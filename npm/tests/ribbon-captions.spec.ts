import { test, expect, type Page } from '@playwright/test';

/**
 * Captions: References > Insert Caption (Figure 1 below a paragraph, Table 1 above a table, the
 * next number shown before OK, captions after it numbered again, New Label / Delete Label kept by
 * the host), Insert Table of Figures (filled at once, with or without label and number), and
 * Cross-reference to a caption's label and number.
 */

const HOST = '#cp-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

async function mount(page: Page, lines: string[], options: Record<string, unknown> = {}) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((opts) => {
    const host = document.createElement('div');
    host.id = 'cp-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    (window as any).__labels = null;
    const r = (window as any).DocxodusEditor.mountRibbon(host, {
      ...opts,
      onCaptionLabelsChange: (labels: string[]) => { (window as any).__labels = labels; },
    });
    r.setExports((window as any).Docxodus);
    r.openBlank('cp.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  }, options);
  await page.locator(blocks).first().click();
  await page.evaluate((ls) => (window as any).__ribbon.exec('pasteText', { text: ls.join('\n') }), lines);
  await page.locator(`${HOST} .dxr-tab[data-tab="references"]`).click();
}

const texts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#cp-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/[‎‏​]/g, '').replace(/\s+/g, ' ').trim()));

const dialog = (page: Page) => page.locator(`${HOST} [data-dxr="captiondialog"]`);

async function caption(page: Page, at: RegExp, text: string) {
  await page.locator(blocks).filter({ hasText: at }).first().click();
  await page.locator(`${HOST} [data-dxr="caption"]`).click();
  await expect(dialog(page)).toBeVisible();
  await dialog(page).locator('[data-cp="text"]').fill(text);
  await dialog(page).locator('button[value="ok"]').click();
  await expect(dialog(page)).toBeHidden();
}

test('Insert Caption puts Figure 1 below the paragraph, its number a field; one undo takes it away', async ({ page }) => {
  await mount(page, ['Picture', 'After']);
  await page.locator(blocks).first().click();
  await page.locator(`${HOST} [data-dxr="caption"]`).click();
  await expect(dialog(page).locator('[data-cp="prefix"]')).toHaveText('Figure 1');
  await expect(dialog(page).locator('[data-cp="position"]')).toHaveValue('below');
  await dialog(page).locator('[data-cp="text"]').fill(': A cat');
  await dialog(page).locator('button[value="ok"]').click();
  await expect.poll(() => texts(page)).toEqual(['Picture', 'Figure 1: A cat', 'After']);
  await expect(page.locator(`${HOST} [data-field="SEQ"]`)).toHaveText('1');
  // (the caret is left at the end of the new caption, so Ctrl+Z works at once)
  expect(await page.evaluate(() => (document.activeElement as HTMLElement | null)?.closest?.('[data-anchor]')?.textContent?.trim() ?? document.activeElement?.tagName)).toContain('Figure 1');
  await page.keyboard.press('Control+z');
  await expect.poll(() => texts(page)).toEqual(['Picture', 'After']);
});

test('A caption before others shows its number and numbers them again', async ({ page }) => {
  await mount(page, ['One', 'Two']);
  await caption(page, /^Two$/, ': Second');
  await page.locator(blocks).first().click();
  await page.locator(`${HOST} [data-dxr="caption"]`).click();
  await expect(dialog(page).locator('[data-cp="prefix"]')).toHaveText('Figure 1');
  await dialog(page).locator('[data-cp="format"]').selectOption('ROMAN');
  await expect(dialog(page).locator('[data-cp="prefix"]')).toHaveText('Figure I');
  await dialog(page).locator('[data-cp="format"]').selectOption('ARABIC');
  await dialog(page).locator('[data-cp="text"]').fill(': First');
  await dialog(page).locator('button[value="ok"]').click();
  await expect.poll(() => texts(page)).toEqual(['One', 'Figure 1: First', 'Two', 'Figure 2: Second']);
});

test('New Label adds a label the host keeps; Delete Label takes it away', async ({ page }) => {
  await mount(page, ['One'], { captionLabels: ['Map'] });
  await page.locator(blocks).first().click();
  await page.locator(`${HOST} [data-dxr="caption"]`).click();
  const labels = dialog(page).locator('[data-cp="label"] option');
  await expect(labels).toHaveText(['Equation', 'Figure', 'Map', 'Table']);
  await expect(dialog(page).locator('[data-cp="deletelabel"]')).toBeDisabled();
  await dialog(page).locator('[data-cp="newlabel"]').click();
  await dialog(page).locator('[data-cp="newname"]').fill('Plate');
  await dialog(page).locator('[data-cp="addlabel"]').click();
  await expect(dialog(page).locator('[data-cp="label"]')).toHaveValue('Plate');
  await expect(dialog(page).locator('[data-cp="prefix"]')).toHaveText('Plate 1');
  expect(await page.evaluate(() => (window as any).__labels)).toEqual(['Map', 'Plate']);
  await dialog(page).locator('[data-cp="deletelabel"]').click();
  expect(await page.evaluate(() => (window as any).__labels)).toEqual(['Map']);
  await dialog(page).locator('[data-cp="label"]').selectOption('Map');
  await dialog(page).locator('[data-cp="exclude"]').check();
  await expect(dialog(page).locator('[data-cp="prefix"]')).toHaveText('1');
  await dialog(page).locator('button[value="ok"]').click();
  await expect.poll(() => texts(page)).toEqual(['One', '1']);
});

test('In a table the caption goes above it, as Table 1', async ({ page }) => {
  await mount(page, ['Intro']);
  await page.locator(blocks).first().click();
  await page.evaluate(() => (window as any).__ribbon.exec('table', { rows: 1, columns: 2 }));
  await page.locator(`${HOST} td [data-anchor]`).first().click();
  await page.locator(`${HOST} [data-dxr="caption"]`).click();
  await expect(dialog(page).locator('[data-cp="position"]')).toHaveValue('above');
  await dialog(page).locator('[data-cp="label"]').selectOption('Table');
  await dialog(page).locator('button[value="ok"]').click();
  // (the block just before the table's first cell)
  await expect.poll(() => page.evaluate(() => {
    const all = Array.from(document.querySelectorAll<HTMLElement>('#cp-host [data-anchor][contenteditable="true"]'));
    const first = all.findIndex((b) => b.closest('table'));
    return first > 0 ? (all[first - 1].textContent ?? '').replace(/[‎‏​]/g, '').trim() : null;
  })).toBe('Table 1');
});

test('Insert Table of Figures lists the captions at once, with or without label and number', async ({ page }) => {
  await mount(page, ['Top', 'Cat', 'Dog']);
  await caption(page, /^Cat$/, ': A cat');
  await caption(page, /^Dog$/, ': A dog');
  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+Home');
  await page.locator(`${HOST} [data-dxr="tof"]`).click();
  const tof = page.locator(`${HOST} [data-dxr="tofdialog"]`);
  await expect(tof.locator('[data-tof="label"]')).toHaveValue('Figure');
  await tof.locator('button[value="ok"]').click();
  await expect.poll(async () => (await texts(page)).slice(0, 2).map((t) => t.replace(/\s/g, ''))).toEqual(['Figure1:Acat1', 'Figure2:Adog1']);
  await page.locator(blocks).last().click();
  await page.keyboard.press('Control+z');
  await expect.poll(async () => (await texts(page))[0]).toBe('Top');
  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+Home');
  await page.locator(`${HOST} [data-dxr="tof"]`).click();
  await tof.locator('[data-tof="include"]').uncheck();
  await tof.locator('button[value="ok"]').click();
  await expect.poll(async () => (await texts(page)).slice(0, 2).map((t) => t.replace(/\s/g, ''))).toEqual(['Acat1', 'Adog1']);
});

test('Cross-reference to a caption: only its label and number', async ({ page }) => {
  await mount(page, ['Cat', 'See ']);
  await caption(page, /^Cat$/, ': A cat');
  const see = page.locator(blocks).filter({ hasText: /^See/ });
  await see.click();
  await page.keyboard.press('End');
  await page.locator(`${HOST} [data-dxr="crossrefcap"]`).click();
  const xr = page.locator(`${HOST} [data-dxr="crossrefdialog"]`);
  await xr.locator('[data-xr="type"]').selectOption('caption:Figure');
  await expect(xr.locator('[data-xr="list"] option')).toHaveText(['Figure 1: A cat']);
  await xr.locator('[data-xr="insert"]').selectOption('labelNumber');
  await xr.locator('[data-xr="insertbtn"]').click();
  await xr.locator('[data-xr="close"]').click();
  await expect.poll(() => texts(page)).toEqual(['Cat', 'Figure 1: A cat', 'See Figure 1']);
  await expect(page.locator(`${HOST} [data-field="REF"]`)).toHaveText('Figure 1');
});
