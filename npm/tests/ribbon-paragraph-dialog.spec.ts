import { test, expect, type Page } from '@playwright/test';

/**
 * The ribbon's Paragraph dialog (Word's): it opens on the selected paragraphs, shows a blank field
 * where they differ, and OK sets only what the user changed — as one `paragraphFormat` command and
 * one undo step.
 */

async function mount(page: Page, options: Record<string, unknown> = {}) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((opts) => {
    const D = (window as any).Docxodus;
    const host = document.createElement('div');
    host.id = 'pd-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const reported: Array<{ label: string; command: unknown }> = [];
    (window as any).__reported = reported;
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {
      ...opts,
      onCommand: (label: string, _ms: number, command?: unknown) => reported.push({ label, command }),
    });
    ribbon.setExports(D);
    ribbon.openBlank('para.docx');
    ribbon.loader.done();
    (window as any).__ribbon = ribbon;
  }, options);
  const block = page.locator('#pd-host [data-anchor][contenteditable="true"]').first();
  await block.click();
}

const dialog = (page: Page) => page.locator('#pd-host dialog[data-dxr="paradialog"]');
const field = (page: Page, name: string) => dialog(page).locator(`[data-pd="${name}"]`);
const formats = (page: Page) => page.evaluate(() => (window as any).__ribbon.exec('paragraphFormats'));

test('one paragraph: OK applies what changed, as one paragraphFormat and one undo step', async ({ page }) => {
  await mount(page);
  await page.keyboard.type('A paragraph to format with the dialog.');
  await page.locator('#pd-host [data-dxr="paralaunch"]').click();
  await expect(dialog(page)).toBeVisible();
  await expect(field(page, 'left')).toHaveValue('0"');
  await expect(field(page, 'special')).toHaveValue('none');
  await expect(field(page, 'widowControl')).toBeChecked();

  await field(page, 'left').fill('1 in');
  await field(page, 'special').selectOption('hanging');
  await expect(field(page, 'by')).toHaveValue('0.5"');
  await field(page, 'line').selectOption('exact');
  await field(page, 'at').fill('18 pt');
  await field(page, 'after').fill('6');
  await dialog(page).locator('[data-pd-tab="breaks"]').click();
  await field(page, 'keepNext').check();
  await dialog(page).locator('button[value="ok"]').click();
  await expect(dialog(page)).toBeHidden();

  const [f] = await formats(page);
  expect(f).toMatchObject({
    leftIndent: 1440, hangingIndent: 720, lineSpacing: 360, lineSpacingRule: 'exact',
    spacingAfter: 120, keepNext: true, widowControl: true, alignment: 'left',
  });
  const commands = await page.evaluate(() => (window as any).__reported.map((r: any) => r.command).filter(Boolean));
  expect(commands.at(-1)).toEqual({
    name: 'paragraphFormat',
    args: { leftIndent: 1440, hangingIndent: 720, spacingAfter: 120, lineSpacing: 360, lineSpacingRule: 'exact', keepNext: true },
  });

  await page.evaluate(() => (window as any).__ribbon.exec('undo'));
  const [back] = await formats(page);
  expect(back).toMatchObject({ leftIndent: 0, firstLineIndent: 0, lineSpacing: 240, spacingAfter: 0, keepNext: false });

  // Cancel and Esc change nothing.
  await page.locator('#pd-host [data-dxr="paralaunch"]').click();
  await field(page, 'left').fill('2 in');
  await page.keyboard.press('Escape');
  await expect(dialog(page)).toBeHidden();
  expect((await formats(page))[0].leftIndent).toBe(0);
});

test('several paragraphs: differing fields are blank and stay as they are', async ({ page }) => {
  await mount(page);
  await page.keyboard.type('First paragraph.');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second paragraph.');
  // The first is indented, the second is not.
  await page.evaluate(() => {
    const r = (window as any).__ribbon;
    r.exec('moveCaret', { to: 'documentStart' });
    r.exec('paragraphFormat', { leftIndent: 720, keepNext: true });
  });
  // Select both paragraphs.
  await page.evaluate(() => {
    const blocks = document.querySelectorAll<HTMLElement>('#pd-host [data-anchor][contenteditable="true"]');
    const range = document.createRange();
    range.setStart(blocks[0].firstChild ?? blocks[0], 0);
    range.setEnd(blocks[1], blocks[1].childNodes.length);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  });
  expect(await formats(page)).toHaveLength(2);
  await page.locator('#pd-host [data-dxr="paralaunch"]').click();
  await expect(field(page, 'left')).toHaveValue('');
  await expect(field(page, 'right')).toHaveValue('0"');
  expect(await field(page, 'keepNext').evaluate((el: HTMLInputElement) => el.indeterminate)).toBe(true);

  await field(page, 'right').fill('2 cm');
  await dialog(page).locator('button[value="ok"]').click();
  await expect(dialog(page)).toBeHidden();
  const after = await formats(page);
  expect(after.map((f: any) => [f.leftIndent, f.rightIndent, f.keepNext])).toEqual([[720, 1134, true], [0, 1134, false]]);

  // Both paragraphs come back with ONE undo.
  await page.evaluate(() => (window as any).__ribbon.exec('undo'));
  await page.evaluate(() => (window as any).__ribbon.exec('moveCaret', { to: 'documentStart' }));
  expect((await formats(page))[0]).toMatchObject({ leftIndent: 720, rightIndent: 0 });
  await page.evaluate(() => (window as any).__ribbon.exec('moveCaret', { to: 'documentEnd' }));
  expect((await formats(page))[0]).toMatchObject({ leftIndent: 0, rightIndent: 0 });
});

test('centimetres, a decimal comma, and a bad entry keeps the dialog open', async ({ page }) => {
  await mount(page, { measurementUnit: 'cm' });
  await page.keyboard.type('Metric paragraph.');
  await page.evaluate(() => (window as any).__ribbon.openParagraphDialog());
  await expect(field(page, 'left')).toHaveValue('0 cm');
  await field(page, 'special').selectOption('firstLine');
  await expect(field(page, 'by')).toHaveValue('1.25 cm');

  await field(page, 'left').fill('abc');
  await dialog(page).locator('button[value="ok"]').click();
  await expect(dialog(page)).toBeVisible();
  await expect(field(page, 'left')).toHaveAttribute('aria-invalid', 'true');
  await expect(dialog(page).locator('[data-pd-msg="length"]')).toBeVisible();

  await field(page, 'left').fill('1,5');
  await field(page, 'left').press('Enter');
  await expect(dialog(page)).toBeHidden();
  expect((await formats(page))[0]).toMatchObject({ leftIndent: 850, firstLineIndent: 709 });
});

test('the dialog moves by its title and opens where it was left', async ({ page }) => {
  await mount(page);
  await page.keyboard.type('A paragraph under the dialog.');
  await page.locator('#pd-host [data-dxr="paralaunch"]').click();
  const before = (await dialog(page).boundingBox())!;
  const title = dialog(page).locator('[data-pd-title]');
  const t = (await title.boundingBox())!;
  await page.mouse.move(t.x + 20, t.y + 8);
  await page.mouse.down();
  await page.mouse.move(t.x + 20 - 150, t.y + 8 + 120, { steps: 5 });
  await page.mouse.up();
  const moved = (await dialog(page).boundingBox())!;
  expect(Math.round(moved.x - before.x)).toBe(-150);
  expect(Math.round(moved.y - before.y)).toBe(120);
  await page.keyboard.press('Escape');
  await page.locator('#pd-host [data-dxr="paralaunch"]').click();
  const again = (await dialog(page).boundingBox())!;
  expect(Math.round(again.x)).toBe(Math.round(moved.x));
  expect(Math.round(again.y)).toBe(Math.round(moved.y));
});

test('in page view, Page break before moves the paragraph to a new page at once', async ({ page }) => {
  await mount(page);
  await page.keyboard.type('First paragraph.');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second paragraph, to start a new page.');
  await page.evaluate(() => (window as any).__ribbon.exec('pageView', { on: true }));
  const pageOf = () => page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#pd-host [data-anchor][contenteditable="true"]'))
    .map((b) => b.closest('[data-page-number], .page-box')?.getAttribute('data-page-number') ?? Array.from(document.querySelectorAll('#pd-host .page-box')).indexOf(b.closest('.page-box')!)));
  await expect.poll(pageOf).toHaveLength(2);
  const [first, second] = await pageOf();
  expect(second).toBe(first);
  await page.locator('#pd-host [data-anchor][contenteditable="true"]').nth(1).click();
  await page.locator('#pd-host [data-dxr="paralaunch"]').click();
  await dialog(page).locator('[data-pd-tab="breaks"]').click();
  await field(page, 'pageBreakBefore').check();
  await dialog(page).locator('button[value="ok"]').click();
  await expect.poll(async () => { const p = await pageOf(); return p[1] !== p[0]; }).toBe(true);
});
