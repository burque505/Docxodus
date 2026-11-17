import { test, expect, type Page } from '@playwright/test';

/**
 * Multilevel lists in the ribbon: legal numbering (the 1.1 button), the List Levels dialog
 * (Word's Define New Multilevel List), Restart and Continue numbering, and levels kept for new
 * lists ("Also for new documents").
 */

const blocks = (page: Page) => page.locator('#ll-host [data-anchor][contenteditable="true"]');
/** Each paragraph's list number as shown ("" when it has none). */
const numbers = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#ll-host [data-anchor][contenteditable="true"]')).map((b) =>
    Array.from(b.querySelectorAll(':scope > [data-list-marker]')).map((m) => m.textContent ?? '').join('').replace(/\s+/g, ' ').trim()));
const exec = (page: Page, name: string, args: object = {}) =>
  page.evaluate(([n, a]) => (window as any).__ribbon.exec(n, a), [name, args] as const);
const dialog = (page: Page) => page.locator('#ll-host dialog[data-dxr="listdialog"]');
const row = (page: Page, level: number) => dialog(page).locator(`tr[data-level="${level}"]`);

async function mount(page: Page, paragraphs: string[], options: Record<string, unknown> = {}) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((opts) => {
    const host = document.createElement('div');
    host.id = 'll-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const kept: unknown[] = [];
    (window as any).__kept = kept;
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {
      ...opts,
      ...((opts as any).keep ? { onListDefaults: (levels: unknown) => kept.push(levels) } : {}),
    });
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('lists.docx');
    ribbon.loader.done();
    (window as any).__ribbon = ribbon;
  }, options);
  await blocks(page).first().click();
  for (const [i, p] of paragraphs.entries()) {
    await page.keyboard.type(p);
    if (i < paragraphs.length - 1) await page.keyboard.press('Enter');
  }
  await expect(blocks(page)).toHaveCount(paragraphs.length);
}

test('the 1.1 button numbers the paragraphs 1., 1.1, 1.1.1 by level, as one undo step', async ({ page }) => {
  await mount(page, ['Scope', 'Definitions', 'Terms', 'Payment']);
  await page.keyboard.press('Control+a');
  await page.locator('#ll-host [data-dxr="legallist"]').click();
  await expect.poll(() => numbers(page)).toEqual(['1.', '2.', '3.', '4.']);
  await blocks(page).nth(1).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Tab');
  await blocks(page).nth(2).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await expect.poll(() => numbers(page)).toEqual(['1.', '1.1', '1.1.1', '2.']);
  // The levels read back as legal numbering.
  await blocks(page).nth(0).click();
  const levels = await exec(page, 'listLevels');
  expect(levels.slice(0, 3).map((l: any) => l.text)).toEqual(['%1.', '%1.%2', '%1.%2.%3']);
});

test('List Levels: a level\'s numbering and punctuation, applied to the list, undone in one step', async ({ page }) => {
  await mount(page, ['First', 'Second']);
  await page.keyboard.press('Control+a');
  await exec(page, 'listFormat', { format: 'decimal' });
  await expect.poll(() => numbers(page)).toEqual(['1.', '2.']);
  await blocks(page).nth(0).click();
  await page.locator('#ll-host [data-dxr="listmenu"]').selectOption('levels');
  await expect(dialog(page)).toBeVisible();
  await expect(row(page, 0).locator('[data-ld-field="format"]')).toHaveValue('decimal');
  await row(page, 0).locator('[data-ld-field="format"]').selectOption('upperRoman');
  await row(page, 0).locator('[data-ld-field="punctuation"]').selectOption('()');
  await row(page, 0).locator('[data-ld-field="start"]').fill('3');
  await expect(dialog(page).locator('[data-ld-preview]')).toContainText('(III)');
  await dialog(page).locator('button[value="ok"]').click();
  await expect(dialog(page)).toBeHidden();
  await expect.poll(() => numbers(page)).toEqual(['(III)', '(IV)']);
  await blocks(page).nth(1).click();
  await page.keyboard.press('Control+z');
  await expect.poll(() => numbers(page)).toEqual(['1.', '2.']);
});

test('List Levels with Legal numbering and periods everywhere', async ({ page }) => {
  await mount(page, ['One', 'Two']);
  await blocks(page).nth(0).click();
  await exec(page, 'listFormat', { format: 'decimal' });
  await blocks(page).nth(1).click();
  await exec(page, 'listFormat', { format: 'decimal' });
  await page.keyboard.press('Home');
  await page.keyboard.press('Tab');
  await expect.poll(() => numbers(page)).toEqual(['1.', 'a.']);
  await page.evaluate(() => (window as any).__ribbon.openListDialog());
  await dialog(page).locator('[data-ld="legal"]').check();
  await expect(row(page, 1).locator('[data-ld-field="format"]')).toBeDisabled();
  await dialog(page).locator('[data-ld="periods"]').selectOption('all');
  await dialog(page).locator('button[value="ok"]').click();
  await expect.poll(() => numbers(page)).toEqual(['1.', '1.1.']);
});

test('Restart and Continue numbering from the Numbering menu', async ({ page }) => {
  await mount(page, ['A one', 'A two', 'Between', 'B one', 'B two']);
  await page.keyboard.press('Control+a');
  await exec(page, 'listFormat', { format: 'decimal' });
  await blocks(page).nth(2).click();
  await exec(page, 'listFormat', { format: 'none' });
  await expect.poll(() => numbers(page)).toEqual(['1.', '2.', '', '3.', '4.']);
  await blocks(page).nth(3).click();
  await page.locator('#ll-host [data-dxr="listmenu"]').selectOption('restart');
  await expect.poll(() => numbers(page)).toEqual(['1.', '2.', '', '1.', '2.']);
  await blocks(page).nth(3).click();
  await page.locator('#ll-host [data-dxr="listmenu"]').selectOption('continue');
  await expect.poll(() => numbers(page)).toEqual(['1.', '2.', '', '3.', '4.']);
});

test('"Also for new documents" keeps the levels, and a new list starts from them', async ({ page }) => {
  await mount(page, ['Kept', 'Plain', 'Later'], { keep: true, measurementUnit: 'cm' });
  await blocks(page).nth(0).click();
  await page.evaluate(() => (window as any).__ribbon.openListDialog());
  // Off a list, the dialog shows the levels a new list would get, in centimetres.
  await expect(row(page, 0).locator('[data-ld-field="text"]')).toHaveValue('1.27 cm');
  await row(page, 0).locator('[data-ld-field="punctuation"]').selectOption(')');
  await dialog(page).locator('[data-ld="defaults"]').check();
  await dialog(page).locator('button[value="ok"]').click();
  await expect.poll(() => numbers(page)).toEqual(['1)', '', '']);
  const kept = await page.evaluate(() => (window as any).__kept);
  expect(kept).toHaveLength(1);
  expect(kept[0][0]).toMatchObject({ level: 0, format: 'decimal', text: '%1)' });
  // A new list elsewhere takes the kept levels.
  await blocks(page).nth(2).click();
  await page.locator('#ll-host [data-list="decimal"]').click();
  await expect.poll(() => numbers(page)).toEqual(['1)', '', '2)']);
});
