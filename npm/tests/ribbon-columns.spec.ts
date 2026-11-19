import { test, expect, type Page } from '@playwright/test';

/**
 * Word's Layout > Columns: the Columns list (One, Two, Three), the More Columns dialog with its
 * "Apply to", and columns around selected text (continuous section breaks, as Word puts them).
 */

const blocks = (page: Page) => page.locator('#co-host [data-anchor][contenteditable="true"]');
const exec = (page: Page, name: string, args: object = {}) =>
  page.evaluate(([n, a]) => (window as any).__ribbon.exec(n, a), [name, args] as const);
const texts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#co-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/ /g, ' ').replace(/‎/g, '').trim()));
/** Each section wrapper's column count (1 when it has none) and whether it draws a line between. */
const sections = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#co-host [data-section-index]'))
    .map((s) => `${s.dataset.cols ?? '1'}${s.dataset.colSep === 'true' ? '|' : ''}`));

async function mount(page: Page, paragraphs: string[]) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'co-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('columns.docx');
    ribbon.loader.done();
    (window as any).__ribbon = ribbon;
  });
  await blocks(page).first().click();
  for (const [i, p] of paragraphs.entries()) {
    await page.keyboard.type(p);
    if (i < paragraphs.length - 1) await page.keyboard.press('Enter');
  }
  await expect(blocks(page)).toHaveCount(paragraphs.length);
}

/** Select characters `start`..`end` of paragraph `index` with the keyboard (offsets in its text). */
async function selectText(page: Page, index: number, start: number, end: number) {
  const length = (await texts(page))[index].length;
  await blocks(page).nth(index).click();
  await page.keyboard.press('End');
  for (let i = 0; i < length - end; i++) await page.keyboard.press('ArrowLeft');
  for (let i = 0; i < end - start; i++) await page.keyboard.press('Shift+ArrowLeft');
}

const pick = async (page: Page, value: string) => {
  await page.locator('#co-host [data-tab="layout"]').click();
  await page.locator('#co-host [data-dxr="columns"]').selectOption(value);
};

test.describe('Columns', () => {
  test('Columns > Two puts the caret\'s section in two columns, and Undo takes it back', async ({ page }) => {
    await mount(page, ['First', 'Second']);
    await blocks(page).nth(1).click();
    await pick(page, '2');
    await expect.poll(() => sections(page)).toEqual(['2']);
    expect(await texts(page)).toEqual(['First', 'Second']);
    await exec(page, 'undo');
    await expect.poll(() => sections(page)).toEqual(['1']);
  });

  test('with text selected, only the selection goes into columns', async ({ page }) => {
    await mount(page, ['One', 'Two two', 'Three']);
    await selectText(page, 1, 0, 7);
    await pick(page, '3');
    await expect.poll(() => sections(page)).toEqual(['1', '3', '1']);
    expect(await texts(page)).toEqual(['One', 'Two two', 'Three']);
    await exec(page, 'undo');
    await expect.poll(() => sections(page)).toEqual(['1']);
  });

  test('a selection inside a paragraph splits it around the columns', async ({ page }) => {
    await mount(page, ['abcdef']);
    await selectText(page, 0, 2, 4);
    await pick(page, '2');
    await expect.poll(() => texts(page)).toEqual(['ab', 'cd', 'ef']);
    expect(await sections(page)).toEqual(['1', '2', '1']);
  });

  test('More columns: count, spacing and a line between, for the section', async ({ page }) => {
    await mount(page, ['Text']);
    await blocks(page).first().click();
    await pick(page, 'more');
    const dialog = page.locator('#co-host [data-dxr="columnsdialog"]');
    await expect(dialog).toBeVisible();
    await expect(dialog.locator('[data-cd="count"]')).toHaveValue('1');
    await expect(dialog.locator('[data-cd="applyTo"]')).toHaveValue('section');
    await dialog.locator('[data-cd="count"]').fill('3');
    await dialog.locator('[data-cd="space"]').fill('0.25"');
    await dialog.locator('[data-cd="separator"]').check();
    await expect(dialog.locator('[data-cd-preview] span')).toHaveCount(3);
    await dialog.locator('button[value="ok"]').click();
    await expect(dialog).toBeHidden();
    await expect.poll(() => sections(page)).toEqual(['3|']);
    const info = await page.evaluate(() => (window as any).__ribbon.editor.sectionInfo());
    expect([info.columns, info.columnSpaceTwips, info.columnSeparator]).toEqual([3, 360, true]);
  });

  test('More columns with text selected starts on Selected text and keeps the selection', async ({ page }) => {
    await mount(page, ['One', 'Two', 'Three']);
    await selectText(page, 1, 0, 3);
    await page.evaluate(() => (window as any).__ribbon.openColumnsDialog());
    const dialog = page.locator('#co-host [data-dxr="columnsdialog"]');
    await expect(dialog.locator('[data-cd="applyTo"]')).toHaveValue('selection');
    await dialog.locator('[data-cd="count"]').fill('2');
    await dialog.locator('button[value="ok"]').click();
    await expect.poll(() => sections(page)).toEqual(['1', '2', '1']);
  });

  test('a bad count is refused in the dialog; Cancel changes nothing', async ({ page }) => {
    await mount(page, ['Text']);
    await blocks(page).first().click();
    await page.evaluate(() => (window as any).__ribbon.openColumnsDialog());
    const dialog = page.locator('#co-host [data-dxr="columnsdialog"]');
    await dialog.locator('[data-cd="count"]').fill('0');
    await dialog.locator('button[value="ok"]').click();
    await expect(dialog).toBeVisible();
    await expect(dialog.locator('[data-cd-msg="count"]')).toBeVisible();
    await dialog.locator('button[value="cancel"]').click();
    await expect(dialog).toBeHidden();
    expect(await sections(page)).toEqual(['1']);
  });

  test('the columns command: This point forward and Whole document; too many is refused', async ({ page }) => {
    await mount(page, ['One', 'Two']);
    await blocks(page).nth(1).click();
    await page.keyboard.press('Home');
    await exec(page, 'columns', { count: 2, applyTo: 'forward' });
    await expect.poll(() => sections(page)).toEqual(['1', '2']);
    await exec(page, 'columns', { count: 3, applyTo: 'document' });
    await expect.poll(() => sections(page)).toEqual(['3', '3']);
    await expect(exec(page, 'columns', { count: 45 })).rejects.toThrow();
    expect(await sections(page)).toEqual(['3', '3']);
  });

  test('page view lays the section out in columns', async ({ page }) => {
    await mount(page, ['Text in two columns']);
    await blocks(page).first().click();
    await pick(page, '2');
    await exec(page, 'pageView', { on: true });
    await expect(page.locator('#co-host [data-pagination-columns="true"]').first()).toBeVisible();
  });
});
