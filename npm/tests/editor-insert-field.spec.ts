import { test, expect, Page } from '@playwright/test';

/**
 * DocxEditor.insertField / the `insertField` command: a field from its code, nested fields in
 * braces; a merge field shown «Name», kept whole (the caret goes past it), one undo step;
 * Highlight Merge Fields shades them; mergeFieldNames lists them.
 */

const blocks = '#ex-editor [data-anchor][contenteditable="true"]';

async function open(page: Page, text: string) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(async (text) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'ex-editor';
    document.body.appendChild(container);
    const ed = D.DocxEditor.openBlank(container, D, {});
    (window as any).__ed = ed;
    container.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.focus();
    await ed.exec('pasteText', { text });
  }, text);
}

async function caretAtEnd(page: Page, index: number) {
  await page.locator(blocks).nth(index).click();
  await page.keyboard.press('End');
}

const blockText = (page: Page, index: number) =>
  page.locator(blocks).nth(index).evaluate((b) => (b.textContent ?? '').replace(/[\s ]+$/, ''));

test('a merge field is inserted at the caret, shown «Name» and kept whole', async ({ page }) => {
  await open(page, 'Dear ');
  await caretAtEnd(page, 0);
  await page.evaluate(() => (window as any).__ed.exec('insertField', { code: 'MERGEFIELD First' }));
  const field = page.locator('#ex-editor [data-field="MERGEFIELD"]');
  await expect(field).toHaveCount(1);
  await expect(field).toHaveText('«First»');
  await expect(field).toHaveAttribute('data-field-target', 'First');
  await expect(field).toHaveAttribute('contenteditable', 'false');
  // the caret is after the field: typing goes after it
  await page.keyboard.type(',');
  await expect.poll(() => blockText(page, 0)).toBe('Dear «First»,');
  expect(await page.evaluate(() => (window as any).__ed.mergeFieldNames())).toEqual(['First']);
});

test('the insert is one undo step', async ({ page }) => {
  await open(page, 'Total: ');
  await caretAtEnd(page, 0);
  await page.evaluate(() => (window as any).__ed.exec('insertField', { code: '= { MERGEFIELD Rate } * 100 \\# "0.00%"' }));
  await expect(page.locator('#ex-editor [data-field="="]')).toHaveText('«= Rate * 100»');
  await page.evaluate(() => (window as any).__ed.exec('undo', {}));
  await expect.poll(() => blockText(page, 0)).toBe('Total:');
  await expect(page.locator('#ex-editor [data-field]')).toHaveCount(0);
});

test('an IF with a nested merge field is one marked field', async ({ page }) => {
  await open(page, 'Text');
  await caretAtEnd(page, 0);
  await page.evaluate(() => (window as any).__ed.exec('insertField', { code: 'IF { MERGEFIELD IsGuaranty } = True "Yes" "No"' }));
  const field = page.locator('#ex-editor [data-field]');
  await expect(field).toHaveCount(1);
  await expect(field).toHaveAttribute('data-field', 'IF');
  await expect(field).toHaveText('«IF IsGuaranty = True "Yes" "No"»');
});

test('a selection is replaced; bad code is refused', async ({ page }) => {
  await open(page, 'Hello NAME!');
  await page.locator(blocks).first().click();
  await page.keyboard.press('Home');
  for (let i = 0; i < 6; i++) await page.keyboard.press('ArrowRight');
  for (let i = 0; i < 4; i++) await page.keyboard.press('Shift+ArrowRight');
  await page.evaluate(() => (window as any).__ed.exec('insertField', { code: 'MERGEFIELD Name' }));
  await expect.poll(() => blockText(page, 0)).toBe('Hello «Name»!');
  const refused = await page.evaluate(async () => {
    try { await (window as any).__ed.exec('insertField', { code: '{ MERGEFIELD X' }); return 'ran'; }
    catch (e: any) { return String(e.message ?? e); }
  });
  expect(refused).toContain('invalid_field_code');
});

test('Highlight Merge Fields shades them grey and toggles off', async ({ page }) => {
  await open(page, 'Dear ');
  await caretAtEnd(page, 0);
  await page.evaluate(() => (window as any).__ed.exec('insertField', { code: 'MERGEFIELD First' }));
  await page.locator(blocks).first().click({ position: { x: 2, y: 2 } });   // (caret away from the field)
  const field = page.locator('#ex-editor [data-field="MERGEFIELD"]');
  const bg = () => field.evaluate((el) => getComputedStyle(el).backgroundColor);
  await page.evaluate(() => (window as any).__ed.exec('highlightMergeFields', { on: true }));
  await expect.poll(bg).toBe('rgb(217, 217, 217)');
  await page.evaluate(() => (window as any).__ed.exec('highlightMergeFields', {}));
  await expect.poll(bg).not.toBe('rgb(217, 217, 217)');
});
