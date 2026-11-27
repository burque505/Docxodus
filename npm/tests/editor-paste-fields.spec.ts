import { test, expect, Page } from '@playwright/test';

/**
 * Merge fields survive copy and paste (the editor's own, and Word's clipboard HTML), and
 * convertMergeText makes «Name» text (pasted before they did) merge fields.
 */

const blocks = '#ex-editor [data-anchor][contenteditable="true"]';
const merge = '#ex-editor [data-field="MERGEFIELD"]';

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

const codes = (page: Page) => page.locator('#ex-editor [data-field-code]').evaluateAll((els) => els.map((e) => e.getAttribute('data-field-code')));
const blockText = (page: Page, index: number) =>
  page.locator(blocks).nth(index).evaluate((b) => (b.textContent ?? '').replace(/[\s ]+$/, ''));

test('a merge field copied in the editor pastes as the field, with its code', async ({ page }) => {
  await open(page, 'Dear \nCopy: ');
  await caretAtEnd(page, 0);
  await page.evaluate(() => (window as any).__ed.exec('insertField', { code: 'MERGEFIELD Amount \\# $,0.00' }));
  await page.keyboard.type(', hello');
  // copy the whole first paragraph (the editor's own copy handler makes the HTML)
  await page.locator(blocks).first().click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  // (within one paragraph the browser copies: a real Ctrl+C, read back from the clipboard)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.keyboard.press('Control+C');
  const html = await page.evaluate(async () => {
    for (const item of await navigator.clipboard.read())
      if (item.types.includes('text/html')) return await (await item.getType('text/html')).text();
    return '';
  });
  expect(html).toContain('data-field-code');
  await caretAtEnd(page, 1);
  expect(await page.evaluate((html) => (window as any).__ed.pasteHtml(html), html)).toBe(true);
  await expect(page.locator(merge)).toHaveCount(2);
  expect(await codes(page)).toEqual(['MERGEFIELD Amount \\# $,0.00', 'MERGEFIELD Amount \\# $,0.00']);
  await expect.poll(() => blockText(page, 1)).toBe('Copy: Dear «Amount», hello');
  // the paste is one undo step
  await page.evaluate(() => (window as any).__ed.exec('undo', {}));
  await expect(page.locator(merge)).toHaveCount(1);
  await expect.poll(() => blockText(page, 1)).toBe('Copy:');
});

const WORD_HTML = `<html><body><!--StartFragment--><p class=MsoNormal>Dear <!--[if supportFields]><span
style='mso-element:field-begin'></span> MERGEFIELD LeFull <span style='mso-element:field-separator'></span><![endif]--><span
style='mso-no-proof:yes'>«LeFull»</span><!--[if supportFields]><span style='mso-element:field-end'></span><![endif]-->,
rate <!--[if supportFields]><span style='mso-element:field-begin'></span>= <span style='mso-element:field-begin'></span>
MERGEFIELD Rate <span style='mso-element:field-separator'></span><![endif]-->«Rate»<!--[if supportFields]><span
style='mso-element:field-end'></span><![endif]--><!--[if supportFields]> * 100 \\# &quot;0.00%&quot; <span
style='mso-element:field-separator'></span><![endif]-->«= Rate * 100»<!--[if supportFields]><span
style='mso-element:field-end'></span><![endif]-->, page <!--[if supportFields]><span style='mso-element:field-begin'></span>
PAGE <span style='mso-element:field-separator'></span><![endif]-->3<!--[if supportFields]><span
style='mso-element:field-end'></span><![endif]--></p><!--EndFragment--></body></html>`;

test("Word's merge fields paste as fields; its other fields as their text", async ({ page }) => {
  await open(page, 'x');
  await caretAtEnd(page, 0);
  await page.keyboard.press('Backspace');
  expect(await page.evaluate((html) => (window as any).__ed.pasteHtml(html), WORD_HTML)).toBe(true);
  expect(await codes(page)).toEqual(['MERGEFIELD LeFull', '= { MERGEFIELD Rate } * 100 \\# "0.00%"']);
  await expect.poll(() => blockText(page, 0)).toMatch(/^Dear «LeFull», rate «= Rate \* 100», page 3$/);
});

test('convertMergeText makes «Name» text fields, and leaves fields and other text alone', async ({ page }) => {
  await open(page, 'Lender: «LeFull» of «Le City»\n«LeShort», said «nobody');
  await caretAtEnd(page, 1);
  await page.evaluate(() => (window as any).__ed.exec('insertField', { code: 'MERGEFIELD Real' }));
  await expect(page.locator(merge)).toHaveCount(1);
  await page.evaluate(() => (window as any).__ed.exec('convertMergeText', {}));
  expect(await codes(page)).toEqual(['MERGEFIELD LeFull', 'MERGEFIELD "Le City"', 'MERGEFIELD LeShort', 'MERGEFIELD Real']);
  await expect.poll(() => blockText(page, 0)).toBe('Lender: «LeFull» of «Le City»');
  await expect.poll(() => blockText(page, 1)).toBe('«LeShort», said «nobody«Real»');
  // one undo step
  await page.evaluate(() => (window as any).__ed.exec('undo', {}));
  expect(await codes(page)).toEqual(['MERGEFIELD Real']);
});
