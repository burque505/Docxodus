import { test, expect, type Page } from '@playwright/test';

/** Ctrl+Home / Ctrl+End (the document's start and end) and Ctrl+PgUp / Ctrl+PgDn (the top of
 *  the previous or next page: page boxes in page view, page and section breaks in continuous). */

const HOST = '#pk-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

async function mount(page: Page, lines: string[], paginated = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'pk-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('pk.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate((ls) => (window as any).__ribbon.exec('pasteText', { text: ls.join('\n') }), lines);
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
}

/** The text of the paragraph the caret is in, and the caret's offset in it. */
const caret = (page: Page) => page.evaluate(() => {
  const sel = getSelection()!;
  const block = (sel.anchorNode instanceof Element ? sel.anchorNode : sel.anchorNode?.parentElement)
    ?.closest('[data-anchor][contenteditable="true"]');
  const range = document.createRange();
  if (block) { range.selectNodeContents(block); range.setEnd(sel.anchorNode!, sel.anchorOffset); }
  return { text: (block?.textContent ?? '').replace(/[\s ]+$/, ''), at: block ? range.toString().length : -1 };
});

/** The page box (1-based) the caret is in. */
const caretPage = (page: Page) => page.evaluate(() => {
  const sel = getSelection()!;
  const el = sel.anchorNode instanceof Element ? sel.anchorNode : sel.anchorNode?.parentElement;
  const boxes = Array.from(document.querySelectorAll('#pk-host .page-box:not([data-section-filler])'));
  return boxes.indexOf(el?.closest('.page-box') as Element) + 1;
});

test('Ctrl+End and Ctrl+Home go to the end and the start of the document', async ({ page }) => {
  await mount(page, ['First line', 'Middle line', 'Last line']);
  await page.locator(blocks).nth(1).click();
  await page.keyboard.press('Control+End');
  expect(await caret(page)).toEqual({ text: 'Last line', at: 9 });
  await page.keyboard.press('Control+Home');
  expect(await caret(page)).toEqual({ text: 'First line', at: 0 });
  await page.keyboard.type('X');
  await expect(page.locator(blocks).first()).toHaveText('XFirst line');
});

test('page view: Ctrl+PgDn and Ctrl+PgUp go to the top of the next and previous page', async ({ page }) => {
  const lines = Array.from({ length: 160 }, (_, i) => `Paragraph ${i + 1} with enough words to take a line of its own on the page.`);
  await mount(page, lines, true);
  const pages = await page.locator(`${HOST} .page-box:not([data-section-filler])`).count();
  expect(pages).toBeGreaterThanOrEqual(3);
  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+PageDown');
  expect(await caretPage(page)).toBe(2);
  expect((await caret(page)).at).toBe(0);
  await page.keyboard.press('Control+PageDown');
  expect(await caretPage(page)).toBe(3);
  await page.keyboard.press('Control+PageUp');
  expect(await caretPage(page)).toBe(2);
  await page.keyboard.press('Control+PageUp');
  expect(await caretPage(page)).toBe(1);
  expect(await caret(page)).toEqual({ text: lines[0], at: 0 });
  // (the top of the page is on screen)
  const shown = await page.evaluate(() => {
    const box = document.querySelectorAll('#pk-host .page-box:not([data-section-filler])')[0].getBoundingClientRect();
    return box.top > -box.height / 2;
  });
  expect(shown).toBe(true);
});

test('continuous view: Ctrl+PgDn goes past the next page break, Ctrl+PgUp back; the last page goes to the end', async ({ page }) => {
  await mount(page, ['Page one', 'More on one', 'Page two', 'Page three']);
  // page breaks before "Page two" and "Page three" (Ctrl+Enter at their start: the break ends a
  // paragraph of its own, and the text goes on on the next page)
  await page.locator(blocks).filter({ hasText: 'Page two' }).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Control+Enter');
  await expect.poll(() => page.locator(`${HOST} [data-docx-page-break]`).count()).toBe(1);
  await page.locator(blocks).filter({ hasText: 'Page three' }).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Control+Enter');
  await expect.poll(() => page.locator(`${HOST} [data-docx-page-break]`).count()).toBe(2);
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Control+PageDown');
  expect((await caret(page)).text).toBe('Page two');
  await page.keyboard.press('Control+PageDown');
  expect((await caret(page)).text).toBe('Page three');
  await page.keyboard.press('Control+PageDown');
  expect(await caret(page)).toEqual({ text: 'Page three', at: 10 });
  await page.keyboard.press('Control+PageUp');
  expect((await caret(page)).text).toBe('Page two');
  await page.keyboard.press('Control+PageUp');
  expect(await caret(page)).toEqual({ text: 'Page one', at: 0 });
});
