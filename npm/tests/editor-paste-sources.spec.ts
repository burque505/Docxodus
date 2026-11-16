import { test, expect, Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Paste from the places people paste from, with what they really put on the clipboard:
 * - Word: its clipboard HTML for samples/verify-mix.docx (fixtures/paste/word-verify-mix.html, as Word 365
 *   wrote it: a table with merged cells and a nested table, a picture, comments, tracked changes);
 * - a web page: Chromium's own copy of a page, through the system clipboard;
 * - this editor: Ctrl+C and Ctrl+V in the document.
 */

const here = dirname(fileURLToPath(import.meta.url));
const WORD_HTML = readFileSync(join(here, 'fixtures/paste/word-verify-mix.html'), 'utf8');

async function openBlank(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'ex-editor';
    document.body.appendChild(container);
    (window as any).__ed = D.DocxEditor.openBlank(container, D, {});
    container.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.focus();
  });
}

/** Fire a paste event at the focused block, as Ctrl+V does, with this clipboard; wait for it to land. */
async function firePaste(page: Page, data: Record<string, string>) {
  const before = await page.evaluate(() => (window as any).__ed.exec('version'));
  await page.evaluate((data) => {
    const dt = new DataTransfer();
    for (const [type, value] of Object.entries(data)) dt.setData(type, value);
    document.activeElement!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, data);
  await expect.poll(() => page.evaluate(() => (window as any).__ed.exec('version'))).not.toBe(before);
}

/** The body's blocks as the editor shows them: top-level ones, and whether each is in a table. */
const blocks = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#ex-editor [data-anchor][contenteditable="true"]')).map((b) => ({
    text: (() => {
      const copy = b.cloneNode(true) as HTMLElement;
      copy.querySelectorAll('[data-list-marker], [aria-hidden="true"]').forEach((m) => m.remove());
      return (copy.textContent ?? '').replace(/[​ ]/g, ' ').replace(/\s+/g, ' ').trim();
    })(),
    tag: b.tagName.toLowerCase(),
    inTable: !!b.closest('table'),
    list: !!b.querySelector(':scope > [data-list-marker]'),
  })));

/** What every paste of Word's verify-mix must show. */
async function expectVerifyMix(page: Page) {
  const all = await blocks(page);
  const body = all.filter((b) => !b.inTable);
  const texts = all.map((b) => b.text).join('\n');
  // Headings and body text
  expect(body.find((b) => b.text === 'Verify mix')?.tag).toBe('h1');
  expect(body.find((b) => b.text === 'Table')?.tag).toBe('h2');
  // The table: one, with a nested one, and every cell's text in a cell
  expect(await page.evaluate(() => document.querySelectorAll('#ex-editor table table').length)).toBe(1);
  for (const cell of ['Item', 'Quantity', 'Notes', 'Apples', '12', 'Merged across two rows on the left', '7',
    'Nested 1.1', 'Nested 2.2', 'After the nested table', 'Total spans two columns', '19']) {
    expect(all.find((b) => b.text === cell)?.inTable, cell).toBe(true);
  }
  expect(await page.evaluate(() => document.querySelector('#ex-editor td[colspan="2"], #ex-editor th[colspan="2"]') !== null)).toBe(true);
  // The picture
  expect(await page.evaluate(() => document.querySelectorAll('#ex-editor img').length)).toBe(1);
  // Comments: the commented text stays, the marks and the comment list go
  expect(texts).toContain('This sentence has a comment on these words and ends here.');
  expect(texts).toContain('A whole paragraph with a second comment.');
  expect(texts).not.toMatch(/\[R\d\]|First comment|Second comment/);
  // Tracked changes read as accepted
  expect(texts).toContain('the meeting is on Tuesday at ten in room 4.');
  expect(texts).not.toContain('Monday');
  // No blank paragraphs but Word's own one after the table (and none from the layout's line breaks)
  expect(body.filter((b) => b.text === '').length).toBeLessThanOrEqual(2);
  expect(body.map((b) => b.text).filter((t) => t !== '')).toEqual([
    'Verify mix', 'A plain paragraph to type into. Add a word here.', 'Table', 'Image', 'The picture above is a gradient.',
    'Comments', 'This sentence has a comment on these words and ends here.', 'A whole paragraph with a second comment.',
    'Tracked changes', 'Tracked changes: the meeting is on Tuesday at ten in room 4.', 'Undo',
    'Type here, then undo and redo.', 'Last paragraph.',
  ]);
}

test.describe('Paste from Word, a web page and the editor', () => {
  test('Word, Merge Formatting (Ctrl+V): every part of a mixed document arrives', async ({ page }) => {
    await openBlank(page);
    await firePaste(page, { 'text/html': WORD_HTML, 'text/plain': 'Verify mix' });
    await expectVerifyMix(page);
    // One undo step takes it all back.
    await page.evaluate(() => (window as any).__ed.exec('undo'));
    await expect.poll(() => page.evaluate(() => (window as any).__ed.exec('text'))).toBe('');
  });

  test('Word, Keep Source Formatting: the same, with Word\'s look', async ({ page }) => {
    await openBlank(page);
    expect(await page.evaluate((html) => (window as any).__ed.exec('pasteHtml', { html, formatting: 'keep' }), WORD_HTML)).toBe(true);
    await expectVerifyMix(page);
  });

  test('a web page copied by the browser: headings, lists, links, a table; the page\'s fonts left behind', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    const source = await context.newPage();
    await source.setContent(`<!doctype html><html><head><style>
      :root { --gap: 8px; }
      body { font-family: Georgia, serif; color: #333; } h2 { color: darkred; } .lead { font-size: 20px; }
      img { border: 1px solid rgb(200, 204, 209); border-image: none 100% / 1 / 0 stretch; max-width: calc(100% - var(--gap));
            --file-upright: 0.75; background-color: rgba(0, 0, 0, 0); margin: clamp(2px, 1vw, 6px); }
      p { margin-block: 0.5em; padding-inline: max(0px, 1px); }
      </style></head><body>
      <figure><img alt="dot" width="20" height="20" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="><figcaption>A caption</figcaption></figure>
      <h2>Web heading</h2>
      <p class="lead">A <b>bold</b> lead with <a href="https://example.com/page">a link</a>.</p>
      <ul><li>Alpha</li><li>Beta<ol><li>Beta one</li><li>Beta two</li></ol></li><li>Gamma</li></ul>
      <table border="1"><tr><th>Name</th><th>Value</th></tr><tr><td>x</td><td>1</td></tr></table>
      <p>Closing words.</p></body></html>`);
    await source.keyboard.press('Control+A');
    await source.keyboard.press('Control+C');
    await openBlank(page);
    await page.bringToFront();
    await page.evaluate(() => (window as any).__ed.exec('insertText', { text: 'Before.' }));
    expect(await page.evaluate(() => (window as any).__ed.exec('paste'))).toBe(true);
    const all = await blocks(page);
    const body = all.filter((b) => !b.inTable);
    expect(body.find((b) => b.text === 'Web heading')?.tag).toBe('h2');
    expect(body.filter((b) => b.list).map((b) => b.text)).toEqual(['Alpha', 'Beta', 'Beta one', 'Beta two', 'Gamma']);
    expect(all.filter((b) => b.inTable).map((b) => b.text)).toEqual(['Name', 'Value', 'x', '1']);
    expect(await page.evaluate(() => document.querySelector('#ex-editor a[href="https://example.com/page"]')?.textContent)).toBe('a link');
    expect(body.filter((b) => b.text === '').length).toBeLessThanOrEqual(2);
    // Today's CSS (custom properties, calc(), border-image) does not cost the picture or the caption
    expect(await page.evaluate(() => document.querySelectorAll('#ex-editor img').length)).toBe(1);
    expect(body.some((b) => b.text === 'A caption')).toBe(true);
    // Merge Formatting: the page's Georgia and its colors are not the document's
    const fonts = await page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#ex-editor [data-anchor] *, #ex-editor [data-anchor]'))
      .map((e) => getComputedStyle(e).fontFamily).filter((f) => /georgia/i.test(f)));
    expect(fonts).toEqual([]);
  });

  test('the editor\'s own copy: Ctrl+C and Ctrl+V keep a list and bold text', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await openBlank(page);
    await page.evaluate(() => (window as any).__ed.exec('pasteHtml', { html: '<p>Plain <b>bold</b> end</p><ol><li>alpha</li><li>beta</li></ol><p>after</p>' }));
    // Select from "Plain" to the end of "beta" and copy.
    await page.evaluate(() => {
      const items = Array.from(document.querySelectorAll<HTMLElement>('#ex-editor [data-anchor][contenteditable="true"]'));
      const first = items.find((b) => b.textContent?.includes('Plain'))!;
      const last = items.find((b) => b.textContent?.includes('beta'))!;
      const range = document.createRange();
      range.setStart(first.firstChild!, 0);
      const walker = document.createTreeWalker(last, NodeFilter.SHOW_TEXT);
      let end: Text | null = null;
      while (walker.nextNode()) if ((walker.currentNode as Text).data.includes('beta')) end = walker.currentNode as Text;
      range.setEnd(end!, end!.data.length);
      const sel = getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
    });
    await page.keyboard.press('Control+C');
    await page.evaluate(() => (window as any).__ed.exec('moveCaret', { to: 'documentEnd' }));
    const before = await page.evaluate(() => (window as any).__ed.exec('version'));
    await page.keyboard.press('Control+V');
    await expect.poll(() => page.evaluate(() => (window as any).__ed.exec('version'))).not.toBe(before);
    const all = await blocks(page);
    expect(all.filter((b) => b.list).map((b) => b.text)).toEqual(['alpha', 'beta', 'alpha', 'beta']);
    expect(await page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#ex-editor [data-anchor] *'))
      .filter((e) => e.childElementCount === 0 && e.textContent === 'bold' && Number(getComputedStyle(e).fontWeight) >= 600).length)).toBe(2);
  });
});
