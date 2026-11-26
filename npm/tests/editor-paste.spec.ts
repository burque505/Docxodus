import { test, expect, Page } from '@playwright/test';

/**
 * Paste (DocxEditor.pasteContent, the paste event, exec paste/pasteHtml/pasteText): the
 * clipboard's HTML from Word, a browser or the editor's own copy goes through the engine's
 * InsertHtml, plain text through InsertPastedText, each as one undo step.
 */

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/** A blank document in a bare DocxEditor with "Hello world" typed, the caret after "Hello ". */
async function openHello(page: Page) {
  await page.goto('/test-harness.html');
  await waitForDocxodus(page);
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'ex-editor';
    document.body.appendChild(container);
    const ed = D.DocxEditor.openBlank(container, D, {});
    (window as any).__ed = ed;
    container.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.focus();
    ed.exec('insertText', { text: 'Hello world' });
    ed.exec('moveCaret', { to: 'paragraphStart' });
  });
  await page.keyboard.press('ArrowRight', { delay: 0 });
  for (let i = 0; i < 5; i++) await page.keyboard.press('ArrowRight');
}

/** Fire a paste event at the focused block, as Ctrl+V does, with this clipboard. */
const firePaste = (page: Page, data: Record<string, string>) => page.evaluate(async (data) => {
  const dt = new DataTransfer();
  for (const [type, value] of Object.entries(data)) dt.setData(type, value);
  const target = document.activeElement as HTMLElement;
  const event = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
  target.dispatchEvent(event);
  await new Promise((r) => setTimeout(r, 50));
  return event.defaultPrevented;
}, data);

const text = (page: Page) => page.evaluate(() => (window as any).__ed.exec('text') as string);

const listMarkers = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#ex-editor [data-anchor][contenteditable="true"]'))
    .map((b) => b.querySelector(':scope > [data-list-marker]'))
    .filter((m): m is Element => !!m)
    .map((m) => (m.textContent ?? '').trim()));

/** Word's clipboard HTML for "Intro", a two-item numbered list with a bullet under item 1. */
const WORD_HTML = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
<head><meta charset="utf-8"><style><!--
@font-face {font-family:"Cambria Math"; panose-1:2 4 5 3 5 4 6 3 2 4;}
p.MsoNormal {margin:0in; font-size:11.0pt; font-family:"Calibri",sans-serif;}
@page WordSection1 {size:8.5in 11.0in; margin:1.0in 1.0in 1.0in 1.0in;}
@list l0:level1 {mso-level-tab-stop:none; mso-level-number-position:left; text-indent:-.25in;}
--></style></head>
<body lang=EN-US style='tab-interval:.5in'>
<!--StartFragment-->
<p class=MsoNormal>Intro<o:p></o:p></p>
<p class=MsoListParagraphCxSpFirst style='text-indent:-.25in;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><![endif]>First item<o:p></o:p></p>
<p class=MsoListParagraphCxSpMiddle style='margin-left:1.0in;mso-add-space:auto;text-indent:-.25in;mso-list:l1 level2 lfo2'><![if !supportLists]><span style='font-family:"Courier New"'><span style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp; </span></span></span><![endif]>Sub bullet<o:p></o:p></p>
<p class=MsoListParagraphCxSpLast style='text-indent:-.25in;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><![endif]>Second item<o:p></o:p></p>
<!--EndFragment-->
</body></html>`;

test.describe('Paste', () => {
  test('one plain paragraph goes into the text at the caret, one undo step', async ({ page }) => {
    await openHello(page);
    expect(await firePaste(page, { 'text/html': '<meta charset="utf-8"><span style="color: red; font-size: 30px">big </span>', 'text/plain': 'big ' }))
      .toBe(true);
    await expect.poll(() => text(page)).toBe('Hello big world');
    await page.keyboard.type('X');
    await page.evaluate(() => (window as any).__ed.exec('undo'));   // the typed X
    await page.evaluate(() => (window as any).__ed.exec('undo'));   // the paste
    await expect.poll(() => text(page)).toBe('Hello world');
  });

  test('Word lists arrive as real lists, and Word markup is left out', async ({ page }) => {
    await openHello(page);
    await firePaste(page, { 'text/html': WORD_HTML, 'text/plain': 'Intro\n1. First item\no Sub bullet\n2. Second item' });
    await expect.poll(() => text(page)).toBe('Hello Intro\nFirst item\nSub bullet\nSecond item\nworld');
    const markers = await listMarkers(page);
    expect(markers.length).toBe(3);
    expect(markers[0]).toMatch(/^1\./);
    expect(markers[2]).toMatch(/^2\./);
  });

  test('several paragraphs split the paragraph; the caret ends after the last', async ({ page }) => {
    await openHello(page);
    await firePaste(page, { 'text/html': '<h2>Head</h2><p>one</p><p>two</p>', 'text/plain': 'Head\none\ntwo' });
    await expect.poll(() => text(page)).toBe('Hello \nHead\none\ntwoworld');
    await page.keyboard.type('!');
    await page.evaluate(() => (document.activeElement as HTMLElement).blur());
    await expect.poll(() => text(page)).toBe('Hello \nHead\none\ntwo!world');
    expect(await page.evaluate(() => !!document.querySelector('#ex-editor h2'))).toBe(true);
  });

  test('Ctrl+Alt+Shift+V and exec pasteText paste the text only', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await openHello(page);
    await page.evaluate(() => navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob(['<h1>Big</h1>'], { type: 'text/html' }),
      'text/plain': new Blob(['plain'], { type: 'text/plain' }),
    })]));
    await page.keyboard.press('Control+Alt+Shift+V');
    await expect.poll(() => text(page)).toBe('Hello plainworld');
    expect(await page.evaluate(() => !!document.querySelector('#ex-editor h1'))).toBe(false);
    expect(await page.evaluate(() => (window as any).__ed.exec('pasteText', { text: 'a\tb\nc' }))).toBe(true);
    await expect.poll(() => text(page)).toBe('Hello plaina\tb\ncworld');
  });

  test('exec pasteHtml takes tables and links', async ({ page }) => {
    await openHello(page);
    const ok = await page.evaluate(() => (window as any).__ed.exec('pasteHtml', {
      html: '<p>See <a href="https://example.com/">this</a></p><table border="1"><tr><td>a</td><td>b</td></tr></table>',
    }));
    expect(ok).toBe(true);
    expect(await page.evaluate(() => document.querySelectorAll('#ex-editor table td').length)).toBe(2);
    expect(await page.evaluate(() => document.querySelector<HTMLAnchorElement>('#ex-editor a[href^="https://example.com"]') !== null)).toBe(true);
  });

  test('the editor\'s own copy of a list pastes back as a list', async ({ page }) => {
    await openHello(page);
    await page.evaluate(() => (window as any).__ed.exec('pasteHtml', { html: '<ol><li>alpha</li><li>beta</li></ol>' }));
    await expect.poll(() => listMarkers(page)).toHaveLength(2);
    // Copy both list items as the browser does, then paste them at the end of the document.
    const copied = await page.evaluate(() => {
      const items = Array.from(document.querySelectorAll<HTMLElement>('#ex-editor [data-anchor][contenteditable="true"]'))
        .filter((b) => b.querySelector('[data-list-marker]'));
      const range = document.createRange();
      range.setStartBefore(items[0]);
      range.setEndAfter(items[items.length - 1]);
      const div = document.createElement('div');
      div.append(range.cloneContents());
      return div.innerHTML;
    });
    await page.evaluate(() => (window as any).__ed.exec('moveCaret', { to: 'documentEnd' }));
    await page.evaluate((html) => (window as any).__ed.exec('pasteHtml', { html }), copied);
    await expect.poll(() => listMarkers(page)).toHaveLength(4);
    expect(await text(page)).toContain('alpha\nbeta');
  });

  test('selectedHtml keeps the selection with its formatting, and pasteHtml takes it back', async ({ page }) => {
    await openHello(page);
    expect(await page.evaluate(() => (window as any).__ed.exec('selectedHtml'))).toBe('');
    // Bold "world", then select "Hello world" and keep it.
    await page.evaluate(() => {
      const ed = (window as any).__ed;
      ed.exec('find', { text: 'world', matchCase: true });
      ed.exec('bold');
      ed.exec('fontFamily', { name: 'Courier New' });
      ed.exec('find', { text: 'Hello world', matchCase: true });
    });
    const kept = await page.evaluate(() => (window as any).__ed.selectionHtml());
    expect(kept.text).toBe('Hello world');
    expect(kept.html).toMatch(/font-weight:\s*(700|bold)/);
    // (the editor's own text color is not the text's: automatic text keeps no color)
    expect(kept.html).not.toMatch(/(^|[^-])color:/);
    expect(await page.evaluate(() => (window as any).__ed.exec('selectedHtml'))).toBe(kept.html);
    await page.evaluate(() => (window as any).__ed.exec('moveCaret', { to: 'documentEnd' }));
    await page.evaluate(async (html) => { await (window as any).__ed.exec('pasteHtml', { html, formatting: 'keep' }); }, kept.html);
    await expect.poll(() => text(page)).toBe('Hello worldHello world');
    // Both "world"s are bold, in Courier New (the font list "Courier New", monospace keeps its first family).
    const boldWorlds = () => page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#ex-editor [data-anchor] *'))
      .filter((s) => s.childElementCount === 0 && s.textContent === 'world' && Number(getComputedStyle(s).fontWeight) >= 700
        && getComputedStyle(s).fontFamily.includes('Courier New')).length);
    await expect.poll(boldWorlds).toBe(2);
  });
});
