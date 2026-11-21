import { test, expect, type Page } from '@playwright/test';

/**
 * Show/Hide ¶ (Word's paragraph marks and tabs: Home > Paragraph, Ctrl+Shift+8) and joining the
 * selected paragraphs into one (text from a PDF, a paragraph a printed line).
 */

const blocks = (page: Page) => page.locator('#mj-host [data-anchor][contenteditable="true"]');
const exec = (page: Page, name: string, args: object = {}) =>
  page.evaluate(([n, a]) => (window as any).__ribbon.exec(n, a), [name, args] as const);
const texts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#mj-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/ /g, ' ').trim()));

async function mount(page: Page, paragraphs: string[], options: object = {}) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((opts) => {
    const host = document.createElement('div');
    host.id = 'mj-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const changes: boolean[] = [];
    (window as any).__marksChanges = changes;
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, { ...opts, onFormattingMarksChange: (on: boolean) => changes.push(on) });
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('marks.docx');
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

/** Select from the start of paragraph `from` to the end of paragraph `to`. */
async function selectParagraphs(page: Page, from: number, to: number) {
  await blocks(page).nth(from).click();
  await page.evaluate(([a, b]) => {
    const all = document.querySelectorAll<HTMLElement>('#mj-host [data-anchor][contenteditable="true"]');
    const range = document.createRange();
    range.setStart(all[a as number], 0);
    range.setEnd(all[b as number], all[b as number].childNodes.length);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }, [from, to] as const);
}

test.describe('Show/Hide ¶', () => {
  test('Ctrl+Shift+8 and the button show a ¶ at each paragraph end, without moving the text', async ({ page }) => {
    await mount(page, ['A first paragraph', 'A second one']);
    const button = page.locator('#mj-host [data-dxr="showmarks"]');
    const geometry = () => page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#mj-host [data-anchor][contenteditable="true"]'))
      .map((b) => { const r = b.getBoundingClientRect(); return [Math.round(r.top), Math.round(r.height)]; }));
    const mark = () => page.evaluate(() => {
      const b = document.querySelector<HTMLElement>('#mj-host [data-anchor][contenteditable="true"]')!;
      return getComputedStyle(b, '::after').content;
    });
    const before = await geometry();
    await blocks(page).first().click();
    await page.keyboard.press('Control+Shift+8');
    await expect(button).toHaveAttribute('aria-pressed', 'true');
    expect(await mark()).toBe('"¶"');
    expect(await geometry()).toEqual(before);
    await button.click();
    await expect(button).toHaveAttribute('aria-pressed', 'false');
    expect(await mark()).not.toBe('"¶"');
    expect(await page.evaluate(() => (window as any).__marksChanges)).toEqual([true, false]);
  });

  test('the option shows them from the start, and the next document keeps them', async ({ page }) => {
    await mount(page, ['Text'], { showFormattingMarks: true });
    await expect(page.locator('#mj-host [data-dxr="showmarks"]')).toHaveAttribute('aria-pressed', 'true');
    expect(await exec(page, 'showMarks', { on: true })).toBe(true);
    await page.evaluate(() => (window as any).__ribbon.openBlank('second.docx'));
    await expect.poll(() => page.evaluate(() => !!document.querySelector('#mj-host .docx-show-marks'))).toBe(true);
  });
});

test('a document opened window by window (openAsync) turns the marks on and off', async ({ page }) => {
  await mount(page, ['Saved paragraph', 'Another']);
  await page.evaluate(async () => {
    const r = (window as any).__ribbon;
    await r.openAsync(r.editor.save(), 'reopened.docx', { windowSize: 1 });
  });
  await blocks(page).first().click();
  const button = page.locator('#mj-host [data-dxr="showmarks"]');
  await page.keyboard.press('Control+Shift+8');
  await expect(button).toHaveAttribute('aria-pressed', 'true');
  await expect.poll(() => page.evaluate(() => !!document.querySelector('#mj-host .docx-show-marks'))).toBe(true);
  await page.keyboard.press('Control+Shift+8');
  await expect(button).toHaveAttribute('aria-pressed', 'false');
  await expect.poll(() => page.evaluate(() => !!document.querySelector('#mj-host .docx-show-marks'))).toBe(false);
});

test.describe('Join paragraphs', () => {
  test('the selected lines become one paragraph: a space where lines meet, none after a hyphen; one undo step', async ({ page }) => {
    await mount(page, ['My imagination takes me where a wall of', 'overloaded shelves lords over their profes-', 'sional life.', 'After']);
    await selectParagraphs(page, 0, 2);
    expect(await exec(page, 'joinParagraphs')).toBe(true);
    await expect.poll(() => texts(page)).toEqual(['My imagination takes me where a wall of overloaded shelves lords over their profes-sional life.', 'After']);
    await exec(page, 'undo');
    await expect.poll(() => texts(page)).toEqual(['My imagination takes me where a wall of', 'overloaded shelves lords over their profes-', 'sional life.', 'After']);
  });

  test('a selection in one paragraph joins nothing', async ({ page }) => {
    await mount(page, ['One line', 'Two']);
    await blocks(page).first().click();
    expect(await exec(page, 'joinParagraphs')).toBe(false);
    expect(await texts(page)).toEqual(['One line', 'Two']);
  });
});

test('a paragraph mark stays at its place under a first-line indent', async ({ page }) => {
  await mount(page, ['Indented', '']);
  await exec(page, 'showFormattingMarks', { on: true }).catch(() => page.evaluate(() => (window as any).__ribbon.editor.setShowFormattingMarks(true)));
  await page.evaluate(() => (window as any).__ribbon.editor.setShowFormattingMarks(true));
  for (const index of [0, 1]) {
    await blocks(page).nth(index).click();
    await page.evaluate(() => (window as any).__ribbon.editor.setParagraphFormat({ firstLineIndent: 1440 }));
  }
  // Where the ¶ is drawn: at the text's end (or the first line's start when empty), not an inch on.
  const gaps = await page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#mj-host [data-anchor][contenteditable="true"]')).map((p) => {
    const after = getComputedStyle(p, '::after');
    const range = document.createRange();
    range.selectNodeContents(p);
    const text = (p.textContent ?? '').replace(/ /g, '').trim();
    const end = text ? range.getBoundingClientRect().right : p.getBoundingClientRect().left + parseFloat(getComputedStyle(p).textIndent);
    return { indent: after.textIndent, end: Math.round(end - p.getBoundingClientRect().left) };
  }));
  for (const g of gaps) expect(g.indent).toBe('0px');
  expect(gaps[1].end).toBe(96); // the empty paragraph's first line starts an inch in
});
