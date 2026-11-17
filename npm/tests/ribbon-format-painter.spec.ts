import { test, expect, type Page } from '@playwright/test';

/**
 * Format Painter (Word's, as AhkWord has it): Ctrl+Shift+C / Ctrl+Shift+V copy and paste
 * formatting; the brush on Home > Clipboard paints once on a click, until Esc on a double-click.
 * A caret copies the paragraph's formatting (alignment, list) with the characters'; a click paints
 * the word under it. Each paint is one undo step.
 */

const blocks = (page: Page) => page.locator('#fp-host [data-anchor][contenteditable="true"]');
const exec = (page: Page, name: string, args: object = {}) =>
  page.evaluate(([n, a]) => (window as any).__ribbon.exec(n, a), [name, args] as const);
const brush = (page: Page) => page.locator('#fp-host [data-dxr="formatpainter"]');
/** Each paragraph's list number as shown ("" when it has none). */
const numbers = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#fp-host [data-anchor][contenteditable="true"]')).map((b) =>
    Array.from(b.querySelectorAll(':scope > [data-list-marker]')).map((m) => m.textContent ?? '').join('').replace(/\s+/g, ' ').trim()));

/** The computed style of `word` in paragraph `index`: weight, color, the paragraph's alignment. */
const look = (page: Page, index: number, word: string) => page.evaluate(([i, w]) => {
  const block = document.querySelectorAll<HTMLElement>('#fp-host [data-anchor][contenteditable="true"]')[i as number];
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if ((n.textContent ?? '').includes(w as string)) {
      const cs = getComputedStyle(n.parentElement!);
      return { bold: Number(cs.fontWeight) >= 600, color: cs.color, align: getComputedStyle(block).textAlign };
    }
  }
  return null;
}, [index, word] as const);

/** Select `word` in paragraph `index` (or put the caret `caretAt` characters into it). */
async function select(page: Page, index: number, word: string, caretAt?: number) {
  await blocks(page).nth(index).click();
  await page.evaluate(([i, w, at]) => {
    const block = document.querySelectorAll<HTMLElement>('#fp-host [data-anchor][contenteditable="true"]')[i as number];
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const k = (n.textContent ?? '').indexOf(w as string);
      if (k < 0) continue;
      const range = document.createRange();
      if (at == null) { range.setStart(n, k); range.setEnd(n, k + (w as string).length); }
      else { range.setStart(n, k + (at as number)); range.collapse(true); }
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
      return;
    }
    throw new Error(`no "${w}" in paragraph ${i}`);
  }, [index, word, caretAt ?? null] as const);
}

/** Click in the middle of `word` in paragraph `index`, as a painting click does. */
async function clickWord(page: Page, index: number, word: string) {
  const box = await page.evaluate(([i, w]) => {
    const block = document.querySelectorAll<HTMLElement>('#fp-host [data-anchor][contenteditable="true"]')[i as number];
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const k = (n.textContent ?? '').indexOf(w as string);
      if (k < 0) continue;
      const range = document.createRange();
      range.setStart(n, k + 1);
      range.setEnd(n, k + 2);
      const r = range.getBoundingClientRect();
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    }
    return null;
  }, [index, word] as const);
  expect(box).not.toBeNull();
  await page.mouse.click(box!.x, box!.y);
}

async function mount(page: Page, paragraphs: string[]) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'fp-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('painter.docx');
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

test.describe('Format Painter', () => {
  test('Ctrl+Shift+C / Ctrl+Shift+V copy the formatting and paint the word at the caret, one undo step', async ({ page }) => {
    await mount(page, ['Alpha beta gamma']);
    await select(page, 0, 'Alpha');
    await exec(page, 'bold', { on: true });
    await exec(page, 'fontColor', { color: '#C00000' });
    await expect.poll(() => look(page, 0, 'Alpha')).toMatchObject({ bold: true, color: 'rgb(192, 0, 0)' });

    await select(page, 0, 'Alpha', 2);
    await page.keyboard.press('Control+Shift+C');
    await select(page, 0, 'gamma', 2);
    await page.keyboard.press('Control+Shift+V');
    await expect.poll(() => look(page, 0, 'gamma')).toMatchObject({ bold: true, color: 'rgb(192, 0, 0)' });
    expect(await look(page, 0, 'beta')).toMatchObject({ bold: false });

    await exec(page, 'undo');
    await expect.poll(() => look(page, 0, 'gamma')).toMatchObject({ bold: false });
    expect(await look(page, 0, 'Alpha')).toMatchObject({ bold: true });
  });

  test('a selection copies the formatting most of it has, and paints only the characters', async ({ page }) => {
    await mount(page, ['Alpha beta', 'Plain words here']);
    await select(page, 0, 'Alpha');
    await exec(page, 'bold', { on: true });
    await exec(page, 'align', { alignment: 'center' });
    // "Alpha be" is mostly bold, and leaves out the paragraph's end: characters only.
    await select(page, 0, 'Alpha', 0);
    for (let i = 0; i < 8; i++) await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Control+Shift+C');
    await select(page, 1, 'words');
    await page.keyboard.press('Control+Shift+V');
    await expect.poll(() => look(page, 1, 'words')).toMatchObject({ bold: true });
    const plain = await look(page, 1, 'Plain');
    expect(plain).toMatchObject({ bold: false });
    expect(plain!.align).not.toBe('center');
  });

  test('the brush paints once on a click: the word, and the paragraph copied with a caret', async ({ page }) => {
    await mount(page, ['Title words', 'Body text here']);
    await select(page, 0, 'Title');
    await exec(page, 'bold', { on: true });
    await exec(page, 'align', { alignment: 'center' });
    await select(page, 0, 'Title', 1);
    await brush(page).click();
    await expect(brush(page)).toHaveAttribute('aria-pressed', 'true');
    await clickWord(page, 1, 'text');
    await expect.poll(() => look(page, 1, 'text')).toMatchObject({ bold: true, align: 'center' });
    expect(await look(page, 1, 'Body')).toMatchObject({ bold: false });
    await expect(brush(page)).toHaveAttribute('aria-pressed', 'false');
    // Off now: another click paints nothing.
    await clickWord(page, 1, 'here');
    await page.waitForTimeout(100);
    expect(await look(page, 1, 'here')).toMatchObject({ bold: false });
  });

  test('a double-click keeps painting until Esc', async ({ page }) => {
    await mount(page, ['Red one', 'two three four']);
    await select(page, 0, 'Red');
    await exec(page, 'fontColor', { color: '#C00000' });
    await select(page, 0, 'Red', 1);
    await brush(page).dblclick();
    await expect(brush(page)).toHaveAttribute('aria-pressed', 'true');
    await clickWord(page, 1, 'two');
    await expect.poll(() => look(page, 1, 'two')).toMatchObject({ color: 'rgb(192, 0, 0)' });
    await clickWord(page, 1, 'four');
    await expect.poll(() => look(page, 1, 'four')).toMatchObject({ color: 'rgb(192, 0, 0)' });
    await expect(brush(page)).toHaveAttribute('aria-pressed', 'true');
    await page.keyboard.press('Escape');
    await expect(brush(page)).toHaveAttribute('aria-pressed', 'false');
    await clickWord(page, 1, 'three');
    await page.waitForTimeout(100);
    expect((await look(page, 1, 'three'))!.color).not.toBe('rgb(192, 0, 0)');
  });

  test('painting from a numbered paragraph makes the paragraph number on in that list', async ({ page }) => {
    await mount(page, ['First', 'Second', 'Between', 'After']);
    await select(page, 0, 'First', 1);
    await exec(page, 'list', { kind: 'decimal' });
    await select(page, 1, 'Second', 1);
    await exec(page, 'list', { kind: 'decimal' });
    await expect.poll(() => numbers(page)).toEqual(['1.', '2.', '', '']);
    await select(page, 1, 'Second', 1);
    await page.keyboard.press('Control+Shift+C');
    await select(page, 3, 'After', 1);
    await page.keyboard.press('Control+Shift+V');
    await expect.poll(() => numbers(page)).toEqual(['1.', '2.', '', '3.']);
    // And from a plain paragraph, a list item leaves its list.
    await select(page, 2, 'Between', 1);
    await page.keyboard.press('Control+Shift+C');
    await select(page, 3, 'After', 1);
    await page.keyboard.press('Control+Shift+V');
    await expect.poll(() => numbers(page)).toEqual(['1.', '2.', '', '']);
  });

  test('exec copyFormatting / pasteFormatting / formatPainter', async ({ page }) => {
    await mount(page, ['Alpha beta']);
    await select(page, 0, 'Alpha');
    await exec(page, 'italic', { on: true });
    await select(page, 0, 'Alpha', 1);
    expect(await exec(page, 'copyFormatting')).toBe(true);
    await select(page, 0, 'beta', 1);
    expect(await exec(page, 'pasteFormatting')).toBe(true);
    expect(await exec(page, 'formatPainter', { mode: 'sticky' })).toBe(true);
    await expect(brush(page)).toHaveAttribute('aria-pressed', 'true');
    expect(await exec(page, 'formatPainter', { mode: 'off' })).toBe(true);
    await expect(brush(page)).toHaveAttribute('aria-pressed', 'false');
  });
});
