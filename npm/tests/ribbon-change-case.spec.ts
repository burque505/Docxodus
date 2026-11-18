import { test, expect, type Page } from '@playwright/test';

/**
 * Change Case (Word's, with AhkWord's Headline Case): Home > Font > Aa▾ and Shift+F3. Only the
 * letters that change are written, so each keeps its formatting; one undo step; with a caret, the
 * word there.
 */

const blocks = (page: Page) => page.locator('#cc-host [data-anchor][contenteditable="true"]');
const exec = (page: Page, name: string, args: object = {}) =>
  page.evaluate(([n, a]) => (window as any).__ribbon.exec(n, a), [name, args] as const);
const texts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#cc-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/ /g, ' ').trim()));
/** Whether the text `word` in paragraph `index` is shown bold. */
const boldOf = (page: Page, index: number, word: string) => page.evaluate(([i, w]) => {
  const block = document.querySelectorAll<HTMLElement>('#cc-host [data-anchor][contenteditable="true"]')[i as number];
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if ((n.textContent ?? '').includes(w as string)) return Number(getComputedStyle(n.parentElement!).fontWeight) >= 600;
  }
  return null;
}, [index, word] as const);

/** Select `word` in paragraph `index` (or put the caret `caretAt` characters into it). */
async function select(page: Page, index: number, word: string, caretAt?: number) {
  await blocks(page).nth(index).click();
  await page.evaluate(([i, w, at]) => {
    const block = document.querySelectorAll<HTMLElement>('#cc-host [data-anchor][contenteditable="true"]')[i as number];
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

/** Select the whole of paragraph `index`. */
async function selectAll(page: Page, index: number) {
  await blocks(page).nth(index).click();
  await page.evaluate((i) => {
    const block = document.querySelectorAll<HTMLElement>('#cc-host [data-anchor][contenteditable="true"]')[i];
    const range = document.createRange();
    range.selectNodeContents(block);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }, index);
}

async function mount(page: Page, paragraphs: string[]) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'cc-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('case.docx');
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

test.describe('Change Case', () => {
  test('the menu changes the selection, keeps each character\'s formatting, one undo step', async ({ page }) => {
    await mount(page, ['alpha beta gamma']);
    await select(page, 0, 'beta');
    await exec(page, 'bold', { on: true });
    await selectAll(page, 0);
    await page.locator('#cc-host [data-dxr="changecase"]').click();
    await page.locator('#cc-host [data-case="upper"]').click();
    await expect.poll(() => texts(page)).toEqual(['ALPHA BETA GAMMA']);
    expect(await boldOf(page, 0, 'BETA')).toBe(true);
    expect(await boldOf(page, 0, 'ALPHA')).toBe(false);
    await exec(page, 'undo');
    await expect.poll(() => texts(page)).toEqual(['alpha beta gamma']);
  });

  test('every mode', async ({ page }) => {
    await mount(page, ['hello WORLD. this is it']);
    const cases: Array<[string, string]> = [
      ['upper', 'HELLO WORLD. THIS IS IT'],
      ['lower', 'hello world. this is it'],
      ['title', 'Hello World. This Is It'],
      ['toggle', 'hELLO wORLD. tHIS iS iT'],
      ['sentence', 'Hello world. This is it'],
    ];
    for (const [mode, want] of cases) {
      await selectAll(page, 0);
      expect(await exec(page, 'changeCase', { mode })).toBe(true);
      await expect.poll(() => texts(page)).toEqual([want]);
    }
  });

  test('Headline Case: small words, abbreviations, all capitals, across paragraphs', async ({ page }) => {
    await mount(page, ['the deed of trust for acme LLC v. the bank of the west', 'DEED OF TRUST']);
    await blocks(page).nth(0).click();
    await page.evaluate(() => {
      const block = (i: number) => document.querySelectorAll<HTMLElement>('#cc-host [data-anchor][contenteditable="true"]')[i];
      const range = document.createRange();
      range.setStart(block(0), 0);
      range.setEnd(block(1), block(1).childNodes.length);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
    });
    expect(await exec(page, 'changeCase', { mode: 'headline' })).toBe(true);
    await expect.poll(() => texts(page)).toEqual(['The Deed of Trust for Acme LLC v. the Bank of the West', 'Deed of Trust']);
  });

  test('Shift+F3 cycles the word at the caret: lower → Capitalize → UPPER → lower', async ({ page }) => {
    await mount(page, ['one two three']);
    await select(page, 0, 'two', 1);
    await page.keyboard.press('Shift+F3');
    await expect.poll(() => texts(page)).toEqual(['one Two three']);
    await page.keyboard.press('Shift+F3');
    await expect.poll(() => texts(page)).toEqual(['one TWO three']);
    await page.keyboard.press('Shift+F3');
    await expect.poll(() => texts(page)).toEqual(['one two three']);
  });

  test('the host\'s own headline lists', async ({ page }) => {
    await mount(page, ['report on widgets']);
    await page.evaluate(() => {
      const r = (window as any).__ribbon;
      r.setHeadlineCaseLists((window as any).DocxodusEditor.parseHeadlineCaseLists('[small]\non\n[exact]\nWIDGETS'));
    });
    await selectAll(page, 0);
    expect(await exec(page, 'changeCase', { mode: 'headline' })).toBe(true);
    await expect.poll(() => texts(page)).toEqual(['Report on WIDGETS']);
  });

  test('page view: a change across paragraphs and its Undo repaint only those paragraphs', async ({ page }) => {
    await mount(page, ['Bystander first', 'one two', 'three four', 'Bystander last']);
    await exec(page, 'pageView', { on: true });
    // (the typing still in the last paragraph went in before the view changed)
    await expect.poll(async () => (await texts(page)).filter((t) => t.includes('Bystander')).length).toBe(2);
    await expect.poll(() => page.locator('#cc-host .page-box, #cc-host [data-page-number]').count()).toBeGreaterThan(0);
    const mark = () => page.evaluate(() => {
      for (const b of document.querySelectorAll<HTMLElement>('#cc-host [data-anchor][contenteditable="true"]'))
        if ((b.textContent ?? '').includes('Bystander')) (b as any).__bystander = true;
    });
    const marked = () => page.evaluate(() =>
      Array.from(document.querySelectorAll<HTMLElement>('#cc-host [data-anchor][contenteditable="true"]'))
        .filter((b) => (b.textContent ?? '').includes('Bystander') && (b as any).__bystander === true).length);
    const index = async (word: string) => (await texts(page)).findIndex((t) => t.includes(word));
    await mark();
    await blocks(page).nth(await index('one')).click();
    await page.evaluate(([a, b]) => {
      const all = document.querySelectorAll<HTMLElement>('#cc-host [data-anchor][contenteditable="true"]');
      const range = document.createRange();
      range.setStart(all[a as number], 0);
      range.setEnd(all[b as number], all[b as number].childNodes.length);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
    }, [await index('one'), await index('three')] as const);
    expect(await exec(page, 'changeCase', { mode: 'upper' })).toBe(true);
    await expect.poll(async () => (await texts(page)).filter((t) => /ONE TWO|THREE FOUR/.test(t)).length).toBe(2);
    expect(await marked()).toBe(2);
    await exec(page, 'undo');
    await expect.poll(async () => (await texts(page)).filter((t) => /one two|three four/.test(t)).length).toBe(2);
    expect(await marked()).toBe(2);
  });
});
