import { test, expect, type Page } from '@playwright/test';

/**
 * Word's Layout > Breaks: page breaks (Ctrl+Enter) and section breaks (Next page, Continuous,
 * Even page, Odd page), the section start of Page Setup, and the continuous view's labelled line.
 */

const blocks = (page: Page) => page.locator('#sb-host [data-anchor][contenteditable="true"]');
const exec = (page: Page, name: string, args: object = {}) =>
  page.evaluate(([n, a]) => (window as any).__ribbon.exec(n, a), [name, args] as const);
const texts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#sb-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/ /g, ' ').replace(/‎/g, '').trim()));
const sections = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#sb-host [data-section-index]'))
    .map((s) => ({ type: s.dataset.sectionType ?? 'nextPage', line: getComputedStyle(s, '::before').content })));

async function mount(page: Page, paragraphs: string[]) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'sb-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('breaks.docx');
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

/** Put the caret `offset` characters into paragraph `index`. */
async function caretAt(page: Page, index: number, offset: number) {
  // From the end: a typed-into blank paragraph can keep a leading placeholder space.
  const length = (await texts(page))[index].length;
  await blocks(page).nth(index).click();
  await page.keyboard.press('End');
  for (let i = 0; i < length - offset; i++) await page.keyboard.press('ArrowLeft');
}

test.describe('Section breaks', () => {
  test('Breaks > Next page splits the paragraph and the continuous view draws a labelled line', async ({ page }) => {
    await mount(page, ['First part', 'Second part']);
    await caretAt(page, 0, 5);
    await page.locator('#sb-host [data-tab="layout"]').click();
    await page.locator('#sb-host [data-dxr="breaks"]').selectOption('nextPage');
    await expect(blocks(page)).toHaveCount(3);
    expect(await texts(page)).toEqual(['First', 'part', 'Second part']);
    const s = await sections(page);
    expect(s.length).toBe(2);
    expect(s[0].line).toBe('none');
    expect(s[1].line).toBe('"Section break (Next page)"');
    const info = await page.evaluate(() => (window as any).__ribbon.editor.sectionInfo());
    expect(info.sectionStart).toBe('nextPage');

    await exec(page, 'undo');
    await expect(blocks(page)).toHaveCount(2);
    expect((await sections(page)).length).toBe(1);
  });

  test('the sectionBreak command and Section start set how the new section starts', async ({ page }) => {
    await mount(page, ['Alpha', 'Beta']);
    await caretAt(page, 1, 0);
    await exec(page, 'sectionBreak', { start: 'continuous' });
    await expect.poll(async () => (await sections(page)).map((s) => s.type)).toEqual(['nextPage', 'continuous']);
    expect((await sections(page))[1].line).toBe('"Section break (Continuous)"');

    // The caret is in the new section: Section start changes that section's start.
    await blocks(page).last().click();
    await page.locator('#sb-host [data-tab="layout"]').click();
    await page.locator('#sb-host [data-dxr="sectionstart"]').selectOption('oddPage');
    await expect.poll(async () => (await sections(page)).map((s) => s.type)).toEqual(['nextPage', 'oddPage']);
    expect((await sections(page))[1].line).toBe('"Section break (Odd page)"');
  });

  test('page view puts the new section on its own page and draws no line', async ({ page }) => {
    await mount(page, ['Alpha', 'Beta']);
    await caretAt(page, 1, 0);
    await exec(page, 'sectionBreak', { start: 'nextPage' });
    await exec(page, 'pageView', { on: true });
    await expect(page.locator('#sb-host .page-box')).toHaveCount(2);
    const lines = await page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#sb-host [data-section-index]'))
      .map((s) => getComputedStyle(s, '::before').content));
    expect(lines.every((l) => l === 'none' || l === 'normal')).toBe(true);
  });

  test('a saved document opened again shows every Next page break, identical sections included', async ({ page }) => {
    await mount(page, ['One', 'Two', 'Three', 'Four']);
    for (const index of [3, 2, 1]) {
      await caretAt(page, index, 0);
      await exec(page, 'sectionBreak', { start: 'nextPage' });
    }
    await expect.poll(async () => (await sections(page)).length).toBe(4);
    await page.evaluate(async () => {
      const ribbon = (window as any).__ribbon;
      await ribbon.openAsync(await ribbon.editor.save(), 'again.docx');
    });
    await expect.poll(async () => (await sections(page)).map((s) => s.line)).toEqual(
      ['none', '"Section break (Next page)"', '"Section break (Next page)"', '"Section break (Next page)"']);
  });

  test('a break is refused in a table and changes nothing', async ({ page }) => {
    await mount(page, ['Before']);
    await exec(page, 'table', { rows: 2, columns: 2 });
    const cell = page.locator('#sb-host td [data-anchor][contenteditable="true"]').first();
    await cell.click();
    const before = await blocks(page).count();
    await expect(exec(page, 'sectionBreak', { start: 'nextPage' })).rejects.toThrow();
    await expect(exec(page, 'insertBreak', { kind: 'page' })).rejects.toThrow();
    expect(await blocks(page).count()).toBe(before);
  });
});

test.describe('Page break', () => {
  test('Ctrl+Enter breaks the page at the caret, and typing keeps the break', async ({ page }) => {
    await mount(page, ['One two']);
    await caretAt(page, 0, 3);
    await page.keyboard.press('Control+Enter');
    await expect(blocks(page)).toHaveCount(2);
    expect(await texts(page)).toEqual(['One', 'two']);
    await exec(page, 'pageView', { on: true });
    await expect(page.locator('#sb-host .page-box')).toHaveCount(2);

    // Typing in the paragraph that ends with the break, then leaving it, keeps the break.
    await blocks(page).first().click();
    await page.keyboard.press('End');
    await page.keyboard.type(' more');
    await blocks(page).last().click();
    await expect.poll(() => texts(page)).toEqual(['One more', 'two']);
    await expect(page.locator('#sb-host .page-box')).toHaveCount(2);

    await exec(page, 'insertBreak', { kind: 'page' });
    await expect(page.locator('#sb-host .page-box')).toHaveCount(3);
  });
});
