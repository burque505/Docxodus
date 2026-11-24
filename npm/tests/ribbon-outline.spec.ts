import { test, expect, type Page } from '@playwright/test';

/**
 * Word's Outline View: the View tab's Outline button and the Outlining tab; headings indented by
 * level with Word's symbols; collapse and Show Level hide parts of the document, and the editing
 * keys step over what is hidden (Enter at a collapsed heading's end starts a heading after its
 * part; Backspace and Delete join nothing hidden); promote, demote and move, by buttons and keys.
 */

const HOST = '#ol-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

async function mount(page: Page, paginated = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'ol-host';
    host.style.height = '700px';
    host.style.width = '900px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('ol.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate(() => (window as any).__ribbon.exec('pasteText', { text: 'Alpha\na1 body\na2 body\nBeta\nb1 body' }));
  for (const n of [0, 3]) {
    await page.locator(blocks).nth(n).click();
    await page.evaluate(() => (window as any).__ribbon.exec('style', { styleId: 'Heading1' }));
  }
  if (paginated) await page.evaluate(() => (window as any).__ribbon.exec('pageView', { on: true }));
  await page.locator(`${HOST} .dxr-tab[data-tab="view"]`).click();
  await page.locator(`${HOST} [data-dxr="viewoutline"]`).click();
  await expect(page.locator(`${HOST} [data-outline-sym]`).first()).toBeAttached();
}

/** Each body paragraph: its text, "#n" for a heading of level n, "~" when hidden. */
const outline = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#ol-host [data-anchor][contenteditable="true"]')).map((b) =>
    `${b.hasAttribute('data-outline-hidden') ? '~' : ''}${b.hasAttribute('data-outline-level') ? `#${b.getAttribute('data-outline-level')} ` : ''}${(b.textContent ?? '').trim()}`));

/** The index of the block the caret is in. */
const caretBlock = (page: Page) => page.evaluate(() => {
  const node = getSelection()!.focusNode;
  const el = (node && (node.nodeType === 1 ? node as Element : node.parentElement))?.closest('[data-anchor][contenteditable="true"]');
  return Array.from(document.querySelectorAll('#ol-host [data-anchor][contenteditable="true"]')).indexOf(el as Element);
});

async function at(page: Page, n: number, end: boolean) {
  await page.locator(blocks).nth(n).click();
  await page.keyboard.press(end ? 'End' : 'Home');
}

async function collapseAlpha(page: Page) {
  await at(page, 0, true);
  await page.keyboard.press('Alt+Shift+Minus');
  await expect.poll(() => outline(page)).toEqual(['#1 Alpha', '~a1 body', '~a2 body', '#1 Beta', 'b1 body']);
}

test.describe('Outline View', () => {
  test('the Outline button shows levels and symbols; Close gives back page view', async ({ page }) => {
    await mount(page, true);
    await expect(page.locator(`${HOST} .dxr-tab[data-tab="outlining"]`)).toBeVisible();
    await expect(page.locator(`${HOST} .page-box`)).toHaveCount(0);
    expect(await outline(page)).toEqual(['#1 Alpha', 'a1 body', 'a2 body', '#1 Beta', 'b1 body']);
    const syms = await page.locator(blocks).evaluateAll((els) => els.map((e) => e.getAttribute('data-outline-sym')));
    expect(syms).toEqual(['plus', 'body', 'body', 'plus', 'body']);
    // (the block drag handle would cover the symbols: Outline View has none)
    await expect(page.locator(`${HOST} .docx-block-handle`)).toHaveCount(0);
    await page.locator(`${HOST} button[data-ol="close"]`).click();
    await expect(page.locator(`${HOST} .page-box`).first()).toBeVisible();
    await expect(page.locator(`${HOST} .dxr-tab[data-tab="outlining"]`)).toBeHidden();
    await expect(page.locator(`${HOST} [data-outline-sym]`)).toHaveCount(0);
  });

  test('collapse and expand by keys and by the symbol', async ({ page }) => {
    await mount(page);
    await collapseAlpha(page);
    await page.keyboard.press('Alt+Shift+Equal');
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', 'a1 body', 'a2 body', '#1 Beta', 'b1 body']);
    const box = await page.locator(blocks).nth(3).boundingBox();
    await page.mouse.click(box!.x - 12, box!.y + 8);
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', 'a1 body', 'a2 body', '#1 Beta', '~b1 body']);
  });

  test('Up and Down step over a collapsed part', async ({ page }) => {
    await mount(page);
    await collapseAlpha(page);
    await page.keyboard.press('ArrowDown');
    expect(await caretBlock(page)).toBe(3);
    await page.keyboard.press('ArrowUp');
    expect(await caretBlock(page)).toBe(0);
    // Right at its end and Left at the next one's start step over it too
    await at(page, 0, true);
    await page.keyboard.press('ArrowRight');
    expect(await caretBlock(page)).toBe(3);
    await page.keyboard.press('ArrowLeft');
    expect(await caretBlock(page)).toBe(0);
  });

  test('Backspace and Delete join nothing hidden', async ({ page }) => {
    await mount(page);
    await collapseAlpha(page);
    await at(page, 3, false);
    await page.keyboard.press('Backspace');
    await at(page, 0, true);
    await page.keyboard.press('Delete');
    await page.waitForTimeout(300);
    expect(await outline(page)).toEqual(['#1 Alpha', '~a1 body', '~a2 body', '#1 Beta', 'b1 body']);
  });

  test('Enter at a collapsed heading\'s end starts a heading after its part', async ({ page }) => {
    await mount(page);
    await collapseAlpha(page);
    await page.keyboard.press('Enter');
    await page.keyboard.type('Gamma');
    await page.locator(blocks).nth(5).click();
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', '~a1 body', '~a2 body', '#1 Gamma', '#1 Beta', 'b1 body']);
  });

  test('a selection over a collapsed heading deletes its part', async ({ page }) => {
    await mount(page);
    await collapseAlpha(page);
    await at(page, 0, false);
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Delete');
    await expect.poll(async () => (await outline(page)).join('|')).not.toContain('a1 body');
    expect((await outline(page)).join('|')).toContain('Beta');
  });

  test('promote, demote and the level box', async ({ page }) => {
    await mount(page);
    await at(page, 1, false);
    await page.keyboard.press('Alt+Shift+ArrowLeft');
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', '#1 a1 body', 'a2 body', '#1 Beta', 'b1 body']);
    await expect(page.locator(`${HOST} [data-dxr="ollevel"]`)).toHaveValue('1');
    await page.keyboard.press('Alt+Shift+ArrowRight');
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', '#2 a1 body', 'a2 body', '#1 Beta', 'b1 body']);
    await page.locator(`${HOST} button[data-ol="body"]`).click();
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', 'a1 body', 'a2 body', '#1 Beta', 'b1 body']);
    // Tab at a heading's start demotes it
    await at(page, 3, false);
    await page.keyboard.press('Tab');
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', 'a1 body', 'a2 body', '#2 Beta', 'b1 body']);
    // and one undo step takes it back
    await page.evaluate(() => (window as any).__ribbon.exec('undo'));
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', 'a1 body', 'a2 body', '#1 Beta', 'b1 body']);
  });

  test('Move Down takes a collapsed heading with its part', async ({ page }) => {
    await mount(page);
    // (both collapsed: Move Down passes the next heading shown, with its hidden part)
    await at(page, 0, true);
    await page.keyboard.press('Alt+Shift+Digit1');
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', '~a1 body', '~a2 body', '#1 Beta', '~b1 body']);
    await page.keyboard.press('Alt+Shift+ArrowDown');
    await expect.poll(() => outline(page)).toEqual(['#1 Beta', '~b1 body', '#1 Alpha', '~a1 body', '~a2 body']);
    await page.locator(`${HOST} button[data-ol="up"]`).click();
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', '~a1 body', '~a2 body', '#1 Beta', '~b1 body']);
  });

  test('Show Level hides body text and lower levels; All Levels shows them', async ({ page }) => {
    await mount(page);
    await at(page, 0, true);
    await page.keyboard.press('Alt+Shift+Digit1');
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', '~a1 body', '~a2 body', '#1 Beta', '~b1 body']);
    await expect(page.locator(`${HOST} [data-dxr="olshow"]`)).toHaveValue('1');
    await page.keyboard.press('Alt+Shift+KeyA');
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', 'a1 body', 'a2 body', '#1 Beta', 'b1 body']);
  });

  test('the caret going into hidden text opens its heading', async ({ page }) => {
    await mount(page);
    await collapseAlpha(page);
    await page.evaluate(() => {
      const b = document.querySelectorAll<HTMLElement>('#ol-host [data-anchor][contenteditable="true"]')[2];
      b.focus();
      getSelection()!.collapse(b.firstChild ?? b, 0);
    });
    await expect.poll(() => outline(page)).toEqual(['#1 Alpha', 'a1 body', 'a2 body', '#1 Beta', 'b1 body']);
  });
});
