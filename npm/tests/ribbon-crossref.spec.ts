import { test, expect, type Page } from '@playwright/test';

/**
 * Insert > Cross-reference: Word's dialog (reference type, "Insert reference to", Insert as
 * hyperlink, Include above/below, the target list; Insert stays open, Cancel becomes Close), one
 * undo per reference, the field kept whole in the text, gray beside the caret, Ctrl+click to go to
 * the target, and the `crossReference` command.
 */

const HOST = '#xr-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

/** A document of `lines`: "1 Title" is a Heading 1, "2 Title" a Heading 2, "# Text" a numbered
 *  item, anything else body text. */
async function mount(page: Page, lines: string[], paginated = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'xr-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('xr.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate((ls) => {
    (window as any).__ribbon.exec('pasteText', { text: ls.map((l: string) => l.replace(/^(\d|#) /, '')).join('\n') });
  }, lines);
  for (let i = 0; i < lines.length; i++) {
    const m = /^(\d|#) /.exec(lines[i]);
    if (!m) continue;
    await page.locator(blocks).nth(i).click();
    await page.evaluate((lvl) => {
      const r = (window as any).__ribbon;
      if (lvl === '#') r.exec('list', { kind: 'decimal' });
      else r.exec('style', { styleId: `Heading${lvl}` });
    }, m[1]);
  }
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
  await page.locator(`${HOST} .dxr-tab[data-tab="insert"]`).click();
}

const dialog = (page: Page) => page.locator(`${HOST} dialog[data-dxr="crossrefdialog"]`);
const xr = (page: Page, name: string) => dialog(page).locator(`[data-xr="${name}"]`);
const fields = (page: Page) => page.locator(`${HOST} [data-field][data-field-target]`);
const text = (page: Page, index: number) => page.locator(blocks).nth(index).evaluate((b) => (b.textContent ?? '').replace(/[\s ]+$/, ''));

/** The caret at the end of paragraph `index`. */
async function caretAtEnd(page: Page, index: number) {
  await page.locator(blocks).nth(index).click();
  await page.keyboard.press('End');
}

async function openDialog(page: Page) {
  await page.locator(`${HOST} [data-dxr="crossref"]`).click();
  await expect(dialog(page)).toBeVisible();
}

test('a reference to a heading: its text, as a hyperlink, through a hidden _Ref bookmark; one undo', async ({ page }) => {
  await mount(page, ['1 Intro', 'Body text', '2 Details', 'See ']);
  await caretAtEnd(page, 3);
  await openDialog(page);
  await expect(xr(page, 'type')).toHaveValue('heading');
  await expect(xr(page, 'insert')).toHaveValue('text');
  await expect(xr(page, 'list').locator('option')).toHaveText(['Intro', '   Details']);
  await expect(xr(page, 'position')).toBeDisabled();
  await xr(page, 'list').selectOption({ index: 1 });
  await xr(page, 'insertbtn').click();
  // (Insert leaves the dialog open; Cancel is now Close)
  await expect(dialog(page)).toBeVisible();
  await expect(dialog(page).locator('[data-xr-close="close"]')).toBeVisible();
  await xr(page, 'close').click();
  await expect(dialog(page)).toBeHidden();
  await expect(fields(page)).toHaveCount(1);
  await expect(fields(page)).toHaveText('Details');
  await expect(fields(page)).toHaveAttribute('data-field', 'REF');
  await expect(fields(page)).toHaveAttribute('data-field-target', /^_Ref\d{9}$/);
  await expect(fields(page)).toHaveAttribute('data-field-hyperlink', 'true');
  expect(await text(page, 3)).toBe('See Details');
  const hidden = await page.evaluate(() => (window as any).__ribbon.exec('bookmarks', { hidden: true }));
  expect(hidden.map((b: any) => b.text)).toEqual(['Details']);
  // (the bookmark and the field went in as one step)
  await page.keyboard.press('Control+z');
  await expect(fields(page)).toHaveCount(0);
  await expect.poll(() => page.evaluate(() => (window as any).__ribbon.exec('bookmarks', { hidden: true }))).toEqual([]);
});

test('Insert twice: the second reference goes after the first; the _Ref bookmark is reused', async ({ page }) => {
  await mount(page, ['1 Intro', 'Refs: ']);
  await caretAtEnd(page, 1);
  await openDialog(page);
  await xr(page, 'list').selectOption({ index: 0 });
  await xr(page, 'insertbtn').click();
  await xr(page, 'insert').selectOption('page');
  await xr(page, 'insertbtn').click();
  await xr(page, 'close').click();
  await expect(fields(page)).toHaveCount(2);
  await expect(fields(page).nth(1)).toHaveAttribute('data-field', 'PAGEREF');
  expect(await text(page, 1)).toBe('Refs: Intro1');
  const hidden = await page.evaluate(() => (window as any).__ribbon.exec('bookmarks', { hidden: true }));
  expect(hidden).toHaveLength(1);
});

test('the dialog is not modal: click and type in the text, insert at the new caret, undo, all while it stays open', async ({ page }) => {
  await mount(page, ['1 Intro', '1 Methods', 'First: ', 'Second: ']);
  await caretAtEnd(page, 2);
  await openDialog(page);
  expect(await dialog(page).evaluate((d) => d.matches(':modal'))).toBe(false);
  await xr(page, 'list').selectOption({ index: 0 });
  await xr(page, 'insertbtn').click();
  // into the text with the dialog still open: type there, then insert at that caret
  await caretAtEnd(page, 3);
  await page.keyboard.type('see ');
  await expect(dialog(page)).toBeVisible();
  await xr(page, 'list').selectOption({ index: 1 });
  await xr(page, 'insertbtn').click();
  await expect(dialog(page)).toBeVisible();
  expect(await text(page, 2)).toBe('First: Intro');
  expect(await text(page, 3)).toBe('Second: see Methods');
  // a heading added meanwhile shows in the list when it is used next
  await page.locator(blocks).nth(3).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Results');
  await page.evaluate(() => (window as any).__ribbon.exec('style', { styleId: 'Heading1' }));
  await xr(page, 'list').focus();
  await expect(xr(page, 'list').locator('option')).toHaveText(['Intro', 'Methods', 'Results']);
  // the document's own keys still work: Ctrl+Z in the text takes back the edits, then the second reference
  await page.locator(blocks).nth(3).click();
  for (let i = 0; i < 10 && (await fields(page).count()) > 1; i++) {
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(100);
  }
  await expect(fields(page)).toHaveCount(1);
  await expect(dialog(page)).toBeVisible();
  // Escape in the dialog closes it
  await xr(page, 'list').focus();
  await page.keyboard.press('Escape');
  await expect(dialog(page)).toBeHidden();
});

test('a bookmark, above/below; the field stays whole when typing beside it, and is gray there', async ({ page }) => {
  await mount(page, ['Alpha one', 'Beta two', 'Gamma ']);
  await page.locator(blocks).first().click();
  await page.evaluate(() => (window as any).__ribbon.exec('addBookmark', { name: 'Target' }));
  await caretAtEnd(page, 2);
  await openDialog(page);
  await xr(page, 'type').selectOption('bookmark');
  await expect(xr(page, 'list').locator('option')).toHaveText(['Target']);
  await xr(page, 'insert').selectOption('position');
  await xr(page, 'insertbtn').click();
  await xr(page, 'close').click();
  await expect(fields(page)).toHaveText('above');
  // the caret is after the field: it is shaded, and what is typed goes outside it
  await expect(fields(page)).toHaveAttribute('data-field-shaded', '');
  await page.keyboard.type('!');
  await expect(fields(page)).toHaveText('above');
  expect(await text(page, 2)).toBe('Gamma above!');
  await page.locator(blocks).first().click();
  await expect(fields(page)).not.toHaveAttribute('data-field-shaded', '');
});

test('Ctrl+click on a reference inserted as a hyperlink goes to its heading', async ({ page }) => {
  await mount(page, ['1 Intro', 'Body', '1 Methods', 'See ']);
  await caretAtEnd(page, 3);
  await page.evaluate(() => (window as any).__ribbon.exec('crossReference', { type: 'heading', target: 'Methods', insert: 'text', hyperlink: true }));
  await expect(fields(page)).toHaveText('Methods');
  await fields(page).click({ modifiers: ['Control'] });
  await expect.poll(() => page.evaluate(() => window.getSelection()?.toString() ?? '')).toBe('Methods');
});

test('numbered items: the list shows number and text; a number with no context', async ({ page }) => {
  await mount(page, ['# First', '# Second', 'Item ']);
  await caretAtEnd(page, 2);
  await openDialog(page);
  await xr(page, 'type').selectOption('numbered');
  await expect(xr(page, 'insert')).toHaveValue('page');
  await expect(xr(page, 'list').locator('option')).toHaveText(['1. First', '2. Second']);
  await xr(page, 'insert').selectOption('numberNoContext');
  await expect(xr(page, 'position')).toBeEnabled();
  await xr(page, 'list').selectOption({ index: 1 });
  await xr(page, 'insertbtn').click();
  await xr(page, 'close').click();
  await expect(fields(page)).toHaveText('2');
  expect(await text(page, 2)).toBe('Item 2');
});

test('in page view a page number is the target\'s page', async ({ page }) => {
  await mount(page, ['1 Intro', 'Page ref: '], true);
  await caretAtEnd(page, 1);
  await page.evaluate(() => (window as any).__ribbon.exec('crossReference', { type: 'heading', target: 'Intro', insert: 'page' }));
  await expect(fields(page)).toHaveText('1');
  await expect(fields(page)).toHaveAttribute('data-field', 'PAGEREF');
});
