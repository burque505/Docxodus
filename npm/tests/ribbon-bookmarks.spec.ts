import { test, expect, type Page } from '@playwright/test';

/**
 * Insert > Bookmark: Word's Bookmark dialog (Add, Delete, Go To, sort, hidden bookmarks), its
 * Ctrl+Shift+F5, the Link popover's "place in this document", and View > Bookmarks (Word's gray
 * brackets).
 */

const HOST = '#bm-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

async function mount(page: Page, paginated = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'bm-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('bm.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  for (const text of ['Alpha one', 'Beta two', 'Gamma three']) {
    await page.keyboard.type(text);
    if (text !== 'Gamma three') await page.keyboard.press('Enter');
  }
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
  await page.locator(`${HOST} .dxr-tab[data-tab="insert"]`).click();
}

/** Select paragraph `index`'s text (End, then Shift+Home). */
async function selectParagraph(page: Page, index: number) {
  await page.locator(blocks).nth(index).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Shift+Home');
}

const dialog = (page: Page) => page.locator(`${HOST} dialog[data-dxr="bookmarkdialog"]`);
const bm = (page: Page, name: string) => dialog(page).locator(`[data-bm="${name}"]`);
const marks = (page: Page, hidden = false) => page.evaluate((h) => (window as any).__ribbon.exec('bookmarks', { hidden: h }), hidden);
const selected = (page: Page) => page.evaluate(() => window.getSelection()?.toString() ?? '');

async function addViaDialog(page: Page, name: string) {
  await page.locator(`${HOST} [data-dxr="bookmark"]`).click();
  await expect(dialog(page)).toBeVisible();
  await bm(page, 'name').fill(name);
  await bm(page, 'add').click();
  await expect(dialog(page)).toBeHidden();
}

test('Add puts a bookmark on the selection; the dialog lists it; Undo takes it away', async ({ page }) => {
  await mount(page);
  await selectParagraph(page, 1);
  await addViaDialog(page, 'Second');
  expect(await marks(page)).toEqual([{ name: 'Second', text: 'Beta two' }]);
  // the text is still there and still editable
  await expect(page.locator(blocks).nth(1)).toHaveText('Beta two');
  await page.locator(`${HOST} [data-dxr="bookmark"]`).click();
  await expect(bm(page, 'list').locator('option')).toHaveText(['Second']);
  await bm(page, 'list').selectOption('Second');
  await expect(bm(page, 'name')).toHaveValue('Second');
  await dialog(page).locator('button[value="cancel"]').click();
  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+z');
  await expect.poll(() => marks(page)).toEqual([]);
});

test('a bookmark can span paragraphs, and Add with a name the document has moves it', async ({ page }) => {
  await mount(page);
  await page.locator(blocks).first().click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+End');
  await addViaDialog(page, 'Span');
  const [span] = await marks(page);
  expect(span.name).toBe('Span');
  expect(span.text).toContain('Alpha one');
  expect(span.text).toContain('Beta two');
  await selectParagraph(page, 2);
  await addViaDialog(page, 'Span');
  expect(await marks(page)).toEqual([{ name: 'Span', text: 'Gamma three' }]);
});

test('Go To closes the dialog and selects the bookmarked text; sort by name or location', async ({ page }) => {
  await mount(page);
  await selectParagraph(page, 2);
  await addViaDialog(page, 'Alpha');
  await selectParagraph(page, 0);
  await addViaDialog(page, 'Zulu');
  await page.locator(blocks).nth(1).click();
  await page.keyboard.press('Control+Shift+F5');
  await expect(dialog(page)).toBeVisible();
  await expect(bm(page, 'list').locator('option')).toHaveText(['Alpha', 'Zulu']);
  await bm(page, 'sort-location').check();
  await expect(bm(page, 'list').locator('option')).toHaveText(['Zulu', 'Alpha']);
  await expect(bm(page, 'goto')).toBeDisabled();
  await bm(page, 'name').fill('alpha');    // (as Word: the case does not matter)
  await expect(bm(page, 'goto')).toBeEnabled();
  await bm(page, 'goto').click();
  await expect(dialog(page)).toBeHidden();
  await expect.poll(() => selected(page)).toBe('Gamma three');
  await page.keyboard.type('X');
  await expect(page.locator(blocks).nth(2)).toHaveText('X');
});

test('a name Word does not take is refused in the dialog', async ({ page }) => {
  await mount(page);
  await selectParagraph(page, 0);
  await page.locator(`${HOST} [data-dxr="bookmark"]`).click();
  for (const bad of ['1st', 'two words', '_Toc12', 'a-b']) {
    await bm(page, 'name').fill(bad);
    await bm(page, 'add').click();
    await expect(dialog(page)).toBeVisible();
    await expect(dialog(page).locator('[data-bm-msg="name"]')).toBeVisible();
  }
  expect(await marks(page, true)).toEqual([]);
});

test('the Link popover links to a bookmark; Delete of a bookmark in use is refused, saying why', async ({ page }) => {
  await mount(page);
  await selectParagraph(page, 2);
  await addViaDialog(page, 'Target');
  await selectParagraph(page, 0);
  await page.locator(`${HOST} [data-dxr="link"]`).click();
  const place = page.locator(`${HOST} [data-dxr="linkbookmark"]`);
  await expect(place.locator('option')).toHaveCount(2);
  await place.selectOption('Target');
  await expect(page.locator(`${HOST} [data-dxr="linkurl"]`)).toBeDisabled();
  await page.locator(`${HOST} [data-dxr="linkform"] button[type="submit"]`).click();
  await expect(page.locator(`${HOST} a[href="#Target"]`)).toHaveCount(1);
  await page.locator(`${HOST} [data-dxr="bookmark"]`).click();
  await bm(page, 'list').selectOption('Target');
  await bm(page, 'delete').click();
  await expect(dialog(page).locator('[data-bm-msg="inUse"]')).toBeVisible();
  expect((await marks(page)).map((m: { name: string }) => m.name)).toEqual(['Target']);
  await dialog(page).locator('button[value="cancel"]').click();
  // an unused one goes
  await selectParagraph(page, 1);
  await addViaDialog(page, 'Spare');
  await page.locator(`${HOST} [data-dxr="bookmark"]`).click();
  await bm(page, 'list').selectOption('Spare');
  await bm(page, 'delete').click();
  await expect(bm(page, 'list').locator('option')).toHaveText(['Target']);
});

for (const paginated of [false, true]) {
  test(`View > Bookmarks draws brackets at each end, an I-beam for an empty one (${paginated ? 'page' : 'continuous'} view)`, async ({ page }) => {
    await mount(page, paginated);
    await selectParagraph(page, 1);
    await addViaDialog(page, 'Range');
    await page.locator(blocks).nth(2).click();
    await page.keyboard.press('Home');
    await addViaDialog(page, 'Point');
    await page.locator(`${HOST} .dxr-tab[data-tab="view"]`).click();
    await page.locator(`${HOST} [data-dxr="showbookmarks"]`).check();
    const sides = () => page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#bm-host .dxr-bkmark'))
      .map((m) => `${m.dataset.name}:${m.dataset.side}`).sort());
    await expect.poll(sides).toEqual(['Point:empty', 'Range:end', 'Range:start']);
    // the brackets sit at the ends of the bookmarked text
    const fit = await page.evaluate(() => {
      const block = document.querySelectorAll<HTMLElement>('#bm-host [data-anchor][contenteditable="true"]')[1];
      const box = block.getBoundingClientRect();
      const start = document.querySelector<HTMLElement>('#bm-host .dxr-bkmark[data-side="start"]')!.getBoundingClientRect();
      return Math.abs(start.top - box.top) < box.height && Math.abs(start.left - box.left) < 12;
    });
    expect(fit).toBe(true);
    await page.locator(`${HOST} [data-dxr="showbookmarks"]`).uncheck();
    await expect(page.locator(`${HOST} .dxr-bkmarks`)).toBeHidden();
  });
}
