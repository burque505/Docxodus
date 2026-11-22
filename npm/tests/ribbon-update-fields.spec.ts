import { test, expect, type Page } from '@playwright/test';

/**
 * Updating fields: References > Table of Contents fills the table at once (one undo with the insert);
 * Update Table and F9 build it again from the headings; Ctrl+A, F9 makes cross-references current;
 * right-click > Update Field; page numbers come from the pages, in page view and (laid out off
 * screen) in continuous view.
 */

const HOST = '#uf-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

/** A document of `lines`: "1 Title" is a Heading 1, "2 Title" a Heading 2, anything else body text. */
async function mount(page: Page, lines: string[], paginated = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'uf-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('uf.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate((ls) => (window as any).__ribbon.exec('pasteText', { text: ls.map((l: string) => l.replace(/^\d /, '')).join('\n') }), lines);
  for (let i = 0; i < lines.length; i++) {
    const m = /^(\d) /.exec(lines[i]);
    if (!m) continue;
    await page.locator(blocks).nth(i).click();
    await page.evaluate((lvl) => (window as any).__ribbon.exec('style', { styleId: `Heading${lvl}` }), m[1]);
  }
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
  await page.locator(`${HOST} .dxr-tab[data-tab="insert"]`).click();
}

/** The table of contents' entries: each paragraph linking to a heading's _Toc bookmark, as its text. */
const entries = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#uf-host [data-anchor][contenteditable="true"]'))
    .filter((b) => b.querySelector('a[href^="#_Toc"]'))
    .map((b) => (b.dataset.committedText ?? b.textContent ?? '').replace(/[‎‏]/g, '').replace(/\s+/g, ' ').trim()));

async function insertContentsAtStart(page: Page) {
  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+Home');
  await page.locator(`${HOST} .dxr-tab[data-tab="references"]`).click();
  await page.locator(`${HOST} [data-dxr="toc"]`).click();
}

test('Insert > Table of Contents fills the table at once; one undo takes it all away', async ({ page }) => {
  await mount(page, ['1 Intro', 'Text', '2 Detail', '1 Methods']);
  await insertContentsAtStart(page);
  await expect.poll(() => entries(page)).toEqual(['Intro 1', 'Detail 1', 'Methods 1']);
  // each entry links to its heading
  await expect(page.locator(`${HOST} a[href^="#_Toc"]`)).toHaveCount(3);
  await page.locator(blocks).last().click();
  await page.keyboard.press('Control+z');
  await expect.poll(() => entries(page)).toEqual([]);
});

test('Update Table and F9 build the table again after the headings change', async ({ page }) => {
  await mount(page, ['1 Intro', 'Text', '1 Methods']);
  await insertContentsAtStart(page);
  await expect.poll(() => entries(page)).toEqual(['Intro 1', 'Methods 1']);
  // rename a heading and add one
  const methods = page.locator(blocks).filter({ hasText: /^Methods$/ });
  await methods.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' and data');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Results');
  await page.evaluate(() => (window as any).__ribbon.exec('style', { styleId: 'Heading1' }));
  await page.locator(`${HOST} [data-dxr="updatetoc"]`).click();
  await expect.poll(() => entries(page)).toEqual(['Intro 1', 'Methods and data 1', 'Results 1']);
  await expect(page.locator(`${HOST} [data-dxr="status"]`)).toContainText('Fields updated');
  // F9 with the caret in the table
  await page.locator(blocks).filter({ hasText: /^Intro$/ }).click();
  await page.keyboard.press('End');
  await page.keyboard.type('duction');
  await page.locator(blocks).filter({ hasText: /^Intro 1$|^Intro\s*1$/ }).first().click();
  await page.keyboard.press('F9');
  await expect.poll(() => entries(page)).toEqual(['Introduction 1', 'Methods and data 1', 'Results 1']);
});

test('Ctrl+A, F9 makes a cross-reference current; right-click > Update Field does too', async ({ page }) => {
  await mount(page, ['1 Intro', 'See ']);
  await page.locator(blocks).nth(1).click();
  await page.keyboard.press('End');
  await page.evaluate(() => (window as any).__ribbon.exec('crossReference', { type: 'heading', target: 'Intro', insert: 'text' }));
  const field = page.locator(`${HOST} [data-field="REF"]`);
  await expect(field).toHaveText('Intro');
  await page.locator(blocks).first().click();
  await page.keyboard.press('End');
  await page.keyboard.type('duction');
  await page.keyboard.press('Control+a');
  await page.keyboard.press('F9');
  await expect(field).toHaveText('Introduction');
  // and once more, from the right-click menu on the field
  await page.locator(blocks).first().click();
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  await field.click({ button: 'right' });
  await expect(page.locator(`${HOST} [data-dxr="fieldmenu"]`)).toBeVisible();
  await page.locator(`${HOST} [data-dxr="fieldmenu"] [data-fm="update"]`).click();
  await expect(field).toHaveText('Introduction!');
});

for (const paginated of [true, false]) {
  test(`page numbers come from the pages (${paginated ? 'page view' : 'continuous view, laid out off screen'})`, async ({ page }) => {
    const filler = Array.from({ length: 70 }, (_, i) => `Filler paragraph ${i + 1} with words enough to fill a line of the page.`);
    await mount(page, ['1 Intro', ...filler, '1 Later'], paginated);
    await insertContentsAtStart(page);
    await expect.poll(() => entries(page)).toHaveLength(2);
    const [first, later] = await entries(page);
    expect(first).toBe('Intro 1');
    expect(later).toMatch(/^Later [2-9]$/);
    expect(later).not.toBe('Later 1');
  });
}
