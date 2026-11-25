import { test, expect, type Page } from '@playwright/test';

/**
 * Home > Styles launcher / View > Styles Pane: Word's Styles pane beside the document — a click
 * applies a style, the caret's style is marked, the filter (Recommended / In use / All styles),
 * Show Preview, New Style…, and a row's menu (Update to Match Selection, Modify…, Delete…, Add to /
 * Remove from Style Gallery).
 */

const HOST = '#sp-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;
const pane = `${HOST} [data-dxr="stylespane"]`;
const rows = `${pane} .dxr-styleitem`;
const menuSel = `${HOST} [data-dxr="stylespanemenu"]`;

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'sp-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const events: Array<[string, unknown]> = [];
    (window as any).__paneEvents = events;
    const r = (window as any).DocxodusEditor.mountRibbon(host, {
      onStylesPaneChange: (on: boolean) => events.push(['pane', on]),
      onStylesPaneSettingsChange: (filter: string, preview: boolean) => events.push(['settings', [filter, preview]]),
    });
    r.setExports((window as any).Docxodus);
    r.openBlank('sp.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate(() => {
    const r = (window as any).__ribbon;
    r.exec('pasteText', { text: 'one\ntwo\nthree' });
    // A custom paragraph style (bold, Georgia) and a custom character style.
    r.exec('setStyle', { name: 'My Quote', run: { bold: true, fontFamily: 'Georgia' } });
    r.exec('setStyle', { name: 'Loud', type: 'character', run: { italic: true } });
  });
}

const openPane = async (page: Page) => {
  await page.locator(`${HOST} [data-dxr="stylespanebtn"]`).click();
  await expect(page.locator(pane)).toBeVisible();
};
const row = (page: Page, name: string) => page.locator(rows).filter({ has: page.locator('.dxr-stylename', { hasText: new RegExp(`^${name}$`) }) });
const names = (page: Page) => page.locator(`${rows} .dxr-stylename`).allTextContents();
const caretStyle = (page: Page) => page.evaluate(() => (window as any).__ribbon.editor.styleAtCaret());
const style = (page: Page, id: string) => page.evaluate((i) => (window as any).__ribbon.editor.styles().find((s: any) => s.id === i) ?? null, id);

async function menu(page: Page, name: string, item: string) {
  await row(page, name).click({ button: 'right' });
  await expect(page.locator(menuSel)).toBeVisible();
  await page.locator(`${menuSel} [data-spcmd="${item}"]`).click();
}

test('opens from the Styles launcher, the View box and the Styles menu; close unchecks the box', async ({ page }) => {
  await mount(page);
  await openPane(page);
  await page.locator(`${HOST} .dxr-tab[data-tab="view"]`).click();
  await expect(page.locator(`${HOST} [data-dxr="showstyles"]`)).toBeChecked();
  await page.locator(`${pane} [data-sp="close"]`).click();
  await expect(page.locator(pane)).toBeHidden();
  await expect(page.locator(`${HOST} [data-dxr="showstyles"]`)).not.toBeChecked();
  await page.locator(`${HOST} [data-dxr="showstyles"]`).check();
  await expect(page.locator(pane)).toBeVisible();
  await page.locator(`${HOST} [data-dxr="showstyles"]`).uncheck();
  await page.locator(`${HOST} .dxr-tab[data-tab="home"]`).click();
  await page.locator(`${HOST} [data-dxr="stylemenubtn"]`).click();
  await page.locator(`${HOST} [data-dxr="stylemenu"] [data-stylecmd="pane"]`).click();
  await expect(page.locator(pane)).toBeVisible();
  expect(await page.evaluate(() => (window as any).__paneEvents.filter((e: any) => e[0] === 'pane').map((e: any) => e[1])))
    .toEqual([true, false, true, false, true]);
});

test('a click applies the style to the caret paragraph and marks it; Enter applies too', async ({ page }) => {
  await mount(page);
  await openPane(page);
  await page.locator(blocks).nth(1).click();
  await expect(row(page, 'Normal')).toHaveAttribute('data-current', '');
  await row(page, 'My Quote').click();
  await expect.poll(() => caretStyle(page)).toBe('MyQuote');
  await expect(row(page, 'My Quote')).toHaveAttribute('data-current', '');
  // (the document kept the caret: typing goes on)
  await page.keyboard.type('X');
  await expect(page.locator(blocks).nth(1)).toContainText('X');
  await page.locator(blocks).nth(2).click();
  await row(page, 'Heading 1').focus();
  await page.keyboard.press('Enter');
  await expect.poll(() => page.evaluate(() => {
    const b = document.querySelectorAll('#sp-host [data-anchor][contenteditable="true"]')[2] as HTMLElement;
    return b.tagName;
  })).toBe('H1');
});

test('the filter: In use lists the applied styles only; All styles adds hidden ones; Show Preview', async ({ page }) => {
  await mount(page);
  await openPane(page);
  await page.locator(blocks).nth(0).click();
  await row(page, 'My Quote').click();
  await page.locator(`${pane} [data-sp="filter"]`).selectOption('inuse');
  await expect.poll(() => names(page)).toEqual(expect.arrayContaining(['Normal', 'My Quote']));
  expect(await names(page)).not.toContain('Loud');
  expect(await names(page)).not.toContain('Heading 1');
  await page.locator(`${pane} [data-sp="filter"]`).selectOption('all');
  await expect.poll(async () => (await names(page)).length).toBeGreaterThan(4);
  expect(await names(page)).toContain('Loud');
  await page.locator(`${pane} [data-sp="filter"]`).selectOption('recommended');
  await page.locator(`${pane} [data-sp="preview"]`).check();
  const look = await row(page, 'My Quote').locator('.dxr-stylename').evaluate((el) => ({ weight: (el as HTMLElement).style.fontWeight, family: (el as HTMLElement).style.fontFamily }));
  expect(look.weight).toBe('700');
  expect(look.family).toContain('Georgia');
  await expect(row(page, 'Loud').locator('.dxr-styleglyph')).toHaveText('a');
  await expect(row(page, 'My Quote').locator('.dxr-styleglyph')).toHaveText('¶');
  const settings = await page.evaluate(() => (window as any).__paneEvents.filter((e: any) => e[0] === 'settings').map((e: any) => e[1]));
  expect(settings[settings.length - 1]).toEqual(['recommended', true]);
});

test('Delete asks, removes a custom style (its paragraphs go back to Normal); one undo; built-ins cannot be deleted', async ({ page }) => {
  await mount(page);
  await openPane(page);
  await page.locator(blocks).nth(0).click();
  await row(page, 'My Quote').click();
  await expect.poll(() => caretStyle(page)).toBe('MyQuote');

  await row(page, 'Heading 1').click({ button: 'right' });
  await expect(page.locator(`${menuSel} [data-spcmd="delete"]`)).toBeDisabled();
  await page.keyboard.press('Escape');

  let asked = '';
  page.once('dialog', (d) => { asked = d.message(); void d.dismiss(); });
  await menu(page, 'My Quote', 'delete');
  expect(asked).toContain('My Quote');
  expect(await style(page, 'MyQuote')).not.toBeNull();

  page.once('dialog', (d) => void d.accept());
  await menu(page, 'My Quote', 'delete');
  await expect.poll(() => style(page, 'MyQuote')).toBeNull();
  await expect(row(page, 'My Quote')).toHaveCount(0);
  await page.locator(blocks).nth(0).click();
  await expect.poll(() => caretStyle(page)).not.toBe('MyQuote');

  await page.evaluate(() => (window as any).__ribbon.exec('undo'));
  await expect.poll(() => style(page, 'MyQuote')).not.toBeNull();
  await expect(row(page, 'My Quote')).toHaveCount(1);
});

test("Remove from / Add to Style Gallery sets the style's gallery flag and the menu offers the other", async ({ page }) => {
  await mount(page);
  await openPane(page);
  await menu(page, 'My Quote', 'removegallery');
  await expect.poll(async () => !!(await style(page, 'MyQuote')).quickFormat).toBe(false);
  await row(page, 'My Quote').click({ button: 'right' });
  await expect(page.locator(`${menuSel} [data-spcmd="removegallery"]`)).toBeHidden();
  await page.locator(`${menuSel} [data-spcmd="addgallery"]`).click();
  await expect.poll(async () => !!(await style(page, 'MyQuote')).quickFormat).toBe(true);
});

test('Update to Match Selection names the style and takes the selection\'s formatting; Modify opens the dialog', async ({ page }) => {
  await mount(page);
  await openPane(page);
  await page.locator(blocks).nth(1).click();
  await row(page, 'My Quote').click();
  await page.locator(blocks).nth(1).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await page.keyboard.press('Control+i');
  await row(page, 'My Quote').click({ button: 'right' });
  await expect(page.locator(`${menuSel} [data-spcmd="update"]`)).toContainText('My Quote');
  await page.locator(`${menuSel} [data-spcmd="update"]`).click();
  await expect.poll(async () => (await style(page, 'MyQuote')).resolvedRun?.italic).toBe(true);

  await menu(page, 'My Quote', 'modify');
  await expect(page.locator(`${HOST} [data-dxr="styledialog"]`)).toBeVisible();
  await page.keyboard.press('Escape');
  await page.locator(`${pane} [data-sp="new"]`).click();
  await expect(page.locator(`${HOST} [data-dxr="styledialog"]`)).toBeVisible();
});
