import { test, expect, type Page } from '@playwright/test';

/**
 * Word's Table Properties in the ribbon: the Table tab's Cell Size boxes and Alignment buttons, and
 * the Table Properties dialog (Table, Row, Column, Cell, Alt Text), whose OK is one undo step.
 */

const HOST = '#tp-host';

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'tp-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('t.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.evaluate(() => (window as any).__ribbon.editor.insertTable(2, 3, { cellContents: ['A', 'B', 'C', 'D', 'E', 'F'] }));
  await cell(page, 'E').click();
  await expect(page.locator(`${HOST} .dxr-tab[data-tab="table"]`)).toBeVisible();
  await page.locator(`${HOST} .dxr-tab[data-tab="table"]`).click();
}

const cell = (page: Page, text: string) =>
  page.locator(`${HOST} td [data-anchor][contenteditable="true"]`, { hasText: new RegExp(`^${text}$`) }).first();
const info = (page: Page) => page.evaluate(() => (window as any).__ribbon.exec('tableInfo'));
const td = (page: Page, name: string) => page.locator(`${HOST} dialog[data-dxr="tabledialog"] [data-td="${name}"]`);

test('the Alignment buttons set the cell\'s vertical alignment, shown on the button and in the page', async ({ page }) => {
  await mount(page);
  await page.locator(`${HOST} .dxr-panel[data-panel="table"] button[data-tt="alignBottom"]`).click();
  await expect.poll(async () => (await info(page))?.verticalAlignment).toBe('bottom');
  await expect(page.locator(`${HOST} button[data-tt="alignBottom"]`)).toHaveClass(/dxr-on/);
  const style = await cell(page, 'E').evaluate((p) => getComputedStyle(p.closest('td')!).verticalAlign);
  expect(style).toBe('bottom');
  await page.locator(`${HOST} button[data-tt="alignTop"]`).click();
  await expect.poll(async () => (await info(page))?.verticalAlignment ?? 'top').toBe('top');
});

test('the Height and Width boxes set the row height and the column width', async ({ page }) => {
  await mount(page);
  const height = page.locator(`${HOST} [data-dxr="tblheight"]`);
  // As a person does: click into the box, then type. The click takes the selection out of the
  // document; the Table tab (and the box) must stay.
  await height.click();
  await page.waitForTimeout(150);
  await expect(page.locator(`${HOST} .dxr-tab[data-tab="table"]`)).toBeVisible();
  await expect(height).toBeVisible();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('0.5"');
  await page.keyboard.press('Enter');
  await expect.poll(async () => (await info(page))?.rowHeightTwips).toBe(720);
  const width = page.locator(`${HOST} [data-dxr="tblwidth"]`);
  await width.click();
  await page.waitForTimeout(150);
  await expect(width).toBeVisible();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('2"');
  await page.keyboard.press('Enter');
  await expect.poll(async () => {
    const i = await info(page);
    return i?.columnWidthsTwips?.[i.columnIndex];
  }).toBe(2880);
});

test('the dialog shows the table, and OK changes width, alignment, row, column, cell and alt text in one undo step', async ({ page }) => {
  await mount(page);
  const before = await info(page);
  await page.locator(`${HOST} [data-dxr="tblprops"]`).click();
  const dialog = page.locator(`${HOST} dialog[data-dxr="tabledialog"]`);
  await expect(dialog).toBeVisible();
  // Off by default: a new table has no preferred width of its own beyond its grid.
  await expect(td(page, 'align-left')).toBeChecked();

  await td(page, 'haswidth').check();
  await td(page, 'widthtype').selectOption('percent');
  await td(page, 'width').fill('80%');
  await td(page, 'align-center').check();
  await dialog.locator('[data-td-tab="row"]').click();
  await expect(dialog.locator('[data-td-which="row"]')).toHaveText('Row 2:', { useInnerText: true });
  await td(page, 'hasheight').check();
  await td(page, 'height').fill('1"');
  await td(page, 'heightrule').selectOption('exact');
  await td(page, 'allowbreak').uncheck();
  await dialog.locator('[data-td-tab="column"]').click();
  await expect(dialog.locator('[data-td-which="column"]')).toHaveText('Column 2:', { useInnerText: true });
  await td(page, 'colwidth').fill('1.5"');
  await dialog.locator('[data-td-tab="cell"]').click();
  await td(page, 'valign-center').check();
  await dialog.locator('[data-td-tab="alt"]').click();
  await td(page, 'alttitle').fill('Letters');
  await td(page, 'altdesc').fill('Six letters in a grid');
  await dialog.locator('button[value="ok"]').click();
  await expect(dialog).toBeHidden();

  await expect.poll(async () => (await info(page))?.alignment).toBe('center');
  const after = await info(page);
  expect(after).toMatchObject({
    widthType: 'percent', width: 80, alignment: 'center', rowHeightTwips: 1440, rowHeightRule: 'exact',
    allowBreakAcrossPages: false, verticalAlignment: 'center', altTitle: 'Letters', altDescription: 'Six letters in a grid',
  });
  expect(after.columnWidthsTwips[1]).toBe(2160);

  // One Ctrl+Z takes the whole dialog back.
  await cell(page, 'E').click();
  await page.keyboard.press('Control+z');
  await expect.poll(async () => (await info(page))?.alignment ?? null).toBe(before.alignment ?? null);
  const undone = await info(page);
  expect(undone).toMatchObject({ rowHeightTwips: null, verticalAlignment: null, altTitle: null, allowBreakAcrossPages: true });
  expect(undone.columnWidthsTwips).toEqual(before.columnWidthsTwips);
});

test('a bad length marks the field and keeps the dialog open; Cancel changes nothing', async ({ page }) => {
  await mount(page);
  const before = JSON.stringify(await info(page));
  await page.locator(`${HOST} [data-dxr="tblprops"]`).click();
  const dialog = page.locator(`${HOST} dialog[data-dxr="tabledialog"]`);
  await td(page, 'indent').fill('lots');
  await dialog.locator('button[value="ok"]').click();
  await expect(dialog).toBeVisible();
  await expect(td(page, 'indent')).toHaveAttribute('aria-invalid', 'true');
  await dialog.locator('button[value="cancel"]').click();
  await expect(dialog).toBeHidden();
  expect(JSON.stringify(await info(page))).toBe(before);
});

test('OK with nothing changed changes nothing, and outside a table the dialog does not open', async ({ page }) => {
  await mount(page);
  const before = JSON.stringify(await info(page));
  await page.locator(`${HOST} [data-dxr="tblprops"]`).click();
  await page.locator(`${HOST} dialog[data-dxr="tabledialog"] button[value="ok"]`).click();
  expect(JSON.stringify(await info(page))).toBe(before);
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]:not(td *)`).first().click();
  expect(await page.evaluate(() => (window as any).__ribbon.openTableDialog())).toBe(false);
});
