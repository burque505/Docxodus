import { test, expect, type Page } from '@playwright/test';

/**
 * Mailings: Insert Merge Field ▾ (the host's columns, then the document's own fields), the Insert
 * Merge Field dialog (formats, \b/\f, code and sample, the last format per field), Highlight Merge
 * Fields, and the host's Start / Select Recipients / Finish buttons (shown only with onMailings).
 */

const HOST = '#mm-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

async function mount(page: Page, withHost: boolean) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((withHost) => {
    const host = document.createElement('div');
    host.id = 'mm-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    (window as any).__mailings = [];
    const r = (window as any).DocxodusEditor.mountRibbon(host, withHost
      ? { onMailings: (c: string) => (window as any).__mailings.push(c), mergeColumns: ['First', 'Amount'] }
      : {});
    r.setExports((window as any).Docxodus);
    r.openBlank('mm.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  }, withHost);
  await page.locator(blocks).first().click();
  await page.evaluate(() => (window as any).__ribbon.editor.exec('pasteText', { text: 'Dear ' }));
  await page.locator(blocks).first().click();
  await page.keyboard.press('End');
  await page.locator(`${HOST} .dxr-tab[data-tab="mailings"]`).click();
}

const field = (page: Page) => page.locator(`${HOST} [data-field="MERGEFIELD"]`);

test('Insert Merge Field ▾ lists the host columns and inserts one', async ({ page }) => {
  await mount(page, true);
  await page.locator(`${HOST} [data-dxr="mergefieldbtn"]`).click();
  const menu = page.locator(`${HOST} [data-dxr="mergefieldmenu"]`);
  await expect(menu.locator('[data-mfname]')).toHaveText(['First', 'Amount']);
  await menu.locator('[data-mfname="First"]').click();
  await expect(field(page)).toHaveText('«First»');
  // the document's own field is listed once, the host's columns first
  await page.evaluate(() => (window as any).__ribbon.setMergeColumns(['Amount']));
  await page.locator(`${HOST} [data-dxr="mergefieldbtn"]`).click();
  await expect(menu.locator('[data-mfname]')).toHaveText(['Amount', 'First']);
});

test('the dialog builds the code, shows a sample and remembers the format', async ({ page }) => {
  await mount(page, true);
  await page.locator(`${HOST} [data-dxr="mergefieldbtn"]`).click();
  await page.locator(`${HOST} [data-mfmenu="dialog"]`).click();
  const dialog = page.locator(`${HOST} [data-dxr="mergefielddialog"]`);
  await expect(dialog).toBeVisible();
  await dialog.locator('select[data-mf="list"]').selectOption('Amount');
  await dialog.locator('select[data-mf="format"]').selectOption('currency');
  await dialog.locator('input[data-mf="sample"]').fill('1234.5');
  await expect(dialog.locator('[data-mf="code"]')).toHaveText('{ MERGEFIELD Amount \\# $,0.00 }');
  await expect(dialog.locator('[data-mf="result"]')).toHaveText('$1,234.50');
  await dialog.locator('select[data-mf="format"]').selectOption('percent');
  await expect(dialog.locator('[data-mf="code"]')).toHaveText('{ = { MERGEFIELD Amount } * 100 \\# "0.00%" }');
  await dialog.locator('select[data-mf="format"]').selectOption('cardtext');
  await dialog.locator('input[data-mf="sample"]').fill('42');
  await expect(dialog.locator('[data-mf="result"]')).toHaveText('forty-two');
  await dialog.locator('select[data-mf="format"]').selectOption('upper');
  await dialog.locator('input[data-mf="before"]').fill('To ');
  await expect(dialog.locator('[data-mf="code"]')).toHaveText('{ MERGEFIELD Amount \\* Upper \\b "To " }');
  await dialog.locator('[data-mf="insert"]').click();
  await expect(field(page)).toHaveCount(1);
  // another field, then back: the last format comes back
  await dialog.locator('select[data-mf="list"]').selectOption('First');
  await expect(dialog.locator('select[data-mf="format"]')).toHaveValue('text');
  await dialog.locator('select[data-mf="list"]').selectOption('Amount');
  await expect(dialog.locator('select[data-mf="format"]')).toHaveValue('upper');
  await expect(dialog.locator('input[data-mf="before"]')).toHaveValue('To ');
});

test('If True / False builds an IF field', async ({ page }) => {
  await mount(page, true);
  await page.evaluate(() => (window as any).__ribbon.openMergeFieldDialog('Member'));
  const dialog = page.locator(`${HOST} [data-dxr="mergefielddialog"]`);
  await dialog.locator('select[data-mf="format"]').selectOption('iftf');
  await dialog.locator('input[data-mf="iftrue"]').fill('Welcome back');
  await dialog.locator('input[data-mf="iffalse"]').fill('Welcome');
  await expect(dialog.locator('input[data-mf="before"]')).toBeHidden();
  await expect(dialog.locator('[data-mf="code"]')).toHaveText('{ IF { MERGEFIELD Member } = True "Welcome back" "Welcome" }');
  await dialog.locator('[data-mf="insert"]').click();
  await expect(page.locator(`${HOST} [data-field="IF"]`)).toHaveCount(1);
});

test('Highlight Merge Fields toggles the shading', async ({ page }) => {
  await mount(page, true);
  const button = page.locator(`${HOST} [data-dxr="highlightmerge"]`);
  await button.click();
  await expect(button).toHaveAttribute('aria-pressed', 'true');
  expect(await page.evaluate(() => (window as any).__ribbon.editor.highlightMergeFields)).toBe(true);
  await button.click();
  expect(await page.evaluate(() => (window as any).__ribbon.editor.highlightMergeFields)).toBe(false);
});

test('the host buttons call onMailings, and are hidden without it', async ({ page }) => {
  await mount(page, true);
  await page.locator(`${HOST} [data-dxr="mailmerge"]`).click();
  await page.locator(`${HOST} [data-dxr="mailrecipients"]`).click();
  await page.locator(`${HOST} [data-dxr="mailfinishbtn"]`).click();
  expect(await page.evaluate(() => (window as any).__mailings)).toEqual(['start', 'recipients', 'finish']);
  await page.evaluate(() => document.getElementById('mm-host')!.remove());
  await mount(page, false);
  await expect(page.locator(`${HOST} [data-dxr="mailstart"]`)).toBeHidden();
  await expect(page.locator(`${HOST} [data-dxr="mailfinish"]`)).toBeHidden();
  await page.locator(`${HOST} [data-dxr="mergefieldbtn"]`).click();
  await expect(page.locator(`${HOST} [data-dxr="mergefieldmenu"] button:disabled`)).toHaveText('(no fields yet)');
});

test('a field goes in right after text typed into a new document', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'mm-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, { mergeColumns: ['First'] });
    r.setExports((window as any).Docxodus);
    r.openBlank('mm.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  // (typed, not pasted: the empty paragraph's placeholder NBSP must not stay in front of the text)
  await page.locator(blocks).first().click();
  await page.keyboard.type('Dear ');
  await page.locator(`${HOST} .dxr-tab[data-tab="mailings"]`).click();
  await page.locator(`${HOST} [data-dxr="mergefieldbtn"]`).click();
  await page.locator(`${HOST} [data-mfname="First"]`).click();
  await expect(field(page)).toHaveText('«First»');
  await page.keyboard.type(',');
  await expect.poll(() => page.locator(blocks).first().evaluate((b) => (b.textContent ?? '').replace(/\s+$/, '')))
    .toBe('Dear «First»,');
});
