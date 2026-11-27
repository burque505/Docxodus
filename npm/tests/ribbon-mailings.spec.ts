import { test, expect, type Page } from '@playwright/test';

/**
 * Mailings: the Merge Fields pane (the host's columns, then the document's own fields; filter;
 * Format with the last format per field; code and sample; Insert; the field at the caret read back
 * and changed), Select Recipients ▾ and the host's Start / Finish (shown only with onMailings),
 * Highlight Merge Fields.
 */

const HOST = '#mm-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;
const COLUMNS = Array.from({ length: 60 }, (_, i) => `Field${String(i + 1).padStart(2, '0')}`).concat(['First', 'Amount']);

async function mount(page: Page, withHost: boolean, type = true) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(({ withHost, columns }) => {
    const host = document.createElement('div');
    host.id = 'mm-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    (window as any).__mailings = [];
    const r = (window as any).DocxodusEditor.mountRibbon(host, withHost
      ? { onMailings: (c: string) => (window as any).__mailings.push(c), mergeColumns: columns, mergeSource: 'people.csv' }
      : {});
    r.setExports((window as any).Docxodus);
    r.openBlank('mm.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  }, { withHost, columns: COLUMNS });
  await page.locator(blocks).first().click();
  // (typed, not pasted: the empty paragraph's placeholder NBSP must not put the field one off)
  if (type) await page.keyboard.type('Dear ');
  await page.locator(`${HOST} .dxr-tab[data-tab="mailings"]`).click();
}

const pane = (page: Page) => page.locator(`${HOST} [data-dxr="mergefieldspane"]`);
const merge = (page: Page) => page.locator(`${HOST} [data-field="MERGEFIELD"]`);
const firstText = (page: Page) => page.locator(blocks).first()
  .evaluate((b) => (b.textContent ?? '').replace(/^[\s ]+/, '').replace(/[\s ]+$/, ''));

async function openPane(page: Page) {
  await page.locator(`${HOST} [data-dxr="mergefieldbtn"]`).click();
  await expect(pane(page)).toBeVisible();
}

test('the pane lists every column, scrolls, filters and inserts on Enter', async ({ page }) => {
  await mount(page, true);
  await openPane(page);
  await expect(pane(page).locator('[data-mf="source"]')).toHaveText('people.csv');
  const items = pane(page).locator('.dxr-mfitem');
  await expect(items).toHaveCount(62);
  // (the list scrolls inside the pane: its last name can be reached)
  await items.last().scrollIntoViewIfNeeded();
  await expect(items.last()).toBeInViewport();
  await pane(page).locator('input[data-mf="filter"]').fill('firs');
  await expect(items).toHaveText(['First']);
  await pane(page).locator('input[data-mf="filter"]').press('Enter');
  await expect(merge(page)).toHaveText('«First»');
  await page.keyboard.type(',');
  await expect.poll(() => firstText(page)).toBe('Dear «First»,');
  // the field is now in the document too: still listed once
  await pane(page).locator('[data-mf="clearfilter"]').click();
  await expect(items).toHaveCount(62);
});

test('a format is chosen in the pane, shown as code and sample, and remembered per field', async ({ page }) => {
  await mount(page, true);
  await openPane(page);
  const p = pane(page);
  await p.locator('input[data-mf="filter"]').fill('Amount');
  await p.locator('select[data-mf="format"]').selectOption('currency');
  await p.locator('input[data-mf="sample"]').fill('1234.5');
  await expect(p.locator('[data-mf="code"]')).toHaveValue('MERGEFIELD Amount \\# $,0.00');
  await expect(p.locator('[data-mf="result"]')).toHaveText('$1,234.50');
  await p.locator('select[data-mf="format"]').selectOption('cardtext');
  await p.locator('input[data-mf="sample"]').fill('42');
  await expect(p.locator('[data-mf="result"]')).toHaveText('forty-two');
  await p.locator('select[data-mf="format"]').selectOption('iftf');
  await expect(p.locator('input[data-mf="before"]')).toBeHidden();
  await p.locator('select[data-mf="format"]').selectOption('upper');
  await p.locator('input[data-mf="before"]').fill('To ');
  await expect(p.locator('[data-mf="code"]')).toHaveValue('MERGEFIELD Amount \\* Upper \\b "To "');
  await p.locator('[data-mf="insert"]').click();
  await expect(merge(page)).toHaveCount(1);
  await p.locator('input[data-mf="filter"]').fill('First');
  await expect(p.locator('select[data-mf="format"]')).toHaveValue('text');
  await p.locator('input[data-mf="filter"]').fill('Amount');
  await expect(p.locator('select[data-mf="format"]')).toHaveValue('upper');
  await expect(p.locator('input[data-mf="before"]')).toHaveValue('To ');
});

test('the field at the caret shows its format, and Apply to Field changes it', async ({ page }) => {
  await mount(page, true);
  await page.evaluate(() => (window as any).__ribbon.editor.exec('insertField', { code: 'MERGEFIELD Amount \\# $,0.00' }));
  await expect(merge(page)).toHaveCount(1);
  await expect(merge(page)).toHaveAttribute('data-field-code', 'MERGEFIELD Amount \\# $,0.00');
  await openPane(page);
  const p = pane(page);
  // (the caret is just after the inserted field)
  await expect(p.locator('[data-mf="at"]')).toBeVisible();
  await expect(p.locator('[data-mf="atcode"]')).toHaveText('{ MERGEFIELD Amount \\# $,0.00 }');
  await expect(p.locator('input[data-mf="name"]')).toHaveValue('Amount');
  await expect(p.locator('select[data-mf="format"]')).toHaveValue('currency');
  await p.locator('select[data-mf="format"]').selectOption('dollartext');
  await p.locator('[data-mf="apply"]').click();
  await expect(merge(page)).toHaveCount(1);
  await expect(merge(page)).toHaveAttribute('data-field-code', 'MERGEFIELD Amount \\* DollarText');
  await expect.poll(() => firstText(page)).toBe('Dear «Amount»');
  // one undo step brings the old format back
  await page.evaluate(() => (window as any).__ribbon.editor.exec('undo', {}));
  await expect(merge(page)).toHaveAttribute('data-field-code', 'MERGEFIELD Amount \\# $,0.00');
});

test('a code the pane has no format for is shown as it is', async ({ page }) => {
  await mount(page, true);
  await page.evaluate(() => (window as any).__ribbon.editor.exec('insertField', { code: 'MERGEFIELD Amount \\# "#,##0" \\* Upper' }));
  await openPane(page);
  const p = pane(page);
  await expect(p.locator('select[data-mf="format"]')).toHaveValue('other');
  await expect(p.locator('[data-mf="code"]')).toHaveValue('MERGEFIELD Amount \\# "#,##0" \\* Upper');
  // (Insert takes the code as it is: a second copy of the field)
  await expect(p.locator('[data-mf="insert"]')).toBeEnabled();
});

test('a code typed from scratch is inserted as typed', async ({ page }) => {
  await mount(page, true);
  await openPane(page);
  const box = pane(page).locator('textarea[data-mf="code"]');
  // (it holds the first name's code: select it and type over it)
  await box.click();
  await page.keyboard.press('Control+A');
  await page.keyboard.type('MERGEFIELD Total \\# "#,##0.0"');
  await expect(box).toHaveAttribute('data-edited', '');
  await pane(page).locator('input[data-mf="sample"]').click();
  await page.keyboard.type('1234.56');
  await expect(pane(page).locator('[data-mf="result"]')).toHaveText('1,234.6');
  await box.click();
  await page.keyboard.press('Enter');
  await expect(merge(page)).toHaveAttribute('data-field-code', 'MERGEFIELD Total \\# "#,##0.0"');
  // a name typed in the Field box makes the code anew, in the format shown (the inserted field's, at the caret)
  await pane(page).locator('input[data-mf="name"]').fill('');
  await pane(page).locator('input[data-mf="name"]').click();
  await page.keyboard.type('Nickname');
  await expect(box).toHaveValue('MERGEFIELD Nickname \\# "#,##0.0"');
  await expect(box).not.toHaveAttribute('data-edited', '');
});

test('the code of the field at the caret is edited in the code box and applied', async ({ page }) => {
  await mount(page, true);
  await page.evaluate(() => (window as any).__ribbon.editor.exec('insertField', { code: 'MERGEFIELD Amount \\# $,0.00' }));
  await openPane(page);
  const box = pane(page).locator('textarea[data-mf="code"]');
  await expect(box).toHaveValue('MERGEFIELD Amount \\# $,0.00');
  await box.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' \\f " USD"');
  await pane(page).locator('[data-mf="apply"]').click();
  await expect(merge(page)).toHaveCount(1);
  await expect(merge(page)).toHaveAttribute('data-field-code', 'MERGEFIELD Amount \\# $,0.00 \\f " USD"');
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

test('Select Recipients, Browse and Clear go to the host; without one they are hidden', async ({ page }) => {
  await mount(page, true);
  await page.locator(`${HOST} [data-dxr="mailmerge"]`).click();
  for (const recip of ['recipients', 'fieldlist', 'editrecipients', 'clearrecipients']) {
    await page.locator(`${HOST} [data-dxr="mailrecipients"]`).click();
    await page.locator(`${HOST} [data-recip="${recip}"]`).click();
  }
  await openPane(page);
  await pane(page).locator('[data-mf="browse"]').click();
  await pane(page).locator('[data-mf="clearsource"]').click();
  await page.locator(`${HOST} [data-dxr="mailfinishbtn"]`).click();
  expect(await page.evaluate(() => (window as any).__mailings))
    .toEqual(['start', 'recipients', 'fieldlist', 'editrecipients', 'clearrecipients', 'recipients', 'clearrecipients', 'finish']);
  // the host clears the source: the pane follows
  await page.evaluate(() => (window as any).__ribbon.setMergeColumns([], ''));
  await expect(pane(page).locator('[data-mf="source"]')).toHaveText('(none)');
  await expect(pane(page).locator('[data-mf="clearsource"]')).toBeDisabled();

  await page.evaluate(() => document.getElementById('mm-host')!.remove());
  await mount(page, false, false);
  await expect(page.locator(`${HOST} [data-dxr="mailstart"]`)).toBeHidden();
  await expect(page.locator(`${HOST} [data-dxr="mailfinish"]`)).toBeHidden();
  await openPane(page);
  await expect(pane(page).locator('[data-mf="browse"]')).toBeHidden();
  await expect(pane(page).locator('.dxr-mfempty')).toBeVisible();
});

// ─── Address Block, Greeting Line ─────────────────────────────────────────────

const PARTY_COLUMNS = ['LeFull', 'LeStatus', 'LeStreet', 'LeCity', 'LeState', 'LeZip', 'LeShort', 'BoFull', 'PropStreet', 'PropCity', 'PropState', 'PropZip', 'PropCounty'];
const PARTY_ROWS = [
  ['Don Dwyer', 'an unmarried man', '12 Elm St', 'Taos', 'NM', '87571', 'Don', 'Ann Lee', '9 Oak Rd', 'Santa Fe', 'NM', '87501', 'Santa Fe'],
  ['Westgro LLC', 'a New Mexico limited liability company', '', 'Albuquerque', 'NM', '87102', 'Westgro', 'Bo Two', '1 Pine', 'Taos', 'NM', '87571', 'Taos'],
];

async function mountParties(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(({ columns, rows }) => {
    const host = document.createElement('div');
    host.id = 'mm-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, { onMailings: () => {}, mergeColumns: columns, mergeSource: 'loan.csv', mergeRows: rows });
    r.setExports((window as any).Docxodus);
    r.openBlank('mm.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  }, { columns: PARTY_COLUMNS, rows: PARTY_ROWS });
  await page.locator(blocks).first().click();
  await page.locator(`${HOST} .dxr-tab[data-tab="mailings"]`).click();
}

const dialog = (page: Page) => page.locator(`${HOST} [data-dxr="addressdialog"]`);
const fieldCodes = (page: Page) => page.locator(`${HOST} [data-field-code]`).evaluateAll((els) => els.map((e) => e.getAttribute('data-field-code')));
const texts = (page: Page) => page.locator(blocks).evaluateAll((bs) => bs.map((b) => (b.textContent ?? '').replace(/^[\s\u00a0]+|[\s\u00a0]+$/g, '')));

test('Address Block: a party of prefixed columns, previewed with its rows, inserted as plain merge fields', async ({ page }) => {
  await mountParties(page);
  await page.locator(`${HOST} [data-dxr="addressblock"]`).click();
  const d = dialog(page);
  await expect(d).toBeVisible();
  await expect(d.locator('select[data-ad="prefix"] option')).toHaveText(['Le…', 'Prop…']);
  await expect(d.locator('select[data-ad="name"]')).toHaveValue('full');
  await expect(d.locator('[data-ad="preview"]')).toHaveText('Don Dwyer\n12 Elm St\nTaos, NM 87571');
  await expect(d.locator('[data-ad="record"]')).toHaveText('1 of 2');
  // the second recipient has no street: its line is left out, as the merge leaves it out
  await d.locator('[data-ad="next"]').click();
  await expect(d.locator('[data-ad="preview"]')).toHaveText('Westgro LLC\nAlbuquerque, NM 87102');
  // Match Fields shows the matches
  await d.locator('summary').click();
  await expect(d.locator('select[data-part="postal"]')).toHaveValue('LeZip');
  await expect(d.locator('select[data-part="first"]')).toHaveValue('');
  await d.locator('[data-ad="ok"]').click();
  await expect(d).toBeHidden();
  expect(await fieldCodes(page)).toEqual(['MERGEFIELD LeFull', 'MERGEFIELD LeStreet', 'MERGEFIELD LeCity', 'MERGEFIELD LeState', 'MERGEFIELD LeZip']);
  expect((await texts(page)).slice(0, 3)).toEqual(['«LeFull»', '«LeStreet»', '«LeCity», «LeState» «LeZip»']);
  // one undo step
  await page.evaluate(() => (window as any).__ribbon.editor.exec('undo', {}));
  expect(await fieldCodes(page)).toEqual([]);
});

test('Greeting Line: salutation, name and punctuation in the paragraph at the caret', async ({ page }) => {
  await mountParties(page);
  await page.locator(`${HOST} [data-dxr="greetingline"]`).click();
  const d = dialog(page);
  await expect(d.locator('[data-ad-mode="greeting"]').first()).toBeVisible();
  await expect(d.locator('select[data-ad="name"]')).toHaveValue('short');
  await expect(d.locator('[data-ad="preview"]')).toHaveText('Dear Don,');
  await d.locator('select[data-ad="prefix"]').selectOption('Le');
  await d.locator('select[data-ad="name"]').selectOption('full');
  await d.locator('select[data-ad="punctuation"]').selectOption(':');
  await expect(d.locator('[data-ad="preview"]')).toHaveText('Dear Don Dwyer:');
  await d.locator('[data-ad="ok"]').click();
  expect(await fieldCodes(page)).toEqual(['MERGEFIELD LeFull']);
  expect((await texts(page))[0]).toBe('Dear «LeFull»:');
});
