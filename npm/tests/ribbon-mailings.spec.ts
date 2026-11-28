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
// (the document's paragraphs: not a header's or footer's)
const texts = (page: Page) => page.locator(blocks).evaluateAll((bs) => bs.filter((b) => !b.hasAttribute('data-hf-anchor') && !b.closest('[data-hf-band]')).map((b) => (b.textContent ?? '').replace(/^[\s\u00a0]+|[\s\u00a0]+$/g, '')));

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

// ─── Rules ────────────────────────────────────────────────────────────────────

const ruleDialog = (page: Page) => page.locator(`${HOST} [data-dxr="ruledialog"]`);
async function rule(page: Page, kind: string) {
  await page.locator(`${HOST} [data-dxr="mergerules"]`).click();
  await page.locator(`${HOST} [data-rule="${kind}"]`).click();
}

test('Rules: If…Then…Else and Skip Record If from a comparison; Next Record at once', async ({ page }) => {
  await mountParties(page);
  await rule(page, 'if');
  const d = ruleDialog(page);
  await expect(d).toBeVisible();
  await d.locator('select[data-rd="field"]').selectOption('LeStatus');
  await d.locator('select[data-rd="op"]').selectOption('notblank');
  await expect(d.locator('input[data-rd="value"]')).toBeHidden();
  await d.locator('input[data-rd="then"]').fill(', ');
  await d.locator('input[data-rd="else"]').fill('');
  await expect(d.locator('[data-rd="code"]')).toHaveText('{ IF "{ MERGEFIELD LeStatus }" <> "" ", " "" }');
  await d.locator('[data-rd="ok"]').click();
  await expect(d).toBeHidden();
  await rule(page, 'skipif');
  await d.locator('select[data-rd="field"]').selectOption('LeZip');
  await d.locator('select[data-rd="op"]').selectOption('<');
  await d.locator('input[data-rd="value"]').fill('80000');
  await expect(d.locator('[data-rd="code"]')).toHaveText('{ SKIPIF "{ MERGEFIELD LeZip }" < "80000" }');
  await d.locator('[data-rd="ok"]').click();
  // (Word compares with wildcards: contains is = "*value*")
  await rule(page, 'if');
  await d.locator('select[data-rd="field"]').selectOption('LeShort');
  await d.locator('select[data-rd="op"]').selectOption('contains');
  await d.locator('input[data-rd="value"]').fill('Flanagan');
  await d.locator('input[data-rd="then"]').fill('Is a Flanagan');
  await d.locator('input[data-rd="else"]').fill('Is not a Flanagan');
  await expect(d.locator('[data-rd="code"]')).toHaveText('{ IF "{ MERGEFIELD LeShort }" = "*Flanagan*" "Is a Flanagan" "Is not a Flanagan" }');
  await d.locator('[data-rd="cancel"], button[value="cancel"]').first().click();
  await rule(page, 'next');
  expect(await fieldCodes(page)).toEqual(['IF "{ MERGEFIELD LeStatus }" <> "" ", " ""', 'SKIPIF "{ MERGEFIELD LeZip }" < "80000"', 'NEXT']);
});

test('Rules: Ask with its answer shown (ASK, then REF); a bookmark name is checked', async ({ page }) => {
  await mountParties(page);
  await rule(page, 'ask');
  const d = ruleDialog(page);
  await d.locator('input[data-rd="bookmark"]').fill('2bad');
  await expect(d.locator('[data-rd="ok"]')).toBeDisabled();
  await d.locator('input[data-rd="bookmark"]').fill('Closing');
  await d.locator('input[data-rd="prompt"]').fill('Closing date?');
  await d.locator('input[data-rd="default"]').fill('soon');
  await d.locator('input[data-rd="once"]').check();
  await expect(d.locator('[data-rd="code"]')).toHaveText('{ ASK Closing "Closing date?" \\d "soon" \\o } { REF Closing }');
  await d.locator('[data-rd="ok"]').click();
  await expect(d).toBeHidden();
  expect(await fieldCodes(page)).toEqual(['ASK Closing "Closing date?" \\d "soon" \\o']);
  await expect(page.locator(`${HOST} [data-field="REF"]`)).toHaveCount(1);
});

// ─── Envelopes ────────────────────────────────────────────────────────────────

const envDialog = (page: Page) => page.locator(`${HOST} [data-dxr="envelopedialog"]`);
/** The section of paragraph `index` (its size in twips). */
const sectionOf = (page: Page, index: number) => page.evaluate(async (index) => {
  const r = (window as any).__ribbon;
  const block = Array.from(document.querySelectorAll<HTMLElement>('#mm-host [data-anchor][contenteditable="true"]'))
    .filter((b) => !b.hasAttribute('data-hf-anchor') && !b.closest('[data-hf-band]'))[index];
  block.focus();
  getSelection()!.collapse(block, 0);
  await new Promise((done) => setTimeout(done, 50));
  const info = r.editor.sectionInfo();
  return { width: info.pageWidthTwips, height: info.pageHeightTwips, landscape: info.landscape, left: info.marginLeftTwips,
    top: info.marginTopTwips, firstPage: r.editor.headerFooterKindEnabled('first') };
}, index);

test('Envelopes: an empty document becomes the envelope, the delivery address merge fields', async ({ page }) => {
  await mountParties(page);
  await page.evaluate(() => {
    const r = (window as any).__ribbon;
    r.setReturnAddress('Burque Law\n1 Plaza');
    (window as any).__kept = [];
  });
  await page.locator(`${HOST} [data-dxr="envelopes"]`).click();
  const d = envDialog(page);
  await expect(d).toBeVisible();
  await expect(d.locator('textarea[data-ev="delivery"]')).toHaveValue('«LeFull»\n«LeStreet»\n«LeCity», «LeState» «LeZip»');
  await expect(d.locator('textarea[data-ev="return"]')).toHaveValue('Burque Law\n1 Plaza');
  await expect(d.locator('[data-ev="where"]')).toHaveText('This document becomes the envelope.');
  await d.locator('select[data-ev="size"]').selectOption('dl');
  // (text typed in the delivery address stays text; «Name» is a field)
  await d.locator('textarea[data-ev="delivery"]').fill('Attn: «LeFull»\n«LeStreet»\n«LeCity», «LeState» «LeZip»');
  await d.locator('[data-ev="ok"]').click();
  await expect(d).toBeHidden();
  expect(await fieldCodes(page)).toEqual(['MERGEFIELD LeFull', 'MERGEFIELD LeStreet', 'MERGEFIELD LeCity', 'MERGEFIELD LeState', 'MERGEFIELD LeZip']);
  expect((await texts(page)).slice(0, 5)).toEqual(['Burque Law', '1 Plaza', 'Attn: «LeFull»', '«LeStreet»', '«LeCity», «LeState» «LeZip»']);
  const section = await sectionOf(page, 0);
  expect(section.width).toBe(Math.round(220 * 1440 / 25.4));
  expect(section.height).toBe(Math.round(110 * 1440 / 25.4));
  expect(section.landscape).toBe(true);
});

test('Envelopes: a document with text gets the envelope in front, as a section of its own', async ({ page }) => {
  await mountParties(page);
  await page.evaluate(() => (window as any).__ribbon.editor.exec('pasteText', { text: 'Dear Sir,\nThe letter.' }));
  await page.locator(`${HOST} [data-dxr="envelopes"]`).click();
  const d = envDialog(page);
  await expect(d.locator('[data-ev="where"]')).toHaveText('The envelope goes in front of the document, as a section of its own.');
  await d.locator('input[data-ev="omit"]').check();
  await d.locator('[data-ev="ok"]').click();
  await expect(d).toBeHidden();
  const all = await texts(page);
  expect(all.slice(0, 3)).toEqual(['«LeFull»', '«LeStreet»', '«LeCity», «LeState» «LeZip»']);
  expect(all.slice(-2)).toEqual(['Dear Sir,', 'The letter.']);
  const envelope = await sectionOf(page, 0);
  const letter = await sectionOf(page, all.length - 1);
  expect(envelope.width).toBe(Math.round(9.5 * 1440));
  expect(letter.width).not.toBe(envelope.width);
});

test('Envelopes: half an inch in, and no page number on the envelope', async ({ page }) => {
  await mountParties(page);
  await page.evaluate(() => (window as any).__ribbon.editor.exec('pasteText', { text: 'The letter.' }));
  // a footer with a page number, as a letter has
  await page.locator(`${HOST} .dxr-tab[data-tab="references"]`).click();
  await page.locator(`${HOST} [data-dxr="pagenum"]`).click();
  await page.locator(`${HOST} .dxr-tab[data-tab="mailings"]`).click();
  await page.locator(`${HOST} [data-dxr="envelopes"]`).click();
  const d = envDialog(page);
  await expect(d.locator('input[data-ev="retleft"]')).toHaveValue('0.5');
  await expect(d.locator('input[data-ev="rettop"]')).toHaveValue('0.5');
  await expect(d.locator('input[data-ev="delleft"]')).toHaveValue('3.99');
  await expect(d.locator('input[data-ev="deltop"]')).toHaveValue('1.98');
  await d.locator('input[data-ev="retleft"]').fill('0.75');
  await d.locator('[data-ev="ok"]').click();
  await expect(d).toBeHidden();
  const envelope = await sectionOf(page, 0);
  expect(envelope.left).toBe(1080);
  expect(envelope.top).toBe(720);
  expect(envelope.firstPage).toBe(true);
  // page view: the envelope's page shows no footer text; the letter's page keeps its number
  await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
  const pages = page.locator(`${HOST} .page-box`);
  await expect(pages).toHaveCount(2);
  const footers = await pages.evaluateAll((ps) => ps.map((p) => Array.from(p.querySelectorAll<HTMLElement>('[data-docx-footer], [data-hf="footer"], .page-footer, footer'))
    .map((f) => (f.textContent ?? '').trim()).join('|')));
  expect(footers[0]).not.toMatch(/\d/);
  expect(footers[1]).toMatch(/\d/);
});

// ─── Labels ───────────────────────────────────────────────────────────────────

const lblDialog = (page: Page) => page.locator(`${HOST} [data-dxr="labelsdialog"]`);
/** The first table's shape and the Table Properties of its first cell. */
const labelTable = (page: Page) => page.evaluate(async () => {
  const r = (window as any).__ribbon;
  const table = document.querySelector<HTMLElement>('#mm-host table')!;
  const rows = Array.from(table.querySelectorAll('tr')).filter((tr) => tr.closest('table') === table);
  const first = table.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!;
  first.focus();
  getSelection()!.collapse(first, 0);
  await new Promise((done) => setTimeout(done, 50));
  const props = r.editor.tableProperties();
  const format = r.editor.paragraphFormats()[0];
  return { rows: rows.length, cells: rows[0].querySelectorAll('td').length, widths: props.columnWidthsTwips, height: props.rowHeightTwips,
    rule: props.rowHeightRule, margins: props.cellMargins, after: format?.spacingAfter ?? null };
});

test('Labels: Avery 5160 on an empty document, a different recipient on each label', async ({ page }) => {
  await mountParties(page);
  await page.locator(`${HOST} [data-dxr="labels"]`).click();
  const d = lblDialog(page);
  await expect(d).toBeVisible();
  await expect(d.locator('select[data-lb="product"]')).toHaveValue('5160');
  await expect(d.locator('textarea[data-lb="text"]')).toHaveValue('«LeFull»\n«LeStreet»\n«LeCity», «LeState» «LeZip»');
  await expect(d.locator('input[data-lb="next"]')).toBeChecked();
  await expect(d.locator('[data-lb="where"]')).toHaveText('This document becomes the sheet of labels.');
  await expect(d.locator('[data-lb="preview"] div')).toHaveCount(30);
  await d.locator('[data-lb="ok"]').click();
  await expect(d).toBeHidden();
  const codes = await fieldCodes(page);
  expect(codes.filter((c) => c === 'NEXT')).toHaveLength(29);
  expect(codes.filter((c) => c === 'MERGEFIELD LeFull')).toHaveLength(30);
  // (the first label has no Next Record; the second starts with it)
  expect(codes.slice(0, 6)).toEqual(['MERGEFIELD LeFull', 'MERGEFIELD LeStreet', 'MERGEFIELD LeCity', 'MERGEFIELD LeState', 'MERGEFIELD LeZip', 'NEXT']);
  const t = await labelTable(page);
  expect(t.rows).toBe(10);
  expect(t.cells).toBe(5);
  expect(t.widths).toEqual([3780, 180, 3780, 180, 3780]);
  expect(t.height).toBe(1440);
  expect(t.rule).toBe('exact');
  expect(t.margins?.leftTwips).toBe(86);
  expect(t.after).toBe(0);
  const section = await sectionOf(page, 0);
  expect(section.width).toBe(12240);
  expect(section.height).toBe(15840);
  expect(section.landscape).toBe(false);
  expect(section.left).toBe(270);
  expect(section.top).toBe(720);
});

test('Labels: a custom A4 sheet, every label the same, in front of a letter', async ({ page }) => {
  await mountParties(page);
  await page.evaluate(() => (window as any).__ribbon.editor.exec('pasteText', { text: 'The letter.' }));
  await page.locator(`${HOST} [data-dxr="labels"]`).click();
  const d = lblDialog(page);
  await expect(d.locator('[data-lb="where"]')).toHaveText('The labels go in front of the document, as a section of its own.');
  await d.locator('select[data-lb="product"]').selectOption('custom');
  // (a custom label starts from the product chosen before)
  await expect(d.locator('input[data-lb="hpitch"]')).toHaveValue('2.75');
  await d.locator('select[data-lb="page"]').selectOption('a4');
  await d.locator('input[data-lb="across"]').fill('2');
  await d.locator('input[data-lb="down"]').fill('3');
  await d.locator('input[data-lb="width"]').fill('3');
  await d.locator('input[data-lb="hpitch"]').fill('2');
  // the pitch less than the label: refused
  await expect(d.locator('[data-lb="ok"]')).toBeDisabled();
  await d.locator('input[data-lb="hpitch"]').fill('3');
  await d.locator('input[data-lb="vpitch"]').fill('2');
  await d.locator('input[data-lb="next"]').uncheck();
  await d.locator('textarea[data-lb="text"]').fill('Burque Law\n«LeFull»');
  await expect(d.locator('[data-lb="preview"] div')).toHaveCount(6);
  await d.locator('[data-lb="ok"]').click();
  await expect(d).toBeHidden();
  const codes = await fieldCodes(page);
  expect(codes).toEqual(Array(6).fill('MERGEFIELD LeFull'));
  const t = await labelTable(page);
  expect(t.rows).toBe(3);
  // (no gap: the pitch is the label's width)
  expect(t.widths).toEqual([4320, 4320]);
  expect(t.height).toBe(2880);
  const all = await texts(page);
  expect(all[all.length - 1]).toBe('The letter.');
  const sheet = await sectionOf(page, 0);
  const letter = await sectionOf(page, all.length - 1);
  expect(sheet.width).toBe(11906);
  expect(letter.width).not.toBe(11906);
});

test('Labels: Next Record stays a field of its own, the name after it (not inside its result)', async ({ page }) => {
  await mountParties(page);
  await page.locator(`${HOST} [data-dxr="labels"]`).click();
  await lblDialog(page).locator('[data-lb="ok"]').click();
  await expect(lblDialog(page)).toBeHidden();
  await expect(page.locator(`${HOST} [data-field-code="NEXT"]`).first()).toHaveText('«Next Record»');
  expect(await page.locator(`${HOST} [data-field-code="NEXT"] [data-field-code]`).count()).toBe(0);
});

test('Preview Results: each label a recipient, stepped through, and off again', async ({ page }) => {
  await mountParties(page);
  await expect(page.locator(`${HOST} [data-dxr="mailpreview"]`)).toBeVisible();
  await page.locator(`${HOST} [data-dxr="labels"]`).click();
  await lblDialog(page).locator('[data-lb="ok"]').click();
  await expect(lblDialog(page)).toBeHidden();
  const shown = () => page.locator(`${HOST} [data-field-code^="MERGEFIELD LeFull"]`).evaluateAll((els) => els.slice(0, 3).map((e) => e.getAttribute('data-merge-preview')));
  await expect(page.locator(`${HOST} [data-dxr="previewnext"]`)).toBeDisabled();
  await page.locator(`${HOST} [data-dxr="previewresults"]`).click();
  await expect.poll(shown).toEqual(['Don Dwyer', 'Westgro LLC', '']);
  await expect(page.locator(`${HOST} [data-dxr="previewcount"]`)).toHaveText('/ 2');
  await page.locator(`${HOST} [data-dxr="previewnext"]`).click();
  await expect.poll(shown).toEqual(['Westgro LLC', '', '']);
  await expect(page.locator(`${HOST} input[data-dxr="previewrecord"]`)).toHaveValue('2');
  await expect(page.locator(`${HOST} [data-dxr="previewnext"]`)).toBeDisabled();
  // (the document's own text is untouched)
  expect(await page.locator(`${HOST} [data-field-code="MERGEFIELD LeFull"]`).first().textContent()).toBe('«LeFull»');
  await page.locator(`${HOST} [data-dxr="previewresults"]`).click();
  await expect.poll(shown).toEqual([null, null, null]);
});
