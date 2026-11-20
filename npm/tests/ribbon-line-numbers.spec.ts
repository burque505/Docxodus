import { test, expect, type Page } from '@playwright/test';

/**
 * Layout > Line Numbers: Word's menu (None, Continuous, Restart Each Page, Restart Each Section,
 * Suppress for Current Paragraph, Line Numbering Options...) and its dialog. The numbers show in
 * page view; in the continuous view the status line says so.
 */

const HOST = '#lr-host';

async function mount(page: Page, paginated: boolean) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'lr-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('lr.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  for (const text of ['Alpha', 'Beta', 'Gamma']) {
    await page.keyboard.type(text);
    if (text !== 'Gamma') await page.keyboard.press('Enter');
  }
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.locator(`${HOST} .dxr-tab[data-tab="layout"]`).click();
}

const numbers = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll('#lr-host .docx-line-numbers span')).map((s) => s.textContent));
const lineNumbering = (page: Page) => page.evaluate(() => (window as any).__ribbon.editor.lineNumbering());
const choose = (page: Page, value: string) => page.locator(`${HOST} [data-dxr="linenumbers"]`).selectOption(value);
const ln = (page: Page, name: string) => page.locator(`${HOST} dialog[data-dxr="linenumbersdialog"] [data-ln="${name}"]`);

test('the menu turns numbering on and off; Undo brings it back', async ({ page }) => {
  await mount(page, true);
  await choose(page, 'continuous');
  await expect.poll(() => numbers(page)).toEqual(['1', '2', '3']);
  expect(await lineNumbering(page)).toMatchObject({ countBy: 1, start: 1, restart: 'continuous' });
  await choose(page, 'none');
  await expect.poll(() => numbers(page)).toEqual([]);
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.keyboard.press('Control+z');
  await expect.poll(() => numbers(page)).toEqual(['1', '2', '3']);
});

test('Suppress for Current Paragraph leaves that paragraph out of the count', async ({ page }) => {
  await mount(page, true);
  await choose(page, 'newPage');
  await expect.poll(() => numbers(page)).toEqual(['1', '2', '3']);
  await page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).nth(1).click();
  await choose(page, 'suppress');
  await expect.poll(() => numbers(page)).toEqual(['1', '2']);
});

test('the options dialog sets start, count by and restart, and checks what is typed', async ({ page }) => {
  await mount(page, true);
  await choose(page, 'more');
  const dialog = page.locator(`${HOST} dialog[data-dxr="linenumbersdialog"]`);
  await expect(dialog).toBeVisible();
  await expect(ln(page, 'start')).toBeDisabled();
  await ln(page, 'on').check();
  await ln(page, 'start').fill('0');
  await dialog.locator('button[value="ok"]').click();
  await expect(dialog).toBeVisible();
  await expect(ln(page, 'start')).toHaveAttribute('aria-invalid', 'true');
  await ln(page, 'start').fill('5');
  await ln(page, 'countBy').fill('2');
  await ln(page, 'restart-newSection').check();
  await ln(page, 'applyTo').selectOption('document');
  await dialog.locator('button[value="ok"]').click();
  await expect(dialog).toBeHidden();
  expect(await lineNumbering(page)).toMatchObject({ countBy: 2, start: 5, restart: 'newSection' });
  // 5, 6, 7: only the multiples of 2 show.
  await expect.poll(() => numbers(page)).toEqual(['6']);
});

test('in the continuous view the status line says line numbers show in page view', async ({ page }) => {
  await mount(page, false);
  await choose(page, 'continuous');
  await expect(page.locator(`${HOST} [data-dxr="status"]`)).toContainText('Line numbers show in page view');
  expect(await numbers(page)).toEqual([]);
  await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
  await expect.poll(() => numbers(page)).toEqual(['1', '2', '3']);
});

import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Plain text, then (continuous breaks) a two-column section, then plain text again. */
function sectionsFixture(): Uint8Array {
  const p = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const pg = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
  const body = p('First') + p('Second')
    + `<w:p><w:pPr><w:sectPr>${pg}</w:sectPr></w:pPr></w:p>`
    + p('Left column') + p('More text')
    + `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${pg}<w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr></w:p>`
    + p('After') + p('Last')
    + `<w:sectPr><w:type w:val="continuous"/>${pg}</w:sectPr>`;
  const rel = `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`;
  return storedZip([
    { name: '[Content_Types].xml', data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>') },
    { name: '_rels/.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`) },
  ]);
}

test('the menu numbers every section, wherever the caret is', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((bytes) => {
    const host = document.createElement('div');
    host.id = 'lr-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.open(new Uint8Array(bytes), 'sections.docx');
    r.loader.done();
    (window as any).__ribbon = r;
    r.editor.setPaginated(true);
  }, Array.from(sectionsFixture()));
  await page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`, { hasText: 'Last' }).click();
  await page.locator(`${HOST} .dxr-tab[data-tab="layout"]`).click();
  await choose(page, 'continuous');
  // First, Second, Left column, More text, After, Last: six lines, numbered straight through.
  await expect.poll(() => numbers(page)).toEqual(['1', '2', '3', '4', '5', '6']);
  await choose(page, 'newSection');
  await expect.poll(() => numbers(page)).toEqual(['1', '2', '1', '2', '1', '2']);
});
