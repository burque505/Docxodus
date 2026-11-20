import { test, expect, type Page } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/**
 * Removing a section break, as Word: Layout > Breaks > Remove section break, Backspace at a
 * section's first paragraph, Delete at a section's last; the text before the break takes the layout
 * of the section after it. With Show ¶ the page view marks each break.
 */

const HOST = '#sb-host';

/** One column (First, Second, then an empty paragraph holding the break), two columns (Left
 *  column, More text, break paragraph), one column (After, Last). Both later sections continuous. */
function fixture(): Uint8Array {
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

async function mount(page: Page, paginated: boolean) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(({ bytes, paginated }) => {
    const host = document.createElement('div');
    host.id = 'sb-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.open(new Uint8Array(bytes), 'sections.docx');
    r.loader.done();
    (window as any).__ribbon = r;
    if (paginated) r.editor.setPaginated(true);
  }, { bytes: Array.from(fixture()), paginated });
}

const block = (page: Page, text: string) =>
  page.locator(`${HOST} [data-anchor][contenteditable="true"]`, { hasText: new RegExp(`^${text}$`) }).first();
/** How many columns the section holding `text` has (the engine's answer). */
const columnsAt = async (page: Page, text: string) => {
  await block(page, text).click();
  return page.evaluate(() => (window as any).__ribbon.editor.sectionInfo()?.columns);
};
const sectionCount = (page: Page) => page.evaluate(() => {
  const keys = new Set(Array.from(document.querySelectorAll<HTMLElement>('#sb-host [data-anchor][contenteditable="true"]'))
    .map((b) => b.closest<HTMLElement>('[data-pagination-section]')?.dataset.paginationSection
      ?? b.closest<HTMLElement>('[data-section-index]')?.dataset.sectionIndex));
  return keys.size;
});
const removeFromMenu = async (page: Page) => {
  await page.locator(`${HOST} .dxr-tab[data-tab="layout"]`).click();
  await page.locator(`${HOST} [data-dxr="breaks"]`).selectOption('removeSection');
};

test('the menu removes the break ending the caret\'s section: its text takes the next section\'s columns; Undo puts it back', async ({ page }) => {
  await mount(page, false);
  expect(await columnsAt(page, 'First')).toBe(1);
  await removeFromMenu(page);
  await expect.poll(() => sectionCount(page)).toBe(2);
  expect(await columnsAt(page, 'First')).toBe(2);
  await page.keyboard.press('Control+z');
  await expect.poll(() => sectionCount(page)).toBe(3);
  expect(await columnsAt(page, 'First')).toBe(1);
});

test('in the last section the menu removes the break before it', async ({ page }) => {
  await mount(page, true);
  expect(await columnsAt(page, 'Left column')).toBe(2);
  await block(page, 'Last').click();
  await removeFromMenu(page);
  await expect.poll(() => sectionCount(page)).toBe(2);
  expect(await columnsAt(page, 'Left column')).toBe(1);
});

test('Backspace at a section\'s first paragraph removes the break before it', async ({ page }) => {
  await mount(page, false);
  await block(page, 'Left column').click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Backspace');
  await expect.poll(() => sectionCount(page)).toBe(2);
  expect(await columnsAt(page, 'Second')).toBe(2);
});

test('Delete at the end of a paragraph before an empty break paragraph joins it and keeps the break', async ({ page }) => {
  await mount(page, false);
  await block(page, 'Second').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Delete');
  // As Word: Second's own mark goes, and the empty paragraph's mark (the break) now ends Second.
  const texts = () => page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#sb-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').trim()));
  await expect.poll(texts).toEqual(['First', 'Second', 'Left column', 'More text', '', 'After', 'Last']);
  expect(await sectionCount(page)).toBe(3);
  expect(await columnsAt(page, 'First')).toBe(1);
  expect(await columnsAt(page, 'Left column')).toBe(2);
});

test('with Show ¶ the page view marks each section break', async ({ page }) => {
  await mount(page, true);
  await page.evaluate(() => (window as any).__ribbon.editor.setShowFormattingMarks?.(true) ?? (window as any).__ribbon.setShowFormattingMarks(true));
  const marks = () => page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#sb-host [data-section-break]'))
    .map((b) => b.dataset.sectionBreak));
  await expect.poll(marks).toEqual(['continuous', 'continuous']);
  const label = await page.evaluate(() => getComputedStyle(document.querySelector('#sb-host [data-section-break]')!, '::after').content);
  expect(label).toContain('Section break (Continuous)');
});
