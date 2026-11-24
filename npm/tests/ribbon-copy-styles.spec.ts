import { test, expect, type Page } from '@playwright/test';

/**
 * Copy Styles from Document (Word's Organizer, from a file into the document): the Styles box's
 * menu opens it, Browse lists the other file's styles beside this document's, Copy brings the
 * chosen ones in, and a style the document has already is replaced only after Word's question.
 */

const HOST = '#cs-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;
const dialog = (page: Page) => page.locator(`${HOST} dialog[data-dxr="copystylesdialog"]`);
const ask = (page: Page) => page.locator(`${HOST} dialog[data-dxr="copystylesask"]`);
const styleOf = (page: Page, id: string) =>
  page.evaluate((sid) => ((window as any).__ribbon.exec('styles') as any[]).find((s) => s.id === sid), id);

/** A ribbon on a blank document; returns the bytes of another document with an italic Normal and
 *  a centered "Fancy Note" style, to copy from. */
async function mount(page: Page): Promise<Buffer> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  const source = await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'cs-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('source.docx');
    ribbon.loader.done();
    (window as any).__ribbon = ribbon;
    ribbon.exec('setStyle', { styleId: 'Normal', run: { italic: true } });
    ribbon.exec('setStyle', { name: 'Fancy Note', paragraph: { alignment: 'center' } });
    const bytes = Array.from(ribbon.save() as Uint8Array);
    ribbon.openBlank('target.docx');
    return bytes;
  });
  await page.locator(blocks).first().click();
  await page.evaluate(() => (window as any).__ribbon.exec('pasteText', { text: 'First paragraph text\nSecond paragraph text' }));
  await page.locator(blocks).first().click();
  return Buffer.from(source);
}

async function openWith(page: Page, bytes: Buffer, name = 'styles.dotx') {
  await page.locator(`${HOST} [data-dxr="stylemenubtn"]`).click();
  await page.locator(`${HOST} [data-dxr="stylemenu"] [data-stylecmd="copy"]`).click();
  await expect(dialog(page)).toBeVisible();
  await dialog(page).locator('[data-cs="file"]').setInputFiles({ name, mimeType: 'application/octet-stream', buffer: bytes });
  await expect(dialog(page).locator('[data-cs="filename"]')).toHaveText(name);
}

const fontStyle = (page: Page, i: number) =>
  page.locator(blocks).nth(i).evaluate((el) => getComputedStyle(el.querySelector('span') ?? el).fontStyle);

test('a new style is copied at once, the dialog stays open, and one Undo takes it back', async ({ page }) => {
  const source = await mount(page);
  await openWith(page, source);
  const from = dialog(page).locator('[data-cs="from"]');
  await expect(from.locator('option', { hasText: 'Fancy Note' })).toHaveCount(1);
  // Normal is in both: marked.
  await expect(from.locator('option[value="Normal"]')).toContainText('✓');

  await from.selectOption('FancyNote');
  await expect(dialog(page).locator('[data-cs="desc"]')).toContainText('Paragraph style, based on Normal.');
  await dialog(page).locator('[data-cs="copy"]').click();
  await expect(ask(page)).toBeHidden();
  await expect(dialog(page)).toBeVisible();
  await expect(dialog(page).locator('[data-cs="desc"]')).toHaveText('Styles copied: 1');
  await expect(dialog(page).locator('[data-cs="to"] option[value="FancyNote"]')).toHaveCount(1);
  expect((await styleOf(page, 'FancyNote')).resolvedParagraph.alignment).toBe('center');
  // This document's Normal was not touched: Fancy Note is based on it as it is.
  expect((await styleOf(page, 'Normal')).resolvedRun.italic ?? false).toBe(false);

  await dialog(page).locator('button[value="cancel"]').click();
  await expect(dialog(page)).toBeHidden();
  await page.locator(`${HOST} [data-dxr="undo"]`).click();
  expect(await styleOf(page, 'FancyNote')).toBeUndefined();
});

test("a style the document has is replaced only after Word's question", async ({ page }) => {
  const source = await mount(page);
  await openWith(page, source);
  const from = dialog(page).locator('[data-cs="from"]');
  await from.selectOption('Normal');
  await expect(dialog(page).locator('[data-cs="desc"]')).toContainText('copying replaces it');

  await dialog(page).locator('[data-cs="copy"]').click();
  await expect(ask(page)).toBeVisible();
  await expect(ask(page).locator('[data-cs-ask]')).toHaveText('Do you want to replace the existing style Normal?');
  await ask(page).locator('button[value="no"]').click();
  await expect(ask(page)).toBeHidden();
  expect((await styleOf(page, 'Normal')).resolvedRun.italic ?? false).toBe(false);

  await from.selectOption('Normal');
  await dialog(page).locator('[data-cs="copy"]').click();
  await ask(page).locator('button[value="yes"]').click();
  await expect(dialog(page).locator('[data-cs="desc"]')).toHaveText('Styles copied: 1');
  expect((await styleOf(page, 'Normal')).resolvedRun.italic).toBe(true);
  await dialog(page).locator('button[value="cancel"]').click();
  for (const i of [0, 1]) expect(await fontStyle(page, i)).toBe('italic');
});

test('Select All with Yes to All copies every style; Cancel copies none', async ({ page }) => {
  const source = await mount(page);
  await openWith(page, source);
  await dialog(page).locator('[data-cs="all"]').click();
  await dialog(page).locator('[data-cs="copy"]').click();
  await ask(page).locator('button[value="cancel"]').click();
  expect(await styleOf(page, 'FancyNote')).toBeUndefined();

  await dialog(page).locator('[data-cs="all"]').click();
  await dialog(page).locator('[data-cs="copy"]').click();
  await ask(page).locator('button[value="all"]').click();
  await expect(ask(page)).toBeHidden();
  await expect(dialog(page).locator('[data-cs="desc"]')).toHaveText('Styles copied: 2');
  expect(await styleOf(page, 'FancyNote')).toBeDefined();
  expect((await styleOf(page, 'Normal')).resolvedRun.italic).toBe(true);
});

test('a file that is not a Word document is refused with a message', async ({ page }) => {
  await mount(page);
  await openWith(page, Buffer.from('not a package'), 'notes.docx');
  await expect(dialog(page).locator('[data-cs-msg="notWord"]')).toBeVisible();
  await dialog(page).locator('[data-cs="copy"]').click();
  await expect(dialog(page).locator('[data-cs-msg="none"]')).toBeVisible();
});
