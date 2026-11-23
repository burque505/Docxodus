import { test, expect, type Page } from '@playwright/test';

/**
 * Styles from the ribbon (Word's): the Font dialog (launcher, Ctrl+D), Modify Style and Create New
 * Style from Formatting (the Styles box's menu), its Format menu onto the Font and Paragraph
 * dialogs, Update to Match Selection, character styles in the Styles box, and Undo of a style edit.
 */

const HOST = '#st-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'st-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const reported: Array<{ label: string; command: any }> = [];
    (window as any).__reported = reported;
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {
      onCommand: (label: string, _ms: number, command?: unknown) => reported.push({ label, command }),
    });
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('styles.docx');
    ribbon.loader.done();
    (window as any).__ribbon = ribbon;
  });
  await page.locator(blocks).first().click();
  await page.evaluate(() => (window as any).__ribbon.exec('pasteText', { text: 'First paragraph text\nSecond paragraph text' }));
  await page.locator(blocks).first().click();
}

const styleDialog = (page: Page) => page.locator(`${HOST} dialog[data-dxr="styledialog"]`);
const fontDialog = (page: Page) => page.locator(`${HOST} dialog[data-dxr="fontdialog"]`);
const styleOf = (page: Page, id: string) =>
  page.evaluate((sid) => ((window as any).__ribbon.exec('styles') as any[]).find((s) => s.id === sid), id);
const commands = (page: Page) => page.evaluate(() => (window as any).__reported.map((r: any) => r.command?.name).filter(Boolean));

async function styleMenu(page: Page, item: 'modify' | 'new' | 'update') {
  await page.locator(`${HOST} [data-dxr="stylemenubtn"]`).click();
  await page.locator(`${HOST} [data-dxr="stylemenu"] [data-stylecmd="${item}"]`).click();
}

test('Ctrl+D opens the Font dialog; OK formats the selection as one fontFormat', async ({ page }) => {
  await mount(page);
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await page.keyboard.press('Control+d');
  await expect(fontDialog(page)).toBeVisible();
  await fontDialog(page).locator('[data-fd="size"]').fill('20');
  await fontDialog(page).locator('[data-fd="style"]').selectOption('bold');
  await fontDialog(page).locator('[data-fd="smallCaps"]').check();
  await fontDialog(page).locator('button[value="ok"]').click();
  await expect(fontDialog(page)).toBeHidden();
  const reported = await page.evaluate(() => (window as any).__reported.at(-1).command);
  expect(reported).toEqual({ name: 'fontFormat', args: { size: 20, bold: true, italic: false, smallCaps: true } });
  const weight = await page.locator(blocks).first().locator('span').first().evaluate((el) => getComputedStyle(el).fontWeight);
  expect(Number(weight)).toBeGreaterThanOrEqual(700);
});

test('Modify Style changes every paragraph in the style, and one Undo takes it back', async ({ page }) => {
  await mount(page);
  await styleMenu(page, 'modify');
  await expect(styleDialog(page)).toBeVisible();
  await expect(styleDialog(page).locator('[data-sd-title]')).toHaveText('Modify Style');
  await expect(styleDialog(page).locator('[data-sd="name"]')).toHaveValue('Normal');
  await styleDialog(page).locator('[data-sd="size"]').fill('15');
  await styleDialog(page).locator('[data-sd-toggle="italic"]').click();
  await expect(styleDialog(page).locator('[data-sd-desc]')).toContainText('15 pt, Italic');
  await styleDialog(page).locator('button[value="ok"]').click();
  await expect(styleDialog(page)).toBeHidden();

  expect((await styleOf(page, 'Normal')).resolvedRun).toMatchObject({ fontSizePts: 15, italic: true });
  expect(await commands(page)).toContain('setStyle');
  // Both paragraphs show it.
  for (const i of [0, 1]) {
    const style = await page.locator(blocks).nth(i).evaluate((el) => getComputedStyle(el.querySelector('span') ?? el).fontStyle);
    expect(style).toBe('italic');
  }

  await page.locator(`${HOST} [data-dxr="undo"]`).click();
  expect((await styleOf(page, 'Normal')).resolvedRun.italic ?? false).toBe(false);
  const after = await page.locator(blocks).nth(1).evaluate((el) => getComputedStyle(el.querySelector('span') ?? el).fontStyle);
  expect(after).toBe('normal');
});

test("Modify Style's Format menu opens the Paragraph dialog onto the style", async ({ page }) => {
  await mount(page);
  await styleMenu(page, 'modify');
  await styleDialog(page).locator('[data-sd-format]').click();
  await styleDialog(page).locator('[data-sd-open="paragraph"]').click();
  const para = page.locator(`${HOST} dialog[data-dxr="paradialog"]`);
  await expect(para).toBeVisible();
  await para.locator('[data-pd="after"]').fill('12');
  await para.locator('button[value="ok"]').click();
  await expect(para).toBeHidden();
  await expect(styleDialog(page)).toBeVisible();
  await expect(styleDialog(page).locator('[data-sd-desc]')).toContainText('Space After: 12 pt');
  await styleDialog(page).locator('button[value="ok"]').click();
  expect((await styleOf(page, 'Normal')).resolvedParagraph.spacingAfterTwips).toBe(240);
});

test('New Style starts from the caret formatting, is given to the paragraph, and refuses a taken name', async ({ page }) => {
  await mount(page);
  await page.evaluate(() => (window as any).__ribbon.exec('paragraphFormat', { alignment: 'center' }));
  await styleMenu(page, 'new');
  await expect(styleDialog(page).locator('[data-sd-title]')).toHaveText('Create New Style from Formatting');
  const name = styleDialog(page).locator('[data-sd="name"]');

  await name.fill('normal');
  await styleDialog(page).locator('button[value="ok"]').click();
  await expect(styleDialog(page)).toBeVisible();
  await expect(styleDialog(page).locator('[data-sd-msg="engine"]')).toBeVisible();

  await name.fill('Centered Note');
  await styleDialog(page).locator('button[value="ok"]').click();
  await expect(styleDialog(page)).toBeHidden();
  const style = await styleOf(page, 'CenteredNote');
  expect(style).toMatchObject({ name: 'Centered Note', basedOn: 'Normal', quickFormat: true });
  expect(style.resolvedParagraph.alignment).toBe('center');
  await expect(page.locator(`${HOST} [data-dxr="style"]`)).toHaveValue('CenteredNote');
});

test('Update to Match Selection puts the paragraph formatting into its style', async ({ page }) => {
  await mount(page);
  await page.evaluate(() => (window as any).__ribbon.exec('paragraphFormat', { leftIndent: 720 }));
  await styleMenu(page, 'update');
  expect((await styleOf(page, 'Normal')).resolvedParagraph.leftIndentTwips).toBe(720);
  // The other paragraph follows its style.
  await page.locator(blocks).nth(1).click();
  const formats = await page.evaluate(() => (window as any).__ribbon.exec('paragraphFormats'));
  expect(formats[0].leftIndent).toBe(720);
});

test('a gallery character style is offered in the Styles box and formats the selection', async ({ page }) => {
  await mount(page);
  await page.evaluate(() => (window as any).__ribbon.exec('setStyle', {
    name: 'Key Term', type: 'character', quickFormat: true, run: { bold: true, smallCaps: true },
  }));
  const box = page.locator(`${HOST} [data-dxr="style"]`);
  // (the box reads the styles again when the caret moves)
  await page.locator(blocks).nth(1).click();
  await expect(box.locator('optgroup option[value="KeyTerm"]')).toHaveCount(1);
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await box.selectOption('KeyTerm');
  const variant = await page.locator(blocks).nth(1).locator('span').first().evaluate((el) => getComputedStyle(el).fontVariantCaps);
  expect(variant).toBe('small-caps');
  expect(await commands(page)).toContain('characterStyle');
});

test('the fonts on the computer are offered in the Font box and both dialogs, whatever font is chosen', async ({ page }) => {
  await mount(page);
  await page.evaluate(() => (window as any).__ribbon.setSystemFonts(['Zephyr Sans', 'Bodoni Test', 'Calibri']));
  await expect(page.locator(`${HOST} [data-dxr="fontfamily"] option[value="Zephyr Sans"]`)).toHaveCount(1);
  await page.keyboard.press('Control+d');
  const font = fontDialog(page).locator('[data-fd="font"]');
  await expect(font.locator('option[value="Bodoni Test"]')).toHaveCount(1);
  await expect(font.locator('option[value="Zephyr Sans"]')).toHaveCount(1);
  await font.selectOption('Zephyr Sans');
  await fontDialog(page).locator('button[value="ok"]').click();
  expect(await page.evaluate(() => (window as any).__reported.at(-1).command.args.font)).toBe('Zephyr Sans');
  await styleMenu(page, 'modify');
  await expect(styleDialog(page).locator('[data-sd="font"] option[value="Bodoni Test"]')).toHaveCount(1);
});

test('a font applied goes first among the recently used fonts, at the top of the Font box and the dialogs', async ({ page }) => {
  await mount(page);
  await page.evaluate(() => {
    const r = (window as any).__ribbon;
    (window as any).__recent = [];
    r.setSystemFonts(['Zephyr Sans', 'Bodoni Test', 'Calibri', 'Arial']);
    r.setRecentFonts(['Arial']);
  });
  const box = page.locator(`${HOST} [data-dxr="fontfamily"]`);
  await expect(box.locator('optgroup').first()).toHaveAttribute('label', 'Recently used fonts');
  await box.selectOption({ label: 'Bodoni Test' });
  const recent = box.locator('optgroup').first().locator('option');
  await expect(recent).toHaveText(['Bodoni Test', 'Arial']);
  // (and still listed below: now in the document, so under "In this document")
  await expect(box.locator('option[value="Bodoni Test"]')).toHaveCount(2);

  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+d');
  const font = fontDialog(page).locator('[data-fd="font"]');
  await expect(font.locator('optgroup').first().locator('option')).toHaveText(['Bodoni Test', 'Arial']);
  await font.selectOption({ label: 'Zephyr Sans' });
  await fontDialog(page).locator('button[value="ok"]').click();
  await expect(recent).toHaveText(['Zephyr Sans', 'Bodoni Test', 'Arial']);
});

test('the recent fonts given at mount are listed, and a change is reported for the host to keep', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'st-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    (window as any).__kept = null;
    const r = (window as any).DocxodusEditor.mountRibbon(host, {
      recentFonts: ['Georgia', 'Arial'],
      onRecentFontsChange: (names: string[]) => { (window as any).__kept = names; },
    });
    r.setExports((window as any).Docxodus);
    r.openBlank('recent.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  const box = page.locator(`${HOST} [data-dxr="fontfamily"]`);
  await expect(box.locator('optgroup').first().locator('option')).toHaveText(['Georgia', 'Arial']);
  await page.locator(blocks).first().click();
  await box.selectOption({ label: 'Arial' });
  expect(await page.evaluate(() => (window as any).__kept)).toEqual(['Arial', 'Georgia']);
});
