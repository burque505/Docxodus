import { test, expect, type Page } from '@playwright/test';

/**
 * Pictures: a click selects one (outline, eight handles, the Picture Format tab); Delete removes
 * it; the handles resize it; Wrap Text makes it float and back; the picture dialog sets size,
 * position and alt text; the right-click menu offers the same.
 */

const HOST = '#pi-host';
const blocks = (page: Page) => page.locator(`${HOST} [data-anchor][contenteditable="true"]`);
const pictures = (page: Page) => page.locator(`${HOST} [data-anchor] img`);
const exec = (page: Page, name: string, args: object = {}) =>
  page.evaluate(([n, a]) => (window as any).__ribbon.exec(n, a), [name, args] as const);
const picture = (page: Page) => exec(page, 'picture') as Promise<any>;
const tab = (page: Page) => page.locator(`${HOST} .dxr-tab[data-tab="picture"]`);
const overlay = (page: Page) => page.locator('.docx-picture-overlay');

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  const png = await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'pi-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('pictures.docx');
    ribbon.loader.done();
    (window as any).__ribbon = ribbon;
    // A 96 x 48 pixel picture: 72 x 36 points at 96 dpi.
    const canvas = document.createElement('canvas');
    canvas.width = 96; canvas.height = 48;
    const g = canvas.getContext('2d')!;
    g.fillStyle = '#c33'; g.fillRect(0, 0, 96, 48);
    return canvas.toDataURL('image/png').split(',')[1];
  });
  await blocks(page).first().click();
  await page.keyboard.type('Before the picture and after it, enough words to wrap around a floating picture.');
  await page.keyboard.press('Home');
  for (let i = 0; i < 'Before the picture'.length; i++) await page.keyboard.press('ArrowRight');
  await exec(page, 'insertPicture', { base64: png, altText: 'red box' });
  await expect(pictures(page)).toHaveCount(1);
  return png;
}

test.describe('Pictures', () => {
  test('a click selects the picture, shows the Picture Format tab, and Escape drops it', async ({ page }) => {
    await mount(page);
    await expect(tab(page)).toBeHidden();
    await pictures(page).first().click();
    await expect(tab(page)).toBeVisible();
    await expect(overlay(page)).toBeVisible();
    await expect(overlay(page).locator('.docx-picture-handle')).toHaveCount(8);
    const info = await picture(page);
    expect(info.wrap).toBe('inline');
    expect(info.altText).toBe('red box');
    expect(info.widthPoints).toBeCloseTo(72, 1);
    expect(info.heightPoints).toBeCloseTo(36, 1);
    await page.keyboard.press('Escape');
    await expect(tab(page)).toBeHidden();
    await expect(overlay(page)).toBeHidden();
    expect(await picture(page)).toBeNull();
  });

  test('Delete removes the selected picture and Undo brings it back', async ({ page }) => {
    await mount(page);
    await pictures(page).first().click();
    await page.keyboard.press('Delete');
    await expect(pictures(page)).toHaveCount(0);
    await expect(tab(page)).toBeHidden();
    await expect(blocks(page).first()).toContainText('Before the picture and after it');
    await exec(page, 'undo');
    await expect(pictures(page)).toHaveCount(1);
  });

  test('the Height box keeps the shape, and Undo takes it back', async ({ page }) => {
    await mount(page);
    await pictures(page).first().click();
    await tab(page).click();
    const height = page.locator(`${HOST} [data-dxr="picheight"]`);
    await height.fill('1"');
    await height.press('Enter');
    await expect.poll(async () => (await picture(page))?.heightPoints).toBeCloseTo(72, 1);
    expect((await picture(page)).widthPoints).toBeCloseTo(144, 1);
    // Still selected, on the new picture element.
    await expect(overlay(page)).toBeVisible();
    await exec(page, 'undo');
    await exec(page, 'selectPicture');
    expect((await picture(page)).widthPoints).toBeCloseTo(72, 1);
  });

  test('dragging a corner handle resizes in proportion', async ({ page }) => {
    await mount(page);
    await pictures(page).first().click();
    const handle = overlay(page).locator('.docx-picture-handle[data-h="se"]');
    const box = (await handle.boundingBox())!;
    const before = (await pictures(page).first().boundingBox())!;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + before.width, box.y + box.height / 2 + 5, { steps: 5 });
    await page.mouse.up();
    await expect.poll(async () => (await picture(page))?.widthPoints).toBeCloseTo(144, 0);
    expect((await picture(page)).heightPoints).toBeCloseTo(72, 0);
  });

  test('Wrap Text > Square floats the picture, and In Line with Text puts it back', async ({ page }) => {
    await mount(page);
    await pictures(page).first().click();
    await tab(page).click();
    await page.locator(`${HOST} [data-dxr="picwrap"]`).selectOption('square');
    await expect.poll(async () => (await picture(page))?.wrap).toBe('square');
    const float = await pictures(page).first().evaluate((img) =>
      getComputedStyle(img.closest('span[style*="float"]') ?? img).float);
    expect(float).toBe('left');
    await page.locator(`${HOST} [data-dxr="picwrap"]`).selectOption('inline');
    await expect.poll(async () => (await picture(page))?.wrap).toBe('inline');
    await expect(pictures(page)).toHaveCount(1);
  });

  test('the dialog sets alt text, and wrapping with a position', async ({ page }) => {
    await mount(page);
    await pictures(page).first().click();
    await tab(page).click();
    await page.locator(`${HOST} [data-dxr="picalt"]`).click();
    const dialog = page.locator(`${HOST} [data-dxr="picturedialog"]`);
    await expect(dialog).toBeVisible();
    await dialog.locator('[data-pd2="alt"]').fill('A red rectangle');
    await dialog.locator('[data-pd2="title"]').fill('Box');
    await dialog.locator('[data-pd2-tab="layout"]').click();
    await dialog.locator('[data-pd2="wrap"]').selectOption('behind');
    await dialog.locator('[data-pd2="hmode-align"]').check();
    await dialog.locator('[data-pd2="halign"]').selectOption('center');
    await dialog.locator('[data-pd2="hfrom-align"]').selectOption('margin');
    await dialog.locator('button[value="ok"]').click();
    await expect(dialog).toBeHidden();
    await expect.poll(async () => (await picture(page))?.wrap).toBe('behind');
    const info = await picture(page);
    expect(info.altText).toBe('A red rectangle');
    expect(info.title).toBe('Box');
    expect(info.layout.horizontalAlignment).toBe('center');
    expect(info.layout.horizontalRelativeFrom).toBe('margin');
    expect(info.layout.behindDocument).toBe(true);
  });

  test('the right-click menu wraps the picture', async ({ page }) => {
    await mount(page);
    await pictures(page).first().click({ button: 'right' });
    const menu = page.locator(`${HOST} [data-dxr="picturemenu"]`);
    await expect(menu).toBeVisible();
    await expect(menu.locator('[data-pm-wrap="inline"]')).toHaveAttribute('aria-checked', 'true');
    await menu.locator('[data-pm-wrap="topAndBottom"]').click();
    await expect(menu).toBeHidden();
    await expect.poll(async () => (await picture(page))?.wrap).toBe('topAndBottom');
  });

  test('typing after selecting a picture goes on after it', async ({ page }) => {
    await mount(page);
    await pictures(page).first().click();
    await page.keyboard.type('X');
    await expect(tab(page)).toBeHidden();
    await page.locator(`${HOST} .dxr-tab[data-tab="home"]`).click();
    // The X is the first thing after the picture.
    await expect.poll(() => pictures(page).first().evaluate((img) => {
      const range = document.createRange();
      range.setStartAfter(img);
      range.setEndAfter(img.closest('[data-anchor]')!.lastChild!);
      return range.toString().slice(0, 1);
    })).toBe('X');
  });
});
