import { test, expect, type Page } from '@playwright/test';

/**
 * Home > Font box shows the font where the caret is, as Word's does (and as the Styles box shows
 * the style): the picked font after a pick, the run's font as the caret moves, blank for a
 * selection of several fonts.
 */

const HOST = '#fb-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;
const box = `${HOST} [data-dxr="fontfamily"]`;

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'fb-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('fb.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate(() => (window as any).__ribbon.exec('pasteText', { text: 'plain words\nsecond line' }));
}

const shown = (page: Page) => page.locator(box).evaluate((s: HTMLSelectElement) => s.options[s.selectedIndex]?.textContent ?? '');

test('the box shows the caret paragraph\'s font, not "Font…"', async ({ page }) => {
  await mount(page);
  await page.locator(blocks).nth(1).click();
  await expect.poll(() => shown(page)).not.toBe('');
  expect(await shown(page)).not.toContain('Font');
});

test('a picked font shows; moving into other text shows its font; a mixed selection shows blank', async ({ page }) => {
  await mount(page);
  const before = await (async () => { await page.locator(blocks).nth(1).click(); await expect.poll(() => shown(page)).not.toBe(''); return shown(page); })();
  await page.locator(blocks).nth(0).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await page.locator(box).selectOption('Georgia');
  await expect.poll(() => shown(page)).toBe('Georgia');

  await page.locator(blocks).nth(1).click();
  await expect.poll(() => shown(page)).toBe(before);
  await page.locator(blocks).nth(0).click();
  await expect.poll(() => shown(page)).toBe('Georgia');

  // From the first paragraph into the second: two fonts.
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Control+Shift+End');
  await expect.poll(() => shown(page)).toBe('');
});
