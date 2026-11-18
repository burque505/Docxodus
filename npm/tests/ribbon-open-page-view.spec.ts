import { test, expect, type Page } from '@playwright/test';

/**
 * Opening a document while the ribbon is in page view renders it straight into pages (once),
 * instead of failing or opening continuous and switching (twice the work).
 */

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'pv-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('first.docx');
    ribbon.loader.done();
    (window as any).__ribbon = ribbon;
  });
}

const surfaces = (page: Page) => page.locator('#pv-host .dxr-surface');
const blocks = (page: Page) => page.locator('#pv-host [data-dxr="editor"] [data-anchor][contenteditable="true"]');

test('New and Open in page view render pages directly, and typing works', async ({ page }) => {
  await mount(page);
  await page.evaluate(() => (window as any).__ribbon.exec('pageView', { on: true }));
  await expect(page.locator('#pv-host [data-dxr="editor"] .page-box').first()).toBeVisible();
  // New, while in page view
  await page.evaluate(() => (window as any).__ribbon.openBlank('second.docx'));
  await expect(surfaces(page)).toHaveCount(1);
  await expect(page.locator('#pv-host [data-dxr="editor"]')).toHaveAttribute('data-view', 'paginated');
  await expect(page.locator('#pv-host [data-dxr="editor"] .page-box').first()).toBeVisible();
  await blocks(page).first().click();
  await page.keyboard.type('Hello pages');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second');
  await expect(blocks(page)).toHaveCount(2);
  // Open the saved bytes, still in page view
  const count = await page.evaluate(() => {
    const ribbon = (window as any).__ribbon;
    const bytes = ribbon.save();
    ribbon.open(bytes, 'third.docx');
    return document.querySelectorAll('#pv-host .dxr-surface').length;
  });
  expect(count).toBe(1);
  await expect(page.locator('#pv-host [data-dxr="editor"] .page-box').first()).toBeVisible();
  await expect(blocks(page)).toHaveText([/Hello pages/, /Second/]);
  await blocks(page).nth(1).click();
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  await expect(blocks(page).nth(1)).toHaveText(/Second!/);
});

test('a failed open in page view leaves no hidden surface behind', async ({ page }) => {
  await mount(page);
  await page.evaluate(() => (window as any).__ribbon.exec('pageView', { on: true }));
  const failed = await page.evaluate(() => {
    try { (window as any).__ribbon.open(new Uint8Array([1, 2, 3]), 'bad.docx'); return false; } catch { return true; }
  });
  expect(failed).toBe(true);
  await expect(surfaces(page)).toHaveCount(1);
});
