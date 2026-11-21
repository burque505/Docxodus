import { test, expect } from '@playwright/test';

/** The Styles box shows the caret paragraph's style, as Word's style gallery marks it. */

const HOST = '#sb-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

test('the Styles box follows the caret, and shows a style just applied', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'sb-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('sb.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate(() => (window as any).__ribbon.exec('pasteText', { text: 'Title here\nBody text' }));
  const box = page.locator(`${HOST} [data-dxr="style"]`);

  await page.locator(blocks).first().click();
  await box.selectOption('Heading2');
  await expect(box).toHaveValue('Heading2');
  await page.locator(blocks).nth(1).click();
  await expect(box).toHaveValue('Normal');
  await page.locator(blocks).first().click();
  await expect(box).toHaveValue('Heading2');
});
