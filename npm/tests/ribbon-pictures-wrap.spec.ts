import { test, expect } from '@playwright/test';
/** Each Wrap Text choice, in continuous and page view: the picture stays selected, and can be
 *  clicked again and put back in line (page view lifts floating pictures out of their paragraph). */
const HOST = '#pi-host';
for (const paginated of [false, true]) for (const wrap of ['square','tight','topAndBottom','behind','inFront']) {
test(`wrap ${wrap} paginated=${paginated}`, async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((pg) => {
    const host = document.createElement('div'); host.id = 'pi-host'; host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, { paginated: pg });
    r.setExports((window as any).Docxodus); r.openBlank('p.docx'); r.loader.done(); (window as any).__ribbon = r;
  }, paginated);
  const png = await page.evaluate(() => { const c = document.createElement('canvas'); c.width=96;c.height=48; const g=c.getContext('2d')!; g.fillStyle='#c33'; g.fillRect(0,0,96,48); return c.toDataURL('image/png').split(',')[1]; });
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.keyboard.type('Some words before and after.');
  await page.keyboard.press('Home');
  await page.evaluate((b) => (window as any).__ribbon.exec('insertPicture', { base64: b }), png);
  await page.locator(`${HOST} [data-anchor] img`).first().click();
  await page.locator(`${HOST} .dxr-tab[data-tab="picture"]`).click();
  await page.locator(`${HOST} [data-dxr="picwrap"]`).selectOption(wrap);
  await page.waitForTimeout(500);
  const info = await page.evaluate(() => (window as any).__ribbon.exec('picture'));
  expect(info?.wrap).toBe(wrap);
  await page.keyboard.press('Escape');
  expect(await page.evaluate(() => (window as any).__ribbon.exec('picture'))).toBeNull();
  if (wrap !== 'behind') {
    await page.locator('#pi-host img').first().click();
    expect((await page.evaluate(() => (window as any).__ribbon.exec('picture')))?.wrap).toBe(wrap);
    await page.locator(`${HOST} .dxr-tab[data-tab="picture"]`).click();
    await page.locator(`${HOST} [data-dxr="picwrap"]`).selectOption('inline');
    await expect.poll(async () => (await page.evaluate(() => (window as any).__ribbon.exec('picture')))?.wrap).toBe('inline');
  }
});
}
