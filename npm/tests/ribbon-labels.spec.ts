import { test, expect, type Page } from '@playwright/test';

/**
 * The ribbon's text in another language (`labels`, `setLabels`): labels, menu entries and
 * tooltips are replaced by their translation, keyed by the English; text without one stays
 * English; the document's own text is never touched, even where it equals a key.
 */

async function mountHosted(page: Page, labels?: Record<string, string>) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.click('#new');
  await page.waitForFunction(() => !!(window as any).__demo.getEditor());
  await page.addScriptTag({ url: '/editor.bundle.js' });
  await page.evaluate((labels) => {
    const { mountRibbon } = (window as any).DocxodusEditor;
    const demo = (window as any).__demo;
    const host = document.createElement('div');
    host.id = 'hosted';
    host.style.cssText = 'height:500px';
    document.body.prepend(host);
    const ribbon = mountRibbon(host, { exports: demo.exports, loader: false, hint: false, idPrefix: 'h-', ...(labels ? { labels } : {}) });
    ribbon.open(demo.getEditor().save(), 'hosted.docx');
    (window as any).__hosted = ribbon;
  }, labels);
}

const spanish = { Home: 'Inicio', File: 'Archivo', 'Bold (Ctrl+B)': 'Negrita (Ctrl+B)', 'Save As…': 'Guardar como…', 'Find': 'Buscar' };
const tab = (page: Page, name: string) => page.locator(`#hosted .dxr-tab[data-tab="${name}"]`);

test.describe('ribbon labels', () => {
  test('labels translate tabs, buttons, menus and tooltips; text without one stays English', async ({ page }) => {
    await mountHosted(page, spanish);
    await expect(tab(page, 'home')).toHaveText('Inicio');
    await expect(tab(page, 'insert')).toHaveText('Insert');
    await expect(page.locator('#hosted [data-dxr="filebutton"]')).toHaveText('Archivo');
    await expect(page.locator('#hosted button[data-cmd="bold"]')).toHaveAttribute('title', 'Negrita (Ctrl+B)');
    await expect(page.locator('#hosted [data-file="saveas"] span')).toHaveText('Guardar como…');
  });

  test('the document is never translated, and setLabels({}) brings the English back', async ({ page }) => {
    await mountHosted(page);
    // A paragraph whose text is exactly a key.
    const block = page.locator('#hosted [data-anchor][contenteditable="true"]').first();
    await block.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.type('Home');
    await page.locator('#hosted .dxr-tab[data-tab="insert"]').click();
    await page.evaluate((labels) => (window as any).__hosted.setLabels(labels), spanish);
    await expect(tab(page, 'home')).toHaveText('Inicio');
    await expect(page.locator('#hosted [data-anchor]').filter({ hasText: /^Home$/ })).toHaveCount(1);

    await page.evaluate(() => (window as any).__hosted.setLabels({}));
    await expect(tab(page, 'home')).toHaveText('Home');
    await expect(page.locator('#hosted button[data-cmd="bold"]')).toHaveAttribute('title', 'Bold (Ctrl+B)');
    // A second language replaces the first, from the English.
    await page.evaluate(() => (window as any).__hosted.setLabels({ Home: 'Accueil' }));
    await expect(tab(page, 'home')).toHaveText('Accueil');
  });
});
