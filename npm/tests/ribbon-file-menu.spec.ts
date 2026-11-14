import { test, expect, type Page } from '@playwright/test';

/**
 * The File menu: the File tab (and Alt+F) opens a menu with New, Open, Save, Save As (with
 * `onSaveAsRequest`), the host's recent documents (`setRecent`, `onOpenRecent`) and the host's own
 * commands (`fileMenu`, `setFileMenu`, `onFileMenu`).
 *
 * Each test mounts a second surface beside /editor.html's own, with the options under test, and
 * records what reaches the host in window.__events.
 */

async function mountHosted(page: Page, withHost: boolean) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.click('#new');
  await page.waitForFunction(() => !!(window as any).__demo.getEditor());
  await page.addScriptTag({ url: '/editor.bundle.js' });
  await page.evaluate((withHost) => {
    const { mountRibbon } = (window as any).DocxodusEditor;
    const demo = (window as any).__demo;
    const events: string[] = [];
    (window as any).__events = events;
    const host = document.createElement('div');
    host.id = 'hosted';
    host.style.cssText = 'height:500px';
    document.body.prepend(host);
    const hooks: Record<string, unknown> = withHost ? {
      onNewRequest: () => events.push('new'),
      onOpenRequest: () => events.push('open'),
      onSave: (_bytes: Uint8Array, name: string) => events.push(`save ${name}`),
      onSaveAsRequest: () => events.push('saveas'),
      onOpenRecent: (entry: { id: string }) => events.push(`recent ${entry.id}`),
      fileMenu: [{ id: 'print', label: 'Print…', shortcut: 'Ctrl+P' }, { id: 'close', label: 'Close', separatorBefore: true }],
      onFileMenu: (id: string) => events.push(`host ${id}`),
    } : {};
    const ribbon = mountRibbon(host, {
      exports: demo.exports, loader: false, hint: false, idPrefix: 'h-', ...hooks,
    });
    ribbon.open(demo.getEditor().save(), 'hosted.docx');
    (window as any).__hosted = ribbon;
  }, withHost);
}

const events = (page: Page) => page.evaluate(() => (window as any).__events as string[]);
const fileTab = (page: Page) => page.locator('#hosted [data-dxr="filebutton"]');
const menu = (page: Page) => page.locator('#hosted [data-dxr="filemenu"]');
const item = (page: Page, name: string) => menu(page).getByRole('menuitem', { name, exact: true });

test.describe('ribbon File menu', () => {
  test('the File tab opens the menu; each item reaches the host and closes the menu', async ({ page }) => {
    await mountHosted(page, true);
    await expect(menu(page)).toBeHidden();

    for (const [name, event] of [
      ['New', 'new'], ['Open…', 'open'], ['Save', 'save hosted.docx'], ['Save As…', 'saveas'], ['Print…', 'host print'], ['Close', 'host close'],
    ] as const) {
      await fileTab(page).click();
      await expect(menu(page)).toBeVisible();
      await expect(fileTab(page)).toHaveAttribute('aria-expanded', 'true');
      await item(page, name).click();
      await expect(menu(page)).toBeHidden();
      await expect.poll(() => events(page)).toContain(event);
    }
    // The File tab is not a panel tab: Home stays selected.
    await expect(page.locator('#hosted .dxr-tab[data-tab="home"]')).toHaveAttribute('aria-selected', 'true');
  });

  test('recent documents appear in the menu and open through onOpenRecent; setFileMenu replaces the host items', async ({ page }) => {
    await mountHosted(page, true);
    await page.evaluate(() => {
      (window as any).__hosted.setRecent([{ id: 'C:/a.docx', name: 'a.docx', title: 'C:/a.docx' }, { id: 'C:/b.docx', name: 'b.docx' }]);
      (window as any).__hosted.setFileMenu([{ id: 'settings', label: 'Settings', shortcut: 'Ctrl+,' }]);
    });
    await fileTab(page).click();
    await expect(menu(page).getByText('Recent')).toBeVisible();
    await expect(item(page, 'Print…')).toHaveCount(0);
    await expect(item(page, 'Settings')).toBeVisible();
    await item(page, 'b.docx').click();
    await expect.poll(() => events(page)).toContain('recent C:/b.docx');
  });

  test('Alt+F opens the menu with the first item focused; arrow keys move, Escape closes, a click outside closes', async ({ page }) => {
    await mountHosted(page, true);
    await page.locator('#hosted [data-anchor][contenteditable="true"]').first().click();
    await page.keyboard.press('Alt+F');
    await expect(menu(page)).toBeVisible();
    await expect(item(page, 'New')).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await expect(item(page, 'Open…')).toBeFocused();
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowUp');
    await expect(item(page, 'Close')).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(menu(page)).toBeHidden();

    await fileTab(page).click();
    await expect(menu(page)).toBeVisible();
    await page.locator('#hosted [data-anchor][contenteditable="true"]').first().click();
    await expect(menu(page)).toBeHidden();
  });

  test('without host hooks the menu has New, Open and Save, and no Save As, Recent or host items', async ({ page }) => {
    await mountHosted(page, false);
    await fileTab(page).click();
    await expect(item(page, 'New')).toBeVisible();
    await expect(item(page, 'Open…')).toBeVisible();
    await expect(item(page, 'Save')).toBeEnabled();
    await expect(item(page, 'Save As…')).toBeHidden();
    await expect(menu(page).getByText('Recent')).toBeHidden();
  });
});
