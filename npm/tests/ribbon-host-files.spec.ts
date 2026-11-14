import { test, expect, type Page } from '@playwright/test';

/**
 * A host that keeps files itself (a desktop program): the ribbon tells it when the document has
 * changed (`isDirty`, `onDirtyChange`, `markSaved`), lets it take over New and Open
 * (`onNewRequest`, `onOpenRequest`), shows its recent documents (`setRecent`, `onOpenRecent`), and
 * saves and opens on Ctrl+S and Ctrl+O.
 *
 * Each test mounts a second surface beside /editor.html's own, with the options under test, and
 * records what reaches the host in window.__events.
 */

async function mountHosted(page: Page, options: Record<string, boolean>) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.click('#new');
  await page.waitForFunction(() => !!(window as any).__demo.getEditor());
  await page.addScriptTag({ url: '/editor.bundle.js' });
  await page.evaluate((options) => {
    const { mountRibbon } = (window as any).DocxodusEditor;
    const demo = (window as any).__demo;
    const events: string[] = [];
    (window as any).__events = events;
    const host = document.createElement('div');
    host.id = 'hosted';
    host.style.cssText = 'height:500px';
    document.body.prepend(host);
    const hooks: Record<string, unknown> = {};
    if (options.dirty) hooks.onDirtyChange = (dirty: boolean) => events.push(`dirty ${dirty}`);
    if (options.onSave) hooks.onSave = (bytes: Uint8Array, name: string) => events.push(`save ${name} ${bytes.length > 0}`);
    if (options.onNewRequest) hooks.onNewRequest = () => events.push('new');
    if (options.onOpenRequest) hooks.onOpenRequest = () => events.push('open');
    if (options.onOpenRecent) hooks.onOpenRecent = (entry: { id: string }) => events.push(`recent ${entry.id}`);
    const ribbon = mountRibbon(host, {
      exports: demo.exports, loader: false, hint: false, idPrefix: 'h-', ...hooks,
    });
    ribbon.open(demo.getEditor().save(), 'hosted.docx');
    (window as any).__hosted = ribbon;
  }, options);
}

const events = (page: Page) => page.evaluate(() => (window as any).__events as string[]);
const isDirty = (page: Page) => page.evaluate(() => (window as any).__hosted.isDirty as boolean);
const firstBlock = (page: Page) => page.locator('#hosted [data-anchor][contenteditable="true"]').first();

test.describe('ribbon for a host that keeps files', () => {
  test('typing, a command and a keyboard undo each make the document changed; markSaved and a new open make it unchanged', async ({ page }) => {
    await mountHosted(page, { dirty: true, onSave: true });
    expect(await isDirty(page)).toBe(false);
    expect(await events(page)).toEqual([]);

    // Typing counts at once: before the editor commits it (that waits for blur).
    await firstBlock(page).click();
    await page.keyboard.press('End');
    await page.keyboard.type('x');
    await expect.poll(() => events(page)).toEqual(['dirty true']);
    expect(await isDirty(page)).toBe(true);
    await page.evaluate(() => (window as any).__hosted.markSaved());
    expect(await isDirty(page)).toBe(false);
    expect(await events(page)).toEqual(['dirty true', 'dirty false']);

    // A ribbon command changes the session's version.
    await firstBlock(page).click();
    await page.keyboard.press('Home');
    await page.keyboard.press('Shift+End');
    await page.locator('#hosted button[data-cmd="bold"]').click();
    await expect.poll(() => isDirty(page)).toBe(true);
    await page.evaluate(() => (window as any).__hosted.markSaved());

    // Undo from the keyboard goes to the editor, not through a ribbon command, and still counts.
    await firstBlock(page).click();
    await page.keyboard.press('Control+z');
    await expect.poll(() => isDirty(page)).toBe(true);

    // Another document opens unchanged.
    await page.evaluate(() => { const r = (window as any).__hosted; r.open(r.save(), 'other.docx'); });
    expect(await isDirty(page)).toBe(false);
    expect((await events(page)).at(-1)).toBe('dirty false');
  });

  test('with onSave the document stays changed until the host marks it saved', async ({ page }) => {
    await mountHosted(page, { dirty: true, onSave: true });
    await firstBlock(page).click();
    await page.keyboard.press('End');
    await page.keyboard.type('y');
    await expect.poll(() => isDirty(page)).toBe(true);
    await page.locator('#hosted [data-dxr="save"]').click();
    expect(await events(page)).toContain('save hosted.docx true');
    expect(await isDirty(page)).toBe(true);    // the host has not stored it yet
  });

  test('Ctrl+S saves and Ctrl+O opens through the host', async ({ page }) => {
    await mountHosted(page, { onSave: true, onOpenRequest: true });
    let pickers = 0;
    page.on('filechooser', () => pickers++);
    await firstBlock(page).click();
    await page.keyboard.press('Control+s');
    await page.keyboard.press('Control+o');
    await expect.poll(() => events(page)).toEqual(['save hosted.docx true', 'open']);
    expect(pickers).toBe(0);
  });

  test('New and Open go to the host instead of replacing the document or showing the browser picker', async ({ page }) => {
    await mountHosted(page, { onNewRequest: true, onOpenRequest: true });
    let pickers = 0;
    page.on('filechooser', () => pickers++);
    const handle = await page.evaluate(() => (window as any).__hosted.editor.sessionHandle);
    await page.locator('#hosted [data-dxr="new"]').click();
    await page.locator('#hosted label:has([data-dxr="file"])').click();
    await page.locator('#hosted label:has([data-dxr="file"])').focus();
    await page.keyboard.press('Enter');
    expect(await events(page)).toEqual(['new', 'open', 'open']);
    expect(pickers).toBe(0);
    expect(await page.evaluate(() => (window as any).__hosted.editor.sessionHandle)).toBe(handle);
  });

  test('the Recent list shows the host entries only with onOpenRecent, and hands back the one picked', async ({ page }) => {
    const entries = [
      { id: 'C:\\docs\\a.docx', name: 'a.docx', title: 'C:\\docs\\a.docx' },
      { id: 'C:\\docs\\b.docx', name: 'b.docx' },
    ];
    await mountHosted(page, {});
    await page.evaluate((entries) => (window as any).__hosted.setRecent(entries), entries);
    await expect(page.locator('#hosted [data-dxr="recent"]')).toBeHidden();

    await mountHosted(page, { onOpenRecent: true });
    const recent = page.locator('#hosted [data-dxr="recent"]');
    await expect(recent).toBeHidden();
    await page.evaluate((entries) => (window as any).__hosted.setRecent(entries), entries);
    await expect(recent).toBeVisible();
    await expect(recent.locator('option')).toHaveText(['Recent…', 'a.docx', 'b.docx']);
    await expect(recent.locator('option').nth(1)).toHaveAttribute('title', 'C:\\docs\\a.docx');
    await recent.selectOption({ label: 'b.docx' });
    expect(await events(page)).toEqual(['recent C:\\docs\\b.docx']);
    await expect(recent).toHaveValue('');      // back to the placeholder: b can be picked again
    await recent.selectOption({ label: 'b.docx' });
    expect(await events(page)).toEqual(['recent C:\\docs\\b.docx', 'recent C:\\docs\\b.docx']);
    await page.evaluate(() => (window as any).__hosted.setRecent([]));
    await expect(recent).toBeHidden();
  });

  // (Last in the file: under Microsoft Edge, as DOCXODUS_CHROMIUM_PATH may point to, the browser
  // closes after a download and the next test cannot start; social-demo.spec.ts shows the same.)
  test('without onSave, Save (a browser download) marks the document saved', async ({ page }) => {
    await mountHosted(page, { dirty: true });
    await firstBlock(page).click();
    await page.keyboard.press('End');
    await page.keyboard.type('z');
    await expect.poll(() => isDirty(page)).toBe(true);
    const download = page.waitForEvent('download');
    await page.locator('#hosted [data-dxr="save"]').click();
    expect((await download).suggestedFilename()).toBe('hosted.docx');
    expect(await isDirty(page)).toBe(false);
  });
});
