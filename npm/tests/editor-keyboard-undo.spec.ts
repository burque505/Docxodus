import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): number[] {
  return Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Undo and Redo from the keyboard behave as the ribbon's buttons do. Typing not yet committed (the
// block commits on blur) is one step of its own: Ctrl+Z undoes it first, not an older step, and
// the caret stays in the document so the next Ctrl+Z / Ctrl+Shift+Z / Ctrl+Y reaches the editor.
for (const paginated of [false, true]) {
  test.describe(`DocxEditor keyboard undo (${paginated ? 'page view' : 'flowing'})`, () => {
    test.beforeEach(async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);
      await page.evaluate(async ({ bytes, paginated }) => {
        const D = (window as any).Docxodus;
        const container = document.createElement('div');
        container.id = 'ku-editor';
        document.body.appendChild(container);
        const editor = D.DocxEditor.open(container, new Uint8Array(bytes), D, { paginated });
        (window as any).__ed = editor;
        const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
        const plain = Array.from(container.querySelectorAll<HTMLElement>('p[data-anchor][contenteditable="true"]'))
          .filter((p) => !p.closest('table') && !p.querySelector('[data-list-marker], a, img, [data-field]')
            && norm(p.textContent || '').length > 12);
        // A repaint replaces a block's node, so the blocks are found again by position.
        const editable = Array.from(container.querySelectorAll('[data-anchor][contenteditable="true"]'));
        (window as any).__kuIndex = plain.slice(0, 2).map((p) => editable.indexOf(p));
      }, { bytes: readTestFile('HC031-Complicated-Document.docx'), paginated });
    });

    const block = async (page: Page, i: number) => page.locator('#ku-editor [data-anchor][contenteditable="true"]')
      .nth(await page.evaluate((i) => (window as any).__kuIndex[i] as number, i));
    // The saved document's text: what the session holds, whatever the DOM shows.
    const saved = (page: Page) => page.evaluate(() => {
      const D = (window as any).Docxodus;
      const bytes: Uint8Array = (window as any).__ed.save();
      const probe = document.createElement('div');
      const reader = D.DocxEditor.open(probe, bytes, D, {});
      const text = probe.textContent || '';
      reader.close();
      return text;
    });
    async function typeAtEnd(page: Page, i: number, text: string) {
      await (await block(page, i)).click();
      await page.keyboard.press('End');
      await page.keyboard.type(text);
    }

    test('Ctrl+Z undoes the typing in progress, and Ctrl+Shift+Z and Ctrl+Y bring it back', async ({ page }) => {
      await typeAtEnd(page, 0, ' KUALPHA');
      await page.keyboard.press('Control+z');
      expect(await page.locator('#ku-editor').innerText()).not.toContain('KUALPHA');
      expect(await saved(page)).not.toContain('KUALPHA');

      await page.keyboard.press('Control+Shift+z');
      expect(await page.locator('#ku-editor').innerText()).toContain('KUALPHA');
      expect(await saved(page)).toContain('KUALPHA');

      await page.keyboard.press('Control+z');
      expect(await saved(page)).not.toContain('KUALPHA');
      await page.keyboard.press('Control+y');
      expect(await saved(page)).toContain('KUALPHA');
    });

    test('Undo steps through blocks newest first; Redo after leaving the block still works', async ({ page }) => {
      await typeAtEnd(page, 0, ' KUONE');
      await typeAtEnd(page, 1, ' KUTWO'); // clicking block 1 commits block 0's typing
      await page.keyboard.press('Control+z');
      let text = await saved(page);
      expect(text).toContain('KUONE');
      expect(text).not.toContain('KUTWO');
      await page.keyboard.press('Control+z');
      text = await saved(page);
      expect(text).not.toContain('KUONE');

      // Leaving the block (a blur with nothing typed) must not clear Redo.
      await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
      await (await block(page, 1)).click();
      await page.keyboard.press('Control+Shift+z');
      await page.keyboard.press('Control+Shift+z');
      text = await saved(page);
      expect(text).toContain('KUONE');
      expect(text).toContain('KUTWO');
    });

    test('a list item: Ctrl+Z takes back its typing and repaints it', async ({ page }) => {
      const index = await page.evaluate(() => {
        const editable = Array.from(document.querySelectorAll('#ku-editor [data-anchor][contenteditable="true"]'));
        return editable.findIndex((el) => el.querySelector(':scope > [data-list-marker]')
          && (el.textContent || '').trim().length > 12);
      });
      expect(index).toBeGreaterThanOrEqual(0);
      await page.locator('#ku-editor [data-anchor][contenteditable="true"]').nth(index).click();
      await page.keyboard.press('End');
      await page.keyboard.type(' KULIST');
      await page.keyboard.press('Control+z');
      expect(await page.locator('#ku-editor').innerText()).not.toContain('KULIST');
      expect(await saved(page)).not.toContain('KULIST');
      await page.keyboard.press('Control+Shift+z');
      expect(await saved(page)).toContain('KULIST');
    });

    test('a table: each Ctrl+Z takes back one step (second cell, first cell, the table)', async ({ page }) => {
      const tables = () => page.evaluate(() => {
        const D = (window as any).Docxodus;
        const probe = document.createElement('div');
        const reader = D.DocxEditor.open(probe, (window as any).__ed.save(), D, {});
        const count = probe.querySelectorAll('table').length;
        reader.close();
        return count;
      });
      const before = await tables();
      await (await block(page, 0)).click();
      await page.evaluate(() => {
        const ed = (window as any).__ed;
        const had = new Set(Array.from(document.querySelectorAll('#ku-editor table')));
        ed.insertTable(2, 2);
        const table = Array.from(document.querySelectorAll('#ku-editor table')).find((t) => !had.has(t))!;
        table.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.id = 'ku-cell';
      });
      expect(await tables()).toBe(before + 1);
      await page.locator('#ku-cell').click();
      await page.keyboard.type('KUCELLA');
      await page.keyboard.press('Tab');
      await page.keyboard.type('KUCELLB');

      await page.keyboard.press('Control+z');
      let text = await saved(page);
      expect(text).toContain('KUCELLA');
      expect(text).not.toContain('KUCELLB');
      await page.keyboard.press('Control+z');
      text = await saved(page);
      expect(text).not.toContain('KUCELLA');
      expect(await tables()).toBe(before + 1);
      await page.keyboard.press('Control+z');
      expect(await tables()).toBe(before);
    });
  });
}

// A header band is its own story, repainted by its own region rather than the body's diff.
test('a header band: Ctrl+Z takes back its typing, on screen and in the document', async ({ page }) => {
  await page.goto('/test-harness.html');
  await waitForDocxodus(page);
  await page.evaluate((bytes) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'ku-editor';
    document.body.appendChild(container);
    (window as any).__ed = D.DocxEditor.open(container, new Uint8Array(bytes), D, { headerFooter: true });
  }, readTestFile('DB003-Only-Default-Header.docx'));
  const header = page.locator('#ku-editor [data-hf-band="header"] [data-anchor][contenteditable="true"]').first();
  await header.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' KUHDR');
  await page.keyboard.press('Control+z');
  const savedHeader = () => page.evaluate(() => {
    const D = (window as any).Docxodus;
    const probe = document.createElement('div');
    const reader = D.DocxEditor.open(probe, (window as any).__ed.save(), D, { headerFooter: true });
    const text = probe.querySelector('[data-hf-band="header"]')?.textContent || '';
    reader.close();
    return text;
  });
  expect(await savedHeader()).not.toContain('KUHDR');
  expect(await page.locator('#ku-editor [data-hf-band="header"]').innerText()).not.toContain('KUHDR');
  await page.keyboard.press('Control+Shift+z');
  expect(await savedHeader()).toContain('KUHDR');
});
