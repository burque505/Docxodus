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

async function open(page: Page, options: object) {
  await page.goto('/test-harness.html');
  await waitForDocxodus(page);
  await page.evaluate(({ bytes, options }) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'tu-editor';
    document.body.appendChild(container);
    (window as any).__ed = D.DocxEditor.open(container, new Uint8Array(bytes), D, options);
    const editable = Array.from(container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'));
    // A short plain paragraph, so End reaches the end of the paragraph and not of a wrapped line.
    const target = editable.find((p) => p.tagName === 'P' && !p.closest('table')
      && !p.querySelector('[data-list-marker], a, img, [data-field]') && /\S/.test(p.textContent || '')
      && (p.textContent || '').trim().length < 40);
    (window as any).__tuIndex = target ? editable.indexOf(target) : -1;
  }, { bytes: readTestFile('HC031-Complicated-Document.docx'), options });
  const index = await page.evaluate(() => (window as any).__tuIndex as number);
  expect(index).toBeGreaterThanOrEqual(0);
  return page.locator('#tu-editor [data-anchor][contenteditable="true"]').nth(index);
}

// The saved document's text: what the session holds, whatever the DOM shows.
const saved = (page: Page) => page.evaluate(() => {
  const D = (window as any).Docxodus;
  const probe = document.createElement('div');
  const reader = D.DocxEditor.open(probe, (window as any).__ed.save(), D, {});
  const text = probe.textContent || '';
  reader.close();
  return text;
});

for (const paginated of [false, true]) {
  test.describe(`DocxEditor typingUndo (${paginated ? 'page view' : 'flowing'})`, () => {
    test('"word": Undo takes typing back a word at a time, and Redo puts it back', async ({ page }) => {
      const block = await open(page, { paginated, typingUndo: 'word' });
      await block.click();
      await page.keyboard.press('End');
      await page.keyboard.type(' tualpha tubeta tugamma');

      await page.keyboard.press('Control+z');
      let text = await saved(page);
      expect(text).toContain('tualpha tubeta');
      expect(text).not.toContain('tugamma');
      await page.keyboard.press('Control+z');
      text = await saved(page);
      expect(text).toContain('tualpha');
      expect(text).not.toContain('tubeta');
      await page.keyboard.press('Control+z');
      expect(await saved(page)).not.toContain('tualpha');

      for (let i = 0; i < 3; i++) await page.keyboard.press('Control+Shift+z');
      text = await saved(page);
      expect(text).toContain('tualpha tubeta tugamma');
      // The spaces typed between the words reach the document as ordinary spaces.
      expect(text).not.toContain(' tubeta');
      expect(text).not.toContain(' tugamma');
    });

    test('"word": a pause ends a step too', async ({ page }) => {
      const block = await open(page, { paginated, typingUndo: 'word' });
      await block.click();
      await page.keyboard.press('End');
      await page.keyboard.type(' tudelta');
      await page.waitForTimeout(1300);
      await page.keyboard.type('tuepsilon');
      await page.keyboard.press('Control+z');
      const text = await saved(page);
      expect(text).toContain('tudelta');
      expect(text).not.toContain('tuepsilon');
    });

    test('"block" (the default): the block\'s typing is one step', async ({ page }) => {
      const block = await open(page, { paginated });
      await block.click();
      await page.keyboard.press('End');
      await page.keyboard.type(' tualpha tubeta');
      await page.keyboard.press('Control+z');
      expect(await saved(page)).not.toContain('tualpha');
    });
  });
}
