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

// save() while the caret is still in the block typed in (Ctrl+S, a host's save on close, an
// autosave): a block commits its typing on blur, so the last words typed were missing from the
// saved document. save() now commits them first, as a step of its own, and leaves the caret.

async function open(page: Page, options: object) {
  await page.goto('/test-harness.html');
  await waitForDocxodus(page);
  await page.evaluate(({ bytes, options }) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'st-editor';
    document.body.appendChild(container);
    (window as any).__ed = D.DocxEditor.open(container, new Uint8Array(bytes), D, options);
    const editable = Array.from(container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'));
    // A short plain paragraph, so End reaches the end of the paragraph and not of a wrapped line.
    const target = editable.find((p) => p.tagName === 'P' && !p.closest('table')
      && !p.querySelector('[data-list-marker], a, img, [data-field]') && /\S/.test(p.textContent || '')
      && (p.textContent || '').trim().length < 40);
    (window as any).__stIndex = target ? editable.indexOf(target) : -1;
  }, { bytes: readTestFile('HC031-Complicated-Document.docx'), options });
  const index = await page.evaluate(() => (window as any).__stIndex as number);
  expect(index).toBeGreaterThanOrEqual(0);
  return page.locator('#st-editor [data-anchor][contenteditable="true"]').nth(index);
}

// The saved document's text, read back through a second editor.
const saved = (page: Page) => page.evaluate(() => {
  const D = (window as any).Docxodus;
  const probe = document.createElement('div');
  const reader = D.DocxEditor.open(probe, (window as any).__ed.save(), D, {});
  const text = probe.textContent || '';
  reader.close();
  return text;
});

for (const paginated of [false, true]) {
  for (const typingUndo of ['block', 'word'] as const) {
    test.describe(`DocxEditor save() with typing in the focused block (${paginated ? 'page view' : 'flowing'}, typingUndo "${typingUndo}")`, () => {
      test('the typing is saved, the caret stays, and Undo takes it back as one step', async ({ page }) => {
        const block = await open(page, { paginated, typingUndo });
        await block.click();
        await page.keyboard.press('End');
        await page.keyboard.type('stalpha');          // one word, so "word" has not committed it yet

        expect(await saved(page)).toContain('stalpha');
        // The block still has the focus, shows the typing once, and typing goes on where it was.
        expect(await page.evaluate(() => {
          const active = document.activeElement as HTMLElement | null;
          return !!active?.isContentEditable && (active.textContent || '').split('stalpha').length - 1;
        })).toBe(1);
        await page.keyboard.type('stbeta');
        expect(await saved(page)).toContain('stalphastbeta');

        // Each save made its own step: Undo takes back the second, then the first.
        await page.keyboard.press('Control+z');
        let text = await saved(page);
        expect(text).toContain('stalpha');
        expect(text).not.toContain('stbeta');
        await page.keyboard.press('Control+z');
        expect(await saved(page)).not.toContain('stalpha');
      });

      test('a save with nothing typed adds no step and keeps Redo', async ({ page }) => {
        const block = await open(page, { paginated, typingUndo });
        await block.click();
        await page.keyboard.press('End');
        await page.keyboard.type('stgamma');
        await page.keyboard.press('Control+z');
        expect(await saved(page)).not.toContain('stgamma');
        // That save found nothing to commit, so Redo still has the typing.
        await page.keyboard.press('Control+Shift+z');
        expect(await saved(page)).toContain('stgamma');
      });
    });
  }
}

test('commitTyping() with no block focused does nothing', async ({ page }) => {
  const block = await open(page, {});
  const before = await page.evaluate(() => (window as any).__ed.version);
  await block.click();
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
  await page.evaluate(() => (window as any).__ed.commitTyping());
  expect(await page.evaluate(() => (window as any).__ed.version)).toBe(before);
});
