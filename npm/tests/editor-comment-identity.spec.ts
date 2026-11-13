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

// Word's form of w:date: UTC, to the second.
const WORD_DATE = /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$/;

// A comment posted from the editor carries the date and the initials Word writes, not only the
// author: w:date (when it was made) and w:initials (the balloon's initials). Replies too.
test.describe('DocxEditor comment identity', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
    await page.evaluate((bytes) => {
      const w = window as any;
      w.__bytes = new Uint8Array(bytes);
      // Open an editor, focus its first plain paragraph, and hand both back.
      w.__ciOpen = (options: object) => {
        const D = w.Docxodus;
        const container = document.createElement('div');
        document.body.appendChild(container);
        const editor = D.DocxEditor.open(container, w.__bytes, D, options);
        const block = Array.from(container.querySelectorAll<HTMLElement>('p[data-anchor][contenteditable="true"]'))
          .find((p) => (p.textContent || '').trim().length > 12)!;
        block.focus();
        return { container, editor };
      };
      // The comments the saved document holds, as the engine lists them.
      w.__ciSaved = (editor: any) => {
        const D = w.Docxodus;
        const h = D.DocxSessionBridge.OpenSession(editor.save(), '{}');
        const list = JSON.parse(D.DocxSessionBridge.ListComments(h));
        D.DocxSessionBridge.CloseSession(h);
        return list.map((c: any) => ({ author: c.author, initials: c.initials ?? null, date: c.date ?? null,
          text: c.text, reply: !!c.parentAnchorId }));
      };
    }, readTestFile('HC031-Complicated-Document.docx'));
  });

  test('a comment and its reply carry the author, the initials given and the time', async ({ page }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { editor } = w.__ciOpen({ commentAuthor: 'Ada Lovelace', commentInitials: 'AdL' });
      const before = Date.now();
      const created = editor.addComment('Please check.');
      editor.addCommentReply(created.anchorId, 'Checked.');
      const out = { comments: w.__ciSaved(editor), before, after: Date.now() };
      editor.close();
      return out;
    });
    const mine = res.comments.filter((c: any) => c.text === 'Please check.' || c.text === 'Checked.');
    expect(mine).toHaveLength(2);
    for (const c of mine) {
      expect(c.author).toBe('Ada Lovelace');
      expect(c.initials).toBe('AdL');
      expect(c.date).toMatch(WORD_DATE);
      const at = Date.parse(c.date);
      expect(at).toBeGreaterThanOrEqual(Math.floor(res.before / 1000) * 1000);
      expect(at).toBeLessThanOrEqual(res.after);
    }
    expect(mine.some((c: any) => c.reply)).toBe(true);
  });

  test('without initials given, they come from the author\'s name', async ({ page }) => {
    const comment = await page.evaluate(() => {
      const w = window as any;
      const { editor } = w.__ciOpen({ commentAuthor: 'Grace Brewster Hopper' });
      editor.addComment('Initials from the name.');
      const out = w.__ciSaved(editor).find((c: any) => c.text === 'Initials from the name.');
      editor.close();
      return out;
    });
    expect(comment.author).toBe('Grace Brewster Hopper');
    expect(comment.initials).toBe('GBH');
    expect(comment.date).toMatch(WORD_DATE);
  });

  test('setCommentAuthor changes who later comments are from, in the gutter too', async ({ page }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__ciOpen({ commentAuthor: 'Ada Lovelace', commentInitials: 'AdL' });
      editor.setCommentAuthor('Grace Hopper');
      editor.addComment('From Grace.');
      const draft = editor.beginComment();
      const bubbleAuthor = container.querySelector('.docx-comment-author')?.textContent ?? null;
      const out = { draft, bubbleAuthor, comment: w.__ciSaved(editor).find((c: any) => c.text === 'From Grace.') };
      editor.close();
      return out;
    });
    expect(res.comment.author).toBe('Grace Hopper');
    expect(res.comment.initials).toBe('GH');
    expect(res.draft).toBe(true);
    expect(res.bubbleAuthor).toBe('Grace Hopper');
  });
});

test('the ribbon passes the initials on, and its Author box changes the comments\' author', async ({ page }) => {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.click('#new');
  await page.waitForFunction(() => !!(window as any).__demo.getEditor());
  await page.addScriptTag({ url: '/editor.bundle.js' });
  const res = await page.evaluate((bytes) => {
    const { mountRibbon } = (window as any).DocxodusEditor;
    const demo = (window as any).__demo;
    const host = document.createElement('div');
    host.style.cssText = 'height:400px';
    document.body.appendChild(host);
    const ribbon = mountRibbon(host, {
      exports: demo.exports, loader: false, hint: false, documentName: 'identity.docx',
      commentAuthor: 'Ada Lovelace', commentInitials: 'AdL',
    });
    ribbon.open(new Uint8Array(bytes), 'identity.docx');
    const listed = () => {
      const bridge = demo.exports.DocxSessionBridge;
      const h = bridge.OpenSession(ribbon.editor.save(), '{}');
      const list = JSON.parse(bridge.ListComments(h));
      bridge.CloseSession(h);
      return list;
    };
    const focus = () => Array.from((ribbon.surface as HTMLElement).querySelectorAll<HTMLElement>('p[data-anchor][contenteditable="true"]'))
      .find((p) => (p.textContent || '').trim().length > 12)!.focus();
    focus();
    ribbon.editor.addComment('From the host\'s user.');
    const author = ribbon.control('author') as HTMLInputElement;
    const boxBefore = author.value;
    author.value = 'Grace Hopper';
    author.dispatchEvent(new Event('change'));
    focus();
    ribbon.editor.addComment('From the Author box.');
    const out = { boxBefore, list: listed().map((c: any) => ({ author: c.author, initials: c.initials, text: c.text })) };
    ribbon.destroy(); host.remove();
    return out;
  }, readTestFile('HC031-Complicated-Document.docx'));
  expect(res.boxBefore).toBe('Ada Lovelace');
  expect(res.list).toContainEqual({ author: 'Ada Lovelace', initials: 'AdL', text: 'From the host\'s user.' });
  expect(res.list).toContainEqual({ author: 'Grace Hopper', initials: 'GH', text: 'From the Author box.' });
});
