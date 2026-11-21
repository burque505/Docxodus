import { test, expect, type Page } from '@playwright/test';

/**
 * View > Navigation Pane: Word's headings tree beside the document — a click goes to a heading,
 * the caret's heading is marked, a drag moves a heading with its text (one undo), and the
 * right-click menu (Promote, Demote, New Heading Before/After, New Subheading, Delete, Select
 * Heading and Content, Expand/Collapse All, Show Heading Levels).
 */

const HOST = '#np-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;
const rows = `${HOST} [data-dxr="navpane"] .dxr-navitem`;

/** A document of `lines`: "1 Title" is a Heading 1, "2 Title" a Heading 2, anything else body text. */
async function mount(page: Page, lines: string[], paginated = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'np-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('np.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate(async (ls) => {
    const r = (window as any).__ribbon;
    r.exec('pasteText', { text: ls.map((l: string) => l.replace(/^\d /, '')).join('\n') });
  }, lines);
  for (let i = 0; i < lines.length; i++) {
    const m = /^(\d) /.exec(lines[i]);
    if (!m) continue;
    await page.locator(blocks).nth(i).click();
    await page.evaluate((lvl) => (window as any).__ribbon.exec('style', { styleId: `Heading${lvl}` }), m[1]);
  }
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
  await page.locator(`${HOST} .dxr-tab[data-tab="view"]`).click();
  await page.locator(`${HOST} [data-dxr="shownav"]`).check();
  await expect(page.locator(`${HOST} [data-dxr="navpane"]`)).toBeVisible();
}

const paneTexts = (page: Page) => page.locator(`${rows} .dxr-navtext`).allTextContents();
const docTexts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#np-host [data-anchor][contenteditable="true"]')).map((b) => (b.textContent ?? '').replace(/[\s ]+$/, '')));
const outline = (page: Page) => page.evaluate(() => (window as any).__ribbon.exec('outline')
);
const row = (page: Page, text: string) => page.locator(rows).filter({ hasText: text }).first();

async function menu(page: Page, text: string, item: string) {
  await row(page, text).click({ button: 'right' });
  await expect(page.locator(`${HOST} [data-dxr="navmenu"]`)).toBeVisible();
  await page.locator(`${HOST} [data-dxr="navmenu"] [data-navcmd="${item}"]`).click();
}

const DOC = ['1 Alpha', 'alpha text', '2 Alpha sub', 'sub text', '1 Beta', 'beta text', '1 Gamma', 'gamma text'];

test('the pane lists the headings by level; a click goes there and marks the caret\'s heading', async ({ page }) => {
  await mount(page, DOC);
  await expect.poll(() => paneTexts(page)).toEqual(['Alpha', 'Alpha sub', 'Beta', 'Gamma']);
  expect(await outline(page)).toEqual([
    { level: 1, text: 'Alpha' }, { level: 2, text: 'Alpha sub' }, { level: 1, text: 'Beta' }, { level: 1, text: 'Gamma' }]);
  await row(page, 'Gamma').click();
  await expect(row(page, 'Gamma')).toHaveAttribute('data-current', '');
  await page.keyboard.type('X');
  await expect.poll(() => docTexts(page)).toContain('XGamma');
  await page.locator(blocks).nth(3).click();    // "sub text"
  await expect(row(page, 'Alpha sub')).toHaveAttribute('data-current', '');
});

test('collapse hides subheadings in the pane only; Show Heading Levels and Expand All', async ({ page }) => {
  await mount(page, DOC);
  await row(page, 'Alpha').locator('.dxr-navtoggle').click();
  await expect.poll(() => paneTexts(page)).toEqual(['Alpha', 'Beta', 'Gamma']);
  expect((await docTexts(page)).length).toBe(8);
  await menu(page, 'Beta', 'expand');
  await expect.poll(() => paneTexts(page)).toEqual(['Alpha', 'Alpha sub', 'Beta', 'Gamma']);
  await row(page, 'Beta').click({ button: 'right' });
  await page.locator(`${HOST} [data-navcmd="levels"]`).click();
  await page.locator(`${HOST} [data-navlevel="1"]`).click();
  await expect.poll(() => paneTexts(page)).toEqual(['Alpha', 'Beta', 'Gamma']);
});

test('dragging a heading moves it with its text and subheadings, as one undo step', async ({ page }) => {
  await mount(page, DOC);
  await row(page, 'Alpha').dragTo(row(page, 'Gamma'), { targetPosition: { x: 20, y: 2 } });
  await expect.poll(() => docTexts(page)).toEqual(['Beta', 'beta text', 'Alpha', 'alpha text', 'Alpha sub', 'sub text', 'Gamma', 'gamma text']);
  await expect.poll(() => paneTexts(page)).toEqual(['Beta', 'Alpha', 'Alpha sub', 'Gamma']);
  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+z');
  await expect.poll(() => docTexts(page)).toEqual(DOC.map((l) => l.replace(/^\d /, '')));
  // to the end, below the last heading
  await row(page, 'Beta').dragTo(row(page, 'Gamma'), { targetPosition: { x: 20, y: 14 } });
  await expect.poll(() => docTexts(page)).toEqual(['Alpha', 'alpha text', 'Alpha sub', 'sub text', 'Gamma', 'gamma text', 'Beta', 'beta text']);
});

test('Promote, Demote, New Heading Before/After, New Subheading', async ({ page }) => {
  await mount(page, DOC);
  await menu(page, 'Alpha sub', 'promote');
  await expect.poll(() => outline(page)).toContainEqual({ level: 1, text: 'Alpha sub' });
  await menu(page, 'Beta', 'demote');
  await expect.poll(() => outline(page)).toContainEqual({ level: 2, text: 'Beta' });
  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+z');
  await expect.poll(() => outline(page)).toContainEqual({ level: 1, text: 'Beta' });

  await menu(page, 'Gamma', 'sub');
  await page.keyboard.type('Gamma child');
  await expect.poll(() => docTexts(page)).toEqual([...DOC.slice(0, 2).map((l) => l.replace(/^\d /, '')), 'Alpha sub', 'sub text', 'Beta', 'beta text', 'Gamma', 'gamma text', 'Gamma child']);
  await page.locator(blocks).first().click();    // (typing reaches the engine when the caret leaves)
  await expect.poll(() => outline(page)).toContainEqual({ level: 2, text: 'Gamma child' });
  await menu(page, 'Beta', 'before');
  await page.keyboard.type('Before beta');
  await page.locator(blocks).first().click();
  await expect.poll(() => outline(page)).toContainEqual({ level: 1, text: 'Before beta' });
  const texts = await docTexts(page);
  expect(texts.indexOf('Before beta')).toBe(texts.indexOf('Beta') - 1);
});

test('Delete takes the heading and its text; Select Heading and Content selects it', async ({ page }) => {
  await mount(page, DOC);
  await menu(page, 'Alpha', 'select');
  // (the range: Chromium's Selection.toString stops at the focused paragraph)
  await expect.poll(() => page.evaluate(() => window.getSelection()?.getRangeAt(0).toString())).toBe('Alphaalpha textAlpha subsub text');
  await menu(page, 'Beta', 'delete');
  await expect.poll(() => docTexts(page)).toEqual(['Alpha', 'alpha text', 'Alpha sub', 'sub text', 'Gamma', 'gamma text']);
  await page.locator(blocks).first().click();
  await page.keyboard.press('Control+z');
  await expect.poll(() => docTexts(page)).toEqual(DOC.map((l) => l.replace(/^\d /, '')));
  // the last heading: the body keeps one (empty) paragraph
  await menu(page, 'Gamma', 'delete');
  await expect.poll(() => paneTexts(page)).toEqual(['Alpha', 'Alpha sub', 'Beta']);
});

test('the pane works in page view, and closes from its × and the View box', async ({ page }) => {
  await mount(page, DOC, true);
  await expect.poll(() => paneTexts(page)).toEqual(['Alpha', 'Alpha sub', 'Beta', 'Gamma']);
  await row(page, 'Beta').click();
  await expect(row(page, 'Beta')).toHaveAttribute('data-current', '');
  await page.locator(`${HOST} [data-dxr="navpane"] [data-nav="close"]`).click();
  await expect(page.locator(`${HOST} [data-dxr="navpane"]`)).toBeHidden();
  await expect(page.locator(`${HOST} [data-dxr="shownav"]`)).not.toBeChecked();
});

test('a document without headings says how to get them', async ({ page }) => {
  await mount(page, ['just text', 'more text']);
  await expect(page.locator(`${HOST} [data-nav="empty"]`)).toBeVisible();
});
