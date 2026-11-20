import { test, expect, type Page } from '@playwright/test';

/**
 * Word's line numbers in the page view: a section's line numbering (w:lnNumType) is drawn in the
 * margin from the laid-out lines, every count-by-th number shown; a suppressed paragraph is neither
 * numbered nor counted; edits redraw them; the continuous view shows none (as Word's Print Layout).
 */

const HOST = '#ln-host';
const LONG = 'This paragraph is long enough that it wraps onto more than one line on the page, '
  + 'so it must take one number for each of its lines and not just one for the paragraph.';

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'ln-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('ln.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  const blocks = page.locator(`${HOST} [data-anchor][contenteditable="true"]`);
  await blocks.first().click();
  await page.keyboard.type('One');
  await page.keyboard.press('Enter');
  await page.keyboard.type(LONG);
  await page.keyboard.press('Enter');
  await page.keyboard.type('Three');
  await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
  await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
}

/** The drawn numbers, in page order. */
const numbers = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll('#ln-host .docx-line-numbers span')).map((s) => s.textContent));

/** How many lines each body paragraph takes on the page. */
const lineCounts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#ln-host .page-box [data-anchor][contenteditable="true"]')).map((p) => {
    const range = document.createRange();
    range.selectNodeContents(p);
    return new Set(Array.from(range.getClientRects()).filter((r) => r.height > 0).map((r) => Math.round(r.top))).size;
  }));

const setLineNumbering = (page: Page, lineNumbering: Record<string, unknown>) =>
  page.evaluate((ln) => {
    const r = (window as any).__ribbon;
    const first = document.querySelector<HTMLElement>('#ln-host .page-box [data-anchor][contenteditable="true"]')!;
    first.focus();
    (r.editor as any).activeBlock = first;
    return r.editor.setPageSetup({ lineNumbering: ln });
  }, lineNumbering);

test('each line of the body gets its number, in the margin left of the text', async ({ page }) => {
  await mount(page);
  expect(await setLineNumbering(page, { countBy: 1 })).toBe(true);
  const counts = await lineCounts(page);
  expect(counts[1]).toBeGreaterThan(1);
  const total = counts.reduce((a, b) => a + b, 0);
  await expect.poll(() => numbers(page)).toEqual(Array.from({ length: total }, (_, i) => String(i + 1)));
  const placed = await page.evaluate(() => {
    const first = document.querySelector('#ln-host .docx-line-numbers span')!.getBoundingClientRect();
    const text = document.querySelector('#ln-host .page-box [data-anchor][contenteditable="true"]')!.getBoundingClientRect();
    return { right: first.right, textLeft: text.left, numberMiddle: (first.top + first.bottom) / 2, textTop: text.top, textBottom: text.bottom };
  });
  expect(placed.right).toBeLessThan(placed.textLeft);
  expect(placed.numberMiddle).toBeGreaterThan(placed.textTop);
  expect(placed.numberMiddle).toBeLessThan(placed.textBottom);
});

test('count by 2 and start at 10 show 10, 12, ...; a suppressed paragraph is not counted', async ({ page }) => {
  await mount(page);
  expect(await setLineNumbering(page, { countBy: 2, start: 10 })).toBe(true);
  const counts = await lineCounts(page);
  const total = counts.reduce((a, b) => a + b, 0);
  const shown = Array.from({ length: total }, (_, i) => i + 10).filter((n) => n % 2 === 0).map(String);
  await expect.poll(() => numbers(page)).toEqual(shown);

  // Suppress the long paragraph: only One and Three are counted (10, 11): 10 shows.
  await page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).nth(1).click();
  await page.evaluate(() => (window as any).__ribbon.editor.setSuppressLineNumbers(true));
  await expect.poll(() => numbers(page)).toEqual(['10']);
});

test('typing a new line redraws the numbers; the continuous view shows none', async ({ page }) => {
  await mount(page);
  expect(await setLineNumbering(page, { countBy: 1 })).toBe(true);
  const before = (await numbers(page)).length;
  const last = page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).last();
  await last.click();
  await page.keyboard.press('End');
  await page.keyboard.press('Shift+Enter');
  await page.keyboard.type('and a second line');
  await expect.poll(async () => (await numbers(page)).length).toBe(before + 1);

  await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(false));
  await expect.poll(async () => (await numbers(page)).length).toBe(0);
});
