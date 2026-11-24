import { test, expect, type Page } from '@playwright/test';

/**
 * Up and Down as in Word: inside a paragraph the browser's own line step, at its first or last
 * line on into the paragraph before or after, near the same place across the line; Ctrl+Down to
 * the next paragraph's start, Ctrl+Up to this one's start (the one before when already there).
 */

const HOST = '#lk-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;
const LONG = 'This paragraph is long enough to wrap onto several lines in the editor, so that the caret can '
  + 'move between its lines with the arrow keys before it leaves for the paragraphs around it; '
  + 'that is the whole point of having it here, and so it goes on for a while longer still.';

async function mount(page: Page, paginated: boolean) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'lk-host';
    host.style.height = '700px';
    host.style.width = '900px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('lk.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  await page.evaluate((t) => (window as any).__ribbon.exec('pasteText', { text: t }), ['First short one', LONG, 'Third short one'].join('\n'));
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
}

/** The index of the block the caret is in, and the caret's offset in its text. */
const caret = (page: Page) => page.evaluate(() => {
  const sel = getSelection()!;
  const node = sel.focusNode;
  const el = (node && (node.nodeType === 1 ? node as Element : node.parentElement))?.closest('[data-anchor][contenteditable="true"]') as HTMLElement | null;
  const all = Array.from(document.querySelectorAll('#lk-host [data-anchor][contenteditable="true"]'));
  if (!el || !node) return { block: -1, offset: -1 };
  const r = document.createRange();
  r.selectNodeContents(el);
  r.setEnd(node, sel.focusOffset);
  return { block: all.indexOf(el), offset: r.toString().length };
});

/** Click `x` px into the `n`th block's first (or last) line. */
async function clickLine(page: Page, n: number, x: number, last = false) {
  const p = await page.evaluate(([n, x, last]) => {
    const b = document.querySelectorAll<HTMLElement>('#lk-host [data-anchor][contenteditable="true"]')[n as number];
    const r = b.getBoundingClientRect();
    return { x: r.left + (x as number), y: last ? r.bottom - 6 : r.top + 6 };
  }, [n, x, last] as [number, number, boolean]);
  await page.mouse.click(p.x, p.y);
}

for (const paginated of [false, true]) {
  test(`Up and Down cross paragraphs and step through a long one's lines (${paginated ? 'page' : 'continuous'} view)`, async ({ page }) => {
    await mount(page, paginated);
    // (the long paragraph wraps: several times the first one's height)
    const heights = await page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#lk-host [data-anchor][contenteditable="true"]'))
      .slice(0, 2).map((b) => b.getBoundingClientRect().height));
    expect(heights[1]).toBeGreaterThan(heights[0] * 1.8);

    await clickLine(page, 0, 40);
    await page.keyboard.press('ArrowDown');
    expect((await caret(page)).block).toBe(1);
    // down through the long paragraph's lines, then out into the third
    let seen = 1;
    for (let i = 0; i < 12 && seen === 1; i++) {
      await page.keyboard.press('ArrowDown');
      seen = (await caret(page)).block;
    }
    expect(seen).toBe(2);
    // and back up: into the long paragraph's last line, then on up to the first
    await page.keyboard.press('ArrowUp');
    const inLong = await caret(page);
    expect(inLong.block).toBe(1);
    expect(inLong.offset).toBeGreaterThan(LONG.length / 2);
    for (let i = 0; i < 12 && seen !== 0; i++) {
      await page.keyboard.press('ArrowUp');
      seen = (await caret(page)).block;
    }
    expect(seen).toBe(0);
    // Up on the first line of the story: its start; typing still goes in there
    await page.keyboard.press('ArrowUp');
    expect(await caret(page)).toEqual({ block: 0, offset: 0 });
  });
}

test('the caret keeps its place across the line when it moves into the next paragraph', async ({ page }) => {
  await mount(page, false);
  await clickLine(page, 1, 120, true);
  const from = await caret(page);
  await page.keyboard.press('ArrowDown');
  const to = await caret(page);
  expect(to.block).toBe(2);
  // ("Third short one" is shorter than 120 px from the left only on narrow fonts: near the end or in it)
  expect(to.offset).toBeGreaterThan(3);
  expect(from.block).toBe(1);
});

test('Ctrl+Down goes to the next paragraph\'s start; Ctrl+Up to this one\'s, then the one before', async ({ page }) => {
  await mount(page, false);
  await clickLine(page, 1, 200);
  await page.keyboard.press('Control+ArrowDown');
  expect(await caret(page)).toEqual({ block: 2, offset: 0 });
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Control+ArrowUp');
  expect(await caret(page)).toEqual({ block: 2, offset: 0 });
  await page.keyboard.press('Control+ArrowUp');
  expect(await caret(page)).toEqual({ block: 1, offset: 0 });
});

test('typing after Down goes into the paragraph the caret moved to', async ({ page }) => {
  await mount(page, false);
  await clickLine(page, 0, 4);
  await page.keyboard.press('Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('Home');
  await page.keyboard.type('X');
  await expect.poll(() => page.evaluate(() =>
    (document.querySelectorAll<HTMLElement>('#lk-host [data-anchor][contenteditable="true"]')[1].textContent ?? '').slice(0, 5))).toBe('XThis');
});
