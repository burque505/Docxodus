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

for (const paginated of [false, true]) {
  test(`Shift+Down over many paragraphs keeps the selection's moving end in view (${paginated ? 'page' : 'continuous'} view)`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
    await page.evaluate(() => {
      const host = document.createElement('div');
      host.id = 'lk-host';
      host.style.height = '420px';
      document.body.appendChild(host);
      const r = (window as any).DocxodusEditor.mountRibbon(host, {});
      r.setExports((window as any).Docxodus);
      r.openBlank('lk.docx');
      r.loader.done();
      (window as any).__ribbon = r;
    });
    await page.locator(blocks).first().click();
    const lines = Array.from({ length: 80 }, (_, i) => `Paragraph ${i + 1} of a long document`);
    await page.evaluate((t) => (window as any).__ribbon.exec('pasteText', { text: t }), lines.join('\n'));
    if (paginated) {
      await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
      await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
    }
    // (switching the view takes the focus: back into the first paragraph)
    await page.locator(blocks).first().click();
    await page.keyboard.press('Control+Home');
    /** Whether the selection's moving end is inside the visible part of the scrolling area. */
    const focusInView = () => page.evaluate(() => {
      const sel = getSelection()!;
      const r = document.createRange();
      r.setStart(sel.focusNode!, sel.focusOffset);
      const rect = r.getClientRects()[0] ?? (sel.focusNode!.nodeType === 1 ? (sel.focusNode as Element) : sel.focusNode!.parentElement!).getBoundingClientRect();
      let el: HTMLElement | null = document.querySelector('#lk-host [data-anchor]');
      while (el && !(el.scrollHeight > el.clientHeight && /auto|scroll/.test(getComputedStyle(el).overflowY))) el = el.parentElement;
      const box = el ? el.getBoundingClientRect() : { top: 0, bottom: innerHeight };
      return rect.top >= box.top - 1 && rect.bottom <= box.bottom + 1;
    });
    for (let i = 0; i < 50; i++) await page.keyboard.press('Shift+ArrowDown');
    const reached = await caret(page);
    expect(reached.block).toBeGreaterThan(40);
    expect(await focusInView()).toBe(true);
    for (let i = 0; i < 30; i++) await page.keyboard.press('Shift+ArrowUp');
    expect((await caret(page)).block).toBeLessThan(reached.block - 20);
    expect(await focusInView()).toBe(true);
    // Ctrl+Shift+End: the story's end, in view
    await page.keyboard.press('Control+Shift+End');
    expect((await caret(page)).block).toBe(79);
    expect(await focusInView()).toBe(true);
  });
}

for (const paginated of [false, true]) {
  test(`PgDn / PgUp move a screenful; with Shift they select a screenful at a time (${paginated ? 'page' : 'continuous'} view)`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
    await page.evaluate(() => {
      const host = document.createElement('div');
      host.id = 'lk-host';
      host.style.height = '420px';
      document.body.appendChild(host);
      const r = (window as any).DocxodusEditor.mountRibbon(host, {});
      r.setExports((window as any).Docxodus);
      r.openBlank('lk.docx');
      r.loader.done();
      (window as any).__ribbon = r;
    });
    await page.locator(blocks).first().click();
    const lines = Array.from({ length: 80 }, (_, i) => `Paragraph ${i + 1} of a long document`);
    await page.evaluate((t) => (window as any).__ribbon.exec('pasteText', { text: t }), lines.join('\n'));
    if (paginated) {
      await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
      await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
    }
    await page.locator(blocks).first().click();
    await page.keyboard.press('Control+Home');
    const inView = () => page.evaluate(() => {
      const sel = getSelection()!;
      const node = sel.focusNode!;
      const rect = (node.nodeType === 1 ? node as Element : node.parentElement!).getBoundingClientRect();
      let el: HTMLElement | null = document.querySelector('#lk-host [data-anchor]');
      while (el && !(el.scrollHeight > el.clientHeight && /auto|scroll/.test(getComputedStyle(el).overflowY))) el = el.parentElement;
      const box = el ? el.getBoundingClientRect() : { top: 0, bottom: innerHeight };
      return rect.bottom >= box.top && rect.top <= box.bottom;
    });
    const collapsed = () => page.evaluate(() => getSelection()!.isCollapsed);

    // PgDn: the caret goes a screenful down, in view, the selection still a caret
    await page.keyboard.press('PageDown');
    const one = await caret(page);
    expect(one.block).toBeGreaterThan(3);
    expect(await inView()).toBe(true);
    expect(await collapsed()).toBe(true);
    // a second PgDn goes further
    await page.keyboard.press('PageDown');
    const two = await caret(page);
    expect(two.block).toBeGreaterThan(one.block);
    // PgUp comes back near the first place
    await page.keyboard.press('PageUp');
    const back = await caret(page);
    expect(back.block).toBeLessThan(two.block);
    expect(Math.abs(back.block - one.block)).toBeLessThanOrEqual(2);

    // Shift+PgDn from the start selects a screenful at a time; the moving end stays in view
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('Shift+PageDown');
    const s1 = await caret(page);
    await page.keyboard.press('Shift+PageDown');
    const s2 = await caret(page);
    expect(s1.block).toBeGreaterThan(3);
    expect(s2.block).toBeGreaterThan(s1.block);
    expect(await inView()).toBe(true);
    expect(await collapsed()).toBe(false);
    // the selection runs from the story's start
    const startsAtTop = await page.evaluate(() => {
      const sel = getSelection()!;
      const n = sel.anchorNode!;
      const el = (n.nodeType === 1 ? n as Element : n.parentElement!).closest('[data-anchor][contenteditable="true"]');
      return el === document.querySelectorAll('#lk-host [data-anchor][contenteditable="true"]')[0] && sel.anchorOffset === 0;
    });
    expect(startsAtTop).toBe(true);
    // Shift+PgUp takes back a screenful
    await page.keyboard.press('Shift+PageUp');
    expect((await caret(page)).block).toBeLessThan(s2.block);

    // PgDn all the way: the story's end
    await page.keyboard.press('Control+Home');
    for (let i = 0; i < 40; i++) await page.keyboard.press('PageDown');
    expect((await caret(page)).block).toBe(79);
  });
}
