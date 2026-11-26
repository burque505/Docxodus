import { test, expect, Page } from '@playwright/test';

/**
 * Cell selection (cell-selection.ts; DocxEditor.cellSelection): Word's rectangle of whole cells,
 * made by dragging from one cell into another, Shift+click, or Shift+arrows past a cell's edge;
 * shaded, grown over merged cells, a nested table a table of its own; Escape or a click ends it.
 *
 * The table:   A1 | B1 | C1
 *              A2 | B2 | C2 (holds a nested table N1 | N2)
 *              ^^ | B3 | C3          (A2 spans two rows)
 *              A4 (two columns) | C4
 */

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

const TABLE = '<p>Before</p><table border="1">'
  + '<tr><td>A1</td><td>B1</td><td>C1</td></tr>'
  + '<tr><td rowspan="2">A2</td><td>B2</td><td>C2<table border="1"><tr><td>N1</td><td>N2</td></tr></table></td></tr>'
  + '<tr><td>B3</td><td>C3</td></tr>'
  + '<tr><td colspan="2">A4</td><td>C4</td></tr>'
  + '</table><p>After</p>';

async function openTable(page: Page) {
  await page.goto('/test-harness.html');
  await waitForDocxodus(page);
  await page.evaluate(async (html) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'ex-editor';
    document.body.appendChild(container);
    const ed = D.DocxEditor.openBlank(container, D, {});
    (window as any).__ed = ed;
    container.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.focus();
    await ed.exec('pasteHtml', { html, formatting: 'keep' });
  }, TABLE);
  await expect.poll(() => page.evaluate(() => document.querySelectorAll('#ex-editor table').length)).toBe(2);
}

/** The middle of the text `text` in a block (the innermost block holding exactly it). */
async function pointOf(page: Page, text: string) {
  const box = await page.evaluate((text) => {
    const blocks = Array.from(document.querySelectorAll<HTMLElement>('#ex-editor [data-anchor][contenteditable="true"]'));
    const block = blocks.find((b) => (b.textContent ?? '').trim() === text)!;
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    const node = walker.nextNode() as Text;
    const range = document.createRange();
    range.selectNodeContents(node);
    const r = range.getBoundingClientRect();
    return { x: r.left + r.width / 2, y: r.top + r.height / 2, right: r.right - 1 };
  }, text);
  return box;
}

async function drag(page: Page, from: string, to: string) {
  const a = await pointOf(page, from), b = await pointOf(page, to);
  await page.mouse.move(a.x, a.y);
  await page.mouse.down();
  await page.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2, { steps: 4 });
  await page.mouse.move(b.x, b.y, { steps: 4 });
  await page.mouse.up();
}

const selection = (page: Page) => page.evaluate(() => (window as any).__ed.cellSelection());
const shaded = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#ex-editor .docx-cell-selected'))
    .map((c) => Array.from(c.querySelectorAll('[data-anchor]')).map((b) => (b.textContent ?? '').trim())[0]));

test.describe('Cell selection', () => {
  test('a drag from one cell into another selects whole cells, grown over a merged cell', async ({ page }) => {
    await openTable(page);
    await drag(page, 'A1', 'B2');
    // A2 spans rows 2-3, so the rectangle takes row 3 as well.
    expect(await selection(page)).toEqual({ rows: [0, 2], columns: [0, 1], cells: 5, tableRows: 4, tableColumns: 3 });
    expect(await shaded(page)).toEqual(['A1', 'B1', 'A2', 'B2', 'B3']);
    // (the browser's own selection is collapsed, hidden)
    expect(await page.evaluate(() => window.getSelection()!.isCollapsed)).toBe(true);
    expect(await page.evaluate(() => document.querySelector('#ex-editor')!.hasAttribute('data-docx-cell-selection'))).toBe(true);
  });

  test('a merged cell across columns grows the rectangle sideways', async ({ page }) => {
    await openTable(page);
    await drag(page, 'B3', 'A4');
    expect(await selection(page)).toMatchObject({ rows: [1, 3], columns: [0, 1] });
  });

  test('a nested table is a table of its own; from it into the outer table takes its cell', async ({ page }) => {
    await openTable(page);
    await drag(page, 'N1', 'N2');
    expect(await selection(page)).toEqual({ rows: [0, 0], columns: [0, 1], cells: 2, tableRows: 1, tableColumns: 2 });
    await page.mouse.click(10, 10);
    await drag(page, 'N1', 'C3');
    expect(await selection(page)).toEqual({ rows: [1, 2], columns: [2, 2], cells: 2, tableRows: 4, tableColumns: 3 });
  });

  test('Shift+click selects from the caret\'s cell; a click ends it', async ({ page }) => {
    await openTable(page);
    const a = await pointOf(page, 'A1'), c = await pointOf(page, 'C4');
    await page.mouse.click(a.x, a.y);
    await page.keyboard.down('Shift');
    await page.mouse.click(c.x, c.y);
    await page.keyboard.up('Shift');
    expect(await selection(page)).toMatchObject({ rows: [0, 3], columns: [0, 2], cells: 10 });
    const b = await pointOf(page, 'B1');
    await page.mouse.click(b.x, b.y);
    expect(await selection(page)).toBeNull();
    expect(await shaded(page)).toEqual([]);
  });

  test('Shift+arrows past a cell\'s edge select cells, then move on the grid; Escape ends it', async ({ page }) => {
    await openTable(page);
    const b = await pointOf(page, 'B1');
    await page.mouse.click(b.right, b.y);
    await page.keyboard.press('End');
    await page.keyboard.press('Shift+ArrowRight');
    expect(await selection(page)).toMatchObject({ rows: [0, 0], columns: [1, 2], cells: 2 });
    await page.keyboard.press('Shift+ArrowDown');
    expect(await selection(page)).toMatchObject({ rows: [0, 1], columns: [1, 2], cells: 4 });
    await page.keyboard.press('Shift+ArrowLeft');
    expect(await selection(page)).toMatchObject({ rows: [0, 1], columns: [1, 1], cells: 2 });
    await page.keyboard.press('Escape');
    expect(await selection(page)).toBeNull();
  });

  test('a drag from a paragraph into the table stays a text selection', async ({ page }) => {
    await openTable(page);
    await drag(page, 'Before', 'B1');
    expect(await selection(page)).toBeNull();
    // and within one cell's text, too
    await page.mouse.click(10, 10);
    const a = await pointOf(page, 'C1');
    await page.mouse.move(a.x - 5, a.y);
    await page.mouse.down();
    await page.mouse.move(a.x + 5, a.y, { steps: 3 });
    await page.mouse.up();
    expect(await selection(page)).toBeNull();
  });

  test('exec selectCells: the cell, row, column or table at the caret, grown over merged cells', async ({ page }) => {
    await openTable(page);
    const b2 = await pointOf(page, 'B2');
    const exec = (part: string) => page.evaluate((part) => { (window as any).__ed.exec('selectCells', { part }); return (window as any).__ed.exec('cellSelection'); }, part);
    await page.mouse.click(b2.x, b2.y);
    expect(await exec('cell')).toMatchObject({ rows: [1, 1], columns: [1, 1], cells: 1 });
    await page.mouse.click(b2.x, b2.y);
    // (A2 spans rows 2-3: the row grows to both)
    expect(await exec('row')).toMatchObject({ rows: [1, 2], columns: [0, 2] });
    await page.mouse.click(b2.x, b2.y);
    // (A4 spans columns 1-2: the column grows to both)
    expect(await exec('column')).toMatchObject({ rows: [0, 3], columns: [0, 1] });
    await page.mouse.click(b2.x, b2.y);
    expect(await exec('table')).toMatchObject({ rows: [0, 3], columns: [0, 2], cells: 10 });
    // Around cells already selected: their rows.
    await page.mouse.click(10, 10);
    await drag(page, 'B1', 'C1');
    expect(await exec('row')).toMatchObject({ rows: [0, 0], columns: [0, 2], cells: 3 });
    // Outside a table: refused.
    const before = await pointOf(page, 'Before');
    await page.mouse.click(before.x, before.y);
    const refused = await page.evaluate(() => { try { (window as any).__ed.exec('selectCells', { part: 'table' }); return 'ran'; } catch (e) { return (e as any).code; } });
    expect(refused).toBe('refused');
  });

  test('the table handle shows at the table under the mouse (the innermost) and selects it', async ({ page }) => {
    await openTable(page);
    const handle = () => page.evaluate(() => {
      const h = document.querySelector<HTMLElement>('.docx-table-handle');
      if (!h || h.hidden) return null;
      const r = h.getBoundingClientRect();
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    });
    const b1 = await pointOf(page, 'B1');
    await page.mouse.move(b1.x, b1.y);
    await expect.poll(handle).not.toBeNull();
    const h = (await handle())!;
    // (on the way to the handle, outside the table, it stays)
    await page.mouse.move(h.x, h.y, { steps: 5 });
    await page.mouse.click(h.x, h.y);
    expect(await selection(page)).toMatchObject({ rows: [0, 3], columns: [0, 2], cells: 10, tableRows: 4 });
    // The nested table has a handle of its own.
    await page.mouse.click(10, 10);
    const n1 = await pointOf(page, 'N1');
    await page.mouse.move(n1.x, n1.y);
    await expect.poll(async () => (await handle())?.y).toBeGreaterThan(h.y + 20);
    const nh = (await handle())!;
    await page.mouse.move(nh.x, nh.y, { steps: 5 });
    await page.mouse.click(nh.x, nh.y);
    expect(await selection(page)).toEqual({ rows: [0, 0], columns: [0, 1], cells: 2, tableRows: 1, tableColumns: 2 });
    // Away from tables, with the caret outside them, it goes.
    const before = await pointOf(page, 'Before');
    await page.mouse.click(before.x, before.y);
    await page.mouse.move(before.x, before.y - 2);
    await expect.poll(handle, { timeout: 3000 }).toBeNull();
  });
});
