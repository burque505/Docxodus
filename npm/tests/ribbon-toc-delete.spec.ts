import { test, expect, type Page } from '@playwright/test';

/**
 * Deleting a selection that reaches into a table of contents from outside: as in Word, the
 * selection takes the whole field and the whole table goes, the page renders on (a delete that
 * took the field's begin and left its end once broke every later render), and one undo brings
 * the table back.
 */

const HOST = '#td-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

/** "Before text", then a contents table, then the headings and text of `lines`. */
async function mount(page: Page, errors: string[], paginated: boolean) {
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'td-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('td.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  const lines = ['Before text', '1 Intro', 'Body one', '1 Detail', 'Body two', '1 Methods', 'Body three'];
  await page.locator(blocks).first().click();
  await page.evaluate((ls) => (window as any).__ribbon.exec('pasteText', { text: ls.map((l: string) => l.replace(/^\d /, '')).join('\n') }), lines);
  for (let i = 0; i < lines.length; i++) {
    if (!/^1 /.test(lines[i])) continue;
    await page.locator(blocks).nth(i).click();
    await page.evaluate(() => (window as any).__ribbon.exec('style', { styleId: 'Heading1' }));
  }
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
  // the table goes before "Intro"
  await page.locator(blocks).nth(1).click();
  await page.keyboard.press('Home');
  await page.locator(`${HOST} .dxr-tab[data-tab="references"]`).click();
  await page.locator(`${HOST} [data-dxr="toc"]`).click();
  await expect.poll(() => entries(page)).toHaveLength(3);
}

const entries = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#td-host [data-anchor][contenteditable="true"]'))
    .filter((b) => b.querySelector('a[href^="#_Toc"]'))
    .map((b) => (b.textContent ?? '').replace(/\s+/g, ' ').trim()));

const texts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#td-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/[ \s]+/g, ' ').trim()));

/** A point in the `n`th contents entry, `x` px in. */
const inEntry = (page: Page, n: number, x = 12) => page.evaluate(([n, x]) => {
  const e = Array.from(document.querySelectorAll<HTMLElement>('#td-host [data-anchor][contenteditable="true"]'))
    .filter((b) => b.querySelector('a[href^="#_Toc"]'))[n];
  const r = e.getBoundingClientRect();
  return { x: r.left + x, y: r.top + r.height / 2 };
}, [n, x] as [number, number]);

for (const paginated of [true, false]) {
  test(`a selection from the text before the contents into them takes the whole table (${paginated ? 'page' : 'continuous'} view)`, async ({ page }) => {
    const errors: string[] = [];
    await mount(page, errors, paginated);
    const before = await texts(page);
    // from inside "Before text" (after "Before") to inside the second entry
    const start = await page.evaluate(() => {
      const b = document.querySelectorAll<HTMLElement>('#td-host [data-anchor][contenteditable="true"]')[0];
      const r = b.getBoundingClientRect();
      return { x: r.left + 1, y: r.top + r.height / 2 };
    });
    await page.mouse.click(start.x, start.y);
    await page.keyboard.press('Home');
    for (let i = 0; i < 6; i++) await page.keyboard.press('ArrowRight');
    const end = await inEntry(page, 1);
    await page.keyboard.down('Shift');
    await page.mouse.click(end.x, end.y);
    await page.keyboard.up('Shift');
    await page.keyboard.press('Backspace');

    await expect.poll(() => entries(page)).toEqual([]);
    const after = await texts(page);
    // (the table's "Contents" title went too: it lay inside the selection)
    expect(before[1]).toBe('Contents');
    expect(after[0]).toBe('Before');
    expect(after.slice(1)).toEqual(before.slice(5));
    expect(errors).toEqual([]);

    // one undo: the text and the whole table back
    await page.keyboard.press('Control+z');
    await expect.poll(() => entries(page)).toHaveLength(3);
    expect(await texts(page)).toEqual(before);
    expect(errors).toEqual([]);
  });
}

test('a selection from inside the contents past them takes the whole table; what is left keeps its own style', async ({ page }) => {
  const errors: string[] = [];
  await mount(page, errors, true);
  const before = await texts(page);
  const start = await inEntry(page, 1);
  await page.mouse.click(start.x, start.y);
  // to just after "Int" in "Intro" (the heading after the table)
  const intro = await page.evaluate(() => {
    const all = Array.from(document.querySelectorAll<HTMLElement>('#td-host [data-anchor][contenteditable="true"]'));
    const h = all.filter((b) => !b.querySelector('a[href^="#_Toc"]'))[2];
    const walker = document.createTreeWalker(h, NodeFilter.SHOW_TEXT);
    const text = walker.nextNode() as Text;
    const r = document.createRange();
    r.setStart(text, 2); r.setEnd(text, 3);
    const box = r.getBoundingClientRect();
    return { x: box.right - 1, y: box.top + box.height / 2, heading: h.textContent };
  });
  expect(intro.heading).toBe('Intro');
  await page.keyboard.down('Shift');
  await page.mouse.click(intro.x, intro.y);
  await page.keyboard.up('Shift');
  await page.keyboard.press('Delete');

  await expect.poll(() => entries(page)).toEqual([]);
  const after = await texts(page);
  expect(after.slice(0, 4)).toEqual(['Before text', 'Contents', 'ro', 'Body one']);
  // ("ro" is still the heading: the contents entry's paragraph went whole)
  const tag = await page.evaluate(() => document.querySelectorAll<HTMLElement>('#td-host [data-anchor][contenteditable="true"]')[2].tagName);
  expect(tag).toBe('H1');
  expect(errors).toEqual([]);
  await page.keyboard.press('Control+z');
  await expect.poll(() => texts(page)).toEqual(before);
});

test('a selection inside the contents deletes only the entries it covers', async ({ page }) => {
  const errors: string[] = [];
  await mount(page, errors, true);
  const a = await inEntry(page, 0, 20), b = await inEntry(page, 1, 20);
  await page.mouse.click(a.x, a.y);
  await page.keyboard.down('Shift');
  await page.mouse.click(b.x, b.y);
  await page.keyboard.up('Shift');
  await page.keyboard.press('Delete');
  await expect.poll(async () => (await entries(page)).length).toBe(2);
  expect(errors).toEqual([]);
});
