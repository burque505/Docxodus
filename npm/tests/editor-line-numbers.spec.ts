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

import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Three columns of long paragraphs, then (a continuous section break) one column; both sections
 *  number their lines and restart each section. */
function columnsFixture(): Uint8Array {
  const lorem = 'Lorem ipsum dolor sit amet consectetuer adipiscing elit maecenas porttitor congue massa fusce posuere magna sed pulvinar ultricies purus lectus malesuada libero sit amet commodo magna eros quis urna.';
  const p = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const ln = '<w:lnNumType w:countBy="1" w:restart="newSection"/>';
  const body = [1, 2, 3, 4, 5, 6].map((i) => p(`${i}. ${lorem}`)).join('')
    + `<w:p><w:pPr><w:sectPr>${ln}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="3" w:space="720"/></w:sectPr></w:pPr></w:p>`
    + p('After the columns, one column again.') + p('A second line of the last section.')
    + `<w:sectPr><w:type w:val="continuous"/>${ln}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
  const rel = (id: string, kind: string, target: string) => `<Relationship Id="${id}" Type="${R_NS}/${kind}" Target="${target}"/>`;
  return storedZip([
    { name: '[Content_Types].xml', data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>') },
    { name: '_rels/.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`) },
  ]);
}

test('in columns each column is numbered on from the one before, and a new section restarts', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((bytes) => {
    const host = document.createElement('div');
    host.id = 'ln-cols';
    document.body.appendChild(host);
    const D = (window as any).Docxodus;
    (window as any).__ed = D.DocxEditor.open(host, new Uint8Array(bytes), D, { paginated: true });
  }, Array.from(columnsFixture()));
  await expect(page.locator('#ln-cols .docx-line-numbers span').first()).toBeVisible();
  const spans = await page.evaluate(() => Array.from(document.querySelectorAll('#ln-cols .docx-line-numbers span')).map((s) => {
    const r = s.getBoundingClientRect();
    return { n: Number(s.textContent), x: Math.round(r.right), y: Math.round(r.top) };
  }));
  // 6 paragraphs of 7 lines in three columns; the empty paragraph that carries the break takes no number.
  expect(spans.length).toBe(42 + 2);
  // No two numbers drawn on top of each other.
  const places = new Set(spans.map((s) => `${s.x},${s.y}`));
  expect(places.size).toBe(spans.length);
  // The last section's two lines restart at 1, and come last.
  const tail = spans.slice(-2).map((s) => s.n);
  expect(tail).toEqual([1, 2]);
  // In the column section, ordered by column (x) then top: 1, 2, 3, ... with no gap or repeat.
  const columns = spans.slice(0, -2).sort((a, b) => a.x - b.x || a.y - b.y).map((s) => s.n);
  expect(columns).toEqual(Array.from({ length: columns.length }, (_, i) => i + 1));
  expect(new Set(spans.slice(0, -2).map((s) => s.x)).size).toBe(3);
});
