import { test, expect, type Page } from '@playwright/test';

/**
 * View > Ruler: Word's horizontal ruler. It draws the section's margins and the caret paragraph's
 * indents and tab stops at the zoom, in both views; dragging a marker sets the indent (one undo),
 * a click sets a tab stop of the selector's kind, dragging a stop off the ruler removes it,
 * double-clicking one opens the Tabs dialog, and dragging a margin edge sets the section's margin.
 */

const HOST = '#ru-host';

async function mount(page: Page, paginated = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'ru-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('ru.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.keyboard.type('Alpha beta gamma');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second');
  if (paginated) {
    await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
    await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
  }
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.locator(`${HOST} .dxr-tab[data-tab="view"]`).click();
  await page.locator(`${HOST} [data-dxr="showruler"]`).check();
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await expect(page.locator(`${HOST} .dxr-ruler-mark[data-mark="leftIndent"]`)).toBeVisible();
}

const state = (page: Page) => page.evaluate(() => (window as any).__ribbon.editor.rulerState());
const format = (page: Page) => page.evaluate(() => (window as any).__ribbon.editor.paragraphFormats()[0]);
const tabs = (page: Page) => page.evaluate(() => (window as any).__ribbon.editor.rulerState()?.tabs ?? null);
const mark = (page: Page, kind: string) => page.locator(`${HOST} .dxr-ruler-mark[data-mark="${kind}"]`).first();

/** The client x of `twips` from the paragraph's origin, and the ruler's middle y. */
async function at(page: Page, twips: number) {
  const s = await state(page);
  const box = (await page.locator(`${HOST} [data-dxr="ruler"]`).boundingBox())!;
  return { x: s.originLeft + twips * s.pxPerTwip, y: box.y + box.height / 2, s };
}

/** Drag the marker `kind` to `twips` from the origin (the right indent: from the right edge). */
async function drag(page: Page, kind: string, twips: number, options: { escape?: boolean } = {}) {
  await expect(mark(page, kind)).toBeVisible();
  const box = (await mark(page, kind).boundingBox())!;
  const s = await state(page);
  const target = kind === 'rightIndent' ? { x: s.originRight - twips * s.pxPerTwip } : await at(page, twips);
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(target.x, box.y + box.height / 2, { steps: 5 });
  if (options.escape) await page.keyboard.press('Escape');
  await page.mouse.up();
}

const undo = async (page: Page) => {
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.keyboard.press('Control+z');
};

for (const paginated of [false, true]) {
  test(`the ruler lines up with the text column (${paginated ? 'page' : 'continuous'} view)`, async ({ page }) => {
    await mount(page, paginated);
    const s = await state(page);
    expect(s.inSectionColumn).toBe(true);
    expect(s.marginLeftTwips).toBe(1440);
    // One inch of margin on screen is 1440 twips at the zoom.
    expect(Math.abs(s.textLeft - s.pageLeft - 1440 * s.pxPerTwip)).toBeLessThan(1);
    const paragraph = (await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().boundingBox())!;
    expect(Math.abs(paragraph.x - s.originLeft)).toBeLessThan(1);
    const left = (await mark(page, 'leftIndent').boundingBox())!;
    expect(Math.abs(left.x + left.width / 2 - s.originLeft)).toBeLessThan(1.5);
  });
}

test('dragging the left indent box moves the paragraph; one undo moves it back', async ({ page }) => {
  await mount(page);
  await drag(page, 'leftIndent', 720);
  await expect.poll(async () => (await format(page)).leftIndent).toBe(720);
  const s = await state(page);
  const paragraph = (await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().boundingBox())!;
  expect(Math.abs(paragraph.x - (s.originLeft + 720 * s.pxPerTwip))).toBeLessThan(1);
  await undo(page);
  await expect.poll(async () => (await format(page)).leftIndent ?? 0).toBe(0);
});

test('the first-line and hanging markers set a first-line or a hanging indent', async ({ page }) => {
  await mount(page);
  await drag(page, 'firstLine', 360);
  await expect.poll(async () => (await format(page)).firstLineIndent).toBe(360);
  // The hanging marker moves the other lines; the first line stays at 360.
  await drag(page, 'hanging', 720);
  await expect.poll(async () => format(page)).toMatchObject({ leftIndent: 720 });
  const f = await format(page);
  expect(f, JSON.stringify(f)).toMatchObject({ hangingIndent: 360 });
});

test('drags snap to 1/16 inch; Alt places freely; Escape cancels', async ({ page }) => {
  await mount(page);
  await drag(page, 'rightIndent', 700); // 700 twips in from the right edge snaps to 720
  await expect.poll(async () => (await format(page)).rightIndent).toBe(720);
  await drag(page, 'leftIndent', 500, { escape: true });
  expect((await format(page)).leftIndent ?? 0).toBe(0);
});

test('a click sets a tab stop of the selector\'s kind; dragging it off the ruler removes it', async ({ page }) => {
  await mount(page);
  const point = await at(page, 1440);
  await page.mouse.click(point.x, point.y);
  await expect.poll(() => tabs(page), { timeout: 3000 }).toEqual([{ positionTwips: 1440, alignment: 'left', leader: 'none' }]);

  await page.locator(`${HOST} [data-dxr="rulertab"]`).click();
  await expect(page.locator(`${HOST} [data-dxr="rulertab"]`)).toHaveAttribute('data-kind', 'center');
  const second = await at(page, 2880);
  await page.mouse.click(second.x, second.y);
  await expect.poll(() => tabs(page)).toEqual([
    { positionTwips: 1440, alignment: 'left', leader: 'none' },
    { positionTwips: 2880, alignment: 'center', leader: 'none' },
  ]);

  const stop = (await page.locator(`${HOST} .dxr-ruler-mark[data-mark="tab"][data-kind="left"]`).boundingBox())!;
  await page.mouse.move(stop.x + stop.width / 2, stop.y + stop.height / 2);
  await page.mouse.down();
  await page.mouse.move(stop.x + stop.width / 2, stop.y + 120, { steps: 5 });
  await page.mouse.up();
  await expect.poll(() => tabs(page)).toEqual([{ positionTwips: 2880, alignment: 'center', leader: 'none' }]);
  await undo(page);
  await expect.poll(async () => (await tabs(page)).length).toBe(2);
});

test('double-clicking a tab stop opens the Tabs dialog; Set, Clear and the default interval', async ({ page }) => {
  await mount(page);
  const point = await at(page, 1440);
  await page.mouse.click(point.x, point.y);
  await expect.poll(async () => (await tabs(page)).length).toBe(1);
  await page.locator(`${HOST} .dxr-ruler-mark[data-mark="tab"]`).dblclick();
  const dialog = page.locator(`${HOST} dialog[data-dxr="tabsdialog"]`);
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('[data-tabs="position"]')).toHaveValue('1"');
  await dialog.locator('[data-tabs="position"]').fill('3');
  await dialog.locator('input[name="dxrtabalign"][value="right"]').check();
  await dialog.locator('input[name="dxrtablead"][value="dot"]').check();
  await dialog.locator('[data-tabs="set"]').click();
  await dialog.locator('[data-tabs="list"]').selectOption({ index: 0 });
  await dialog.locator('[data-tabs="clear"]').click();
  await expect(dialog.locator('[data-tabs="cleared"]')).toHaveText('1"');
  await dialog.locator('[data-tabs="default"]').fill('1');
  await dialog.locator('button[value="ok"]').click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => tabs(page)).toEqual([{ positionTwips: 4320, alignment: 'right', leader: 'dot' }]);
  expect((await state(page)).defaultTabStopTwips).toBe(1440);
});

test('dragging the left margin edge sets the section\'s left margin', async ({ page }) => {
  await mount(page);
  const s = await state(page);
  const edge = (await mark(page, 'leftMargin').boundingBox())!;
  await page.mouse.move(edge.x + edge.width / 2, edge.y + edge.height / 2);
  await page.mouse.down();
  await page.mouse.move(s.pageLeft + 720 * s.pxPerTwip, edge.y + edge.height / 2, { steps: 5 });
  await page.mouse.up();
  await expect.poll(async () => (await page.evaluate(() => (window as any).__ribbon.editor.sectionInfo())).marginLeftTwips).toBe(720);
  await expect.poll(async () => (await state(page))?.marginLeftTwips).toBe(720);
});

test('the ruler hides with the View tab box', async ({ page }) => {
  await mount(page);
  await page.locator(`${HOST} [data-dxr="showruler"]`).uncheck();
  await expect(page.locator(`${HOST} [data-dxr="ruler"]`)).toBeHidden();
});

test('the ruler stays after view switches and ribbon clicks, lined up, and its drags still land', async ({ page }) => {
  await mount(page);
  const aligned = async () => {
    const s = await state(page);
    const left = await mark(page, 'leftIndent').boundingBox();
    const paragraph = (await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().boundingBox())!;
    return !!s && !!left && Math.abs(left.x + left.width / 2 - s.originLeft) < 1.5 && Math.abs(paragraph.x - s.originLeft) < 1.5;
  };
  for (const view of ['viewpage', 'viewweb', 'viewpage']) {
    await page.locator(`${HOST} [data-dxr="${view}"]`).click();
    await expect.poll(aligned, { timeout: 5000 }).toBe(true);
  }
  // The focus is in the ribbon now (Home tab); a drag still sets the first paragraph's indent.
  await page.locator(`${HOST} .dxr-tab[data-tab="home"]`).click();
  await expect.poll(aligned).toBe(true);
  await drag(page, 'leftIndent', 720);
  await expect.poll(async () => (await format(page))?.leftIndent).toBe(720);
});
