import { test, expect, type Page } from '@playwright/test';

/**
 * Page view lays the pages out again when an edit overfills one: Enter or typing at a page's foot
 * no longer leaves text cut off past the page's text area. The caret stays where it was.
 */

const HOST = '#pf-host';
const blocks = `${HOST} [data-anchor][contenteditable="true"]`;

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'pf-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('pf.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
  await page.locator(blocks).first().click();
  const lines = Array.from({ length: 60 }, (_, i) => `Line ${i + 1} of the document`);
  await page.evaluate((t) => (window as any).__ribbon.exec('pasteText', { text: t }), lines.join('\n'));
  await page.evaluate(() => (window as any).__ribbon.editor.setPaginated(true));
  await expect(page.locator(`${HOST} .page-box [data-anchor][contenteditable="true"]`).first()).toBeVisible();
}

/** The blocks reaching past their page's text area. */
const overflowing = (page: Page) => page.evaluate(() => {
  const out: string[] = [];
  for (const content of Array.from(document.querySelectorAll<HTMLElement>('#pf-host .page-box .page-content'))) {
    const area = content.getBoundingClientRect();
    for (const b of Array.from(content.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))) {
      const r = b.getBoundingClientRect();
      if (r.height > 0 && r.bottom > area.bottom + 2) out.push(b.textContent ?? '');
    }
  }
  return out;
});

/** The text of the last block on the first page. */
const lastOnFirstPage = (page: Page) => page.evaluate(() => {
  const box = document.querySelector<HTMLElement>('#pf-host .page-box .page-content')!;
  const bl = box.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]');
  return (bl[bl.length - 1].textContent ?? '').trim();
});

test('Enter at the foot of a page moves what no longer fits onto the next page; the caret stays', async ({ page }) => {
  await mount(page);
  const last = await lastOnFirstPage(page);
  expect(last).toMatch(/^Line \d+ of the document$/);
  const n = Number(/\d+/.exec(last)![0]);
  // the caret after "Line N" on the first page's last line, then Enter three times
  const p = await page.evaluate(() => {
    const box = document.querySelector<HTMLElement>('#pf-host .page-box .page-content')!;
    const bl = box.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]');
    const b = bl[bl.length - 2];
    b.scrollIntoView({ block: 'center' });
    const r = b.getBoundingClientRect();
    return { x: r.left + 4, y: r.top + r.height / 2 };
  });
  await page.mouse.click(p.x, p.y);
  await page.keyboard.press('End');
  for (let i = 0; i < 3; i++) await page.keyboard.press('Enter');
  await page.keyboard.type('Typed');
  // laid out again once the edits pause: nothing cut off, the last lines moved on
  await expect.poll(() => overflowing(page), { timeout: 5000 }).toEqual([]);
  expect(await lastOnFirstPage(page)).not.toBe(last);
  // the caret is still in the new paragraph: more typing goes there
  await page.keyboard.type(' more');
  await expect.poll(() => page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLElement>('#pf-host [data-anchor][contenteditable="true"]'))
      .map((b) => (b.textContent ?? '').trim()).filter((t) => t.startsWith('Typed')))).toEqual(['Typed more']);
  // and the line that was the first page's last is now on the second page, in order
  const order = await page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#pf-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').trim()).filter((t) => t));
  expect(order.indexOf('Typed more')).toBe(order.indexOf(`Line ${n - 1} of the document`) + 1);
});

test('no lay-out while nothing overflows: typing mid-page leaves the pages as they are', async ({ page }) => {
  await mount(page);
  await page.evaluate(() => { (window as any).__pageBox = document.querySelector('#pf-host .page-box'); });
  await page.locator(blocks).nth(3).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' x');
  await page.waitForTimeout(900);
  // (the same page box element: no remount)
  expect(await page.evaluate(() => (window as any).__pageBox === document.querySelector('#pf-host .page-box'))).toBe(true);
});
