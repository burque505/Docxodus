import { test, expect, type Page } from '@playwright/test';

/**
 * Find options (Word's More >>): match case, whole words, prefix/suffix, ignore punctuation and
 * white space, the ^ codes (^p finds over paragraph marks), Word wildcards and regular
 * expressions; the find bar's Options ▾ and Special ▾ menus.
 */

const blocks = (page: Page) => page.locator('#fo-host [data-anchor][contenteditable="true"]');
const exec = (page: Page, name: string, args: object = {}) =>
  page.evaluate(([n, a]) => (window as any).__ribbon.exec(n, a), [name, args] as const);
const texts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#fo-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/ /g, ' ').trim()));

async function boot(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

async function mount(page: Page, paragraphs: string[], options: object = {}) {
  await boot(page);
  await page.evaluate((opts) => {
    const host = document.createElement('div');
    host.id = 'fo-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const changes: object[] = [];
    const commands: object[] = [];
    (window as any).__findOptionChanges = changes;
    (window as any).__commands = commands;
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {
      ...opts,
      onFindOptionsChange: (o: object) => changes.push(o),
      onCommand: (_label: string, _ms: number, command?: object) => { if (command) commands.push(command); },
    });
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('find.docx');
    ribbon.loader.done();
    (window as any).__ribbon = ribbon;
  }, options);
  await blocks(page).first().click();
  for (const [i, p] of paragraphs.entries()) {
    await page.keyboard.type(p);
    if (i < paragraphs.length - 1) await page.keyboard.press('Enter');
  }
  await expect(blocks(page)).toHaveCount(paragraphs.length);
}

/** Every hit of `query` in `text` (paragraphs joined by \r), as the compiler finds them. */
function hits(page: Page, query: string, text: string, options: object = {}) {
  return page.evaluate(([q, t, o]) => {
    const compiled = (window as any).DocxodusEditor.compileFind(q, o);
    if (!compiled.ok) return 'ERROR: ' + compiled.error;
    return Array.from((t as string).matchAll(compiled.regex), (m: RegExpMatchArray) => m[0]);
  }, [query, text, options] as const);
}

/** `text` with every hit of `query` replaced as Replace all would. */
function replaced(page: Page, query: string, text: string, replacement: string, options: object = {}) {
  return page.evaluate(([q, t, r, o]) => {
    const { compileFind, expandReplacement } = (window as any).DocxodusEditor;
    const compiled = compileFind(q, o);
    if (!compiled.ok) return 'ERROR: ' + compiled.error;
    return (t as string).replace(compiled.regex, (...args: any[]) => {
      const groups = typeof args[args.length - 1] === 'object' ? args[args.length - 1] : undefined;
      const match = Object.assign(args.slice(0, typeof groups === 'object' ? -3 : -2), { groups }) as unknown as RegExpExecArray;
      return expandReplacement(r, match, o);
    });
  }, [query, text, replacement, options] as const);
}

test.describe('find options: the query compiler', () => {
  test('plain text: case, whole words, prefix and suffix', async ({ page }) => {
    await boot(page);
    expect(await hits(page, 'cat', 'Cat cat concat cats')).toEqual(['Cat', 'cat', 'cat', 'cat']);
    expect(await hits(page, 'cat', 'Cat cat concat cats', { matchCase: true })).toEqual(['cat', 'cat', 'cat']);
    expect(await hits(page, 'cat', 'Cat cat concat cats', { wholeWord: true })).toEqual(['Cat', 'cat']);
    expect(await hits(page, 'cat', 'cat concat cats', { matchPrefix: true })).toEqual(['cat', 'cat']);
    expect((await hits(page, 'cat', 'cat concat cats', { matchSuffix: true })).length).toBe(2);
    // Regex characters in a plain query are only text.
    expect(await hits(page, 'a.b (c)', 'a.b (c) axb (c)')).toEqual(['a.b (c)']);
    expect(await hits(page, 'café', 'Café CAFÉ', { wholeWord: true })).toEqual(['Café', 'CAFÉ']);
  });

  test('ignore punctuation and white space', async ({ page }) => {
    await boot(page);
    expect(await hits(page, 'email', 'email e-mail e.mail', { ignorePunctuation: true })).toEqual(['email', 'e-mail', 'e.mail']);
    expect(await hits(page, 'e-mail', 'email', { ignorePunctuation: true })).toEqual(['email']);
    expect(await hits(page, 'ice cream', 'icecream ice  cream ice\tcream', { ignoreSpace: true })).toEqual(['icecream', 'ice  cream', 'ice\tcream']);
    // A paragraph mark is never skipped.
    expect(await hits(page, 'ice cream', 'ice\rcream', { ignoreSpace: true })).toEqual([]);
  });

  test('Word ^ codes', async ({ page }) => {
    await boot(page);
    expect(await hits(page, 'end^pnext', 'the end\rnext one')).toEqual(['end\rnext']);
    expect(await hits(page, '^t', 'a\tb')).toEqual(['\t']);
    expect(await hits(page, '^#^#', 'a12b3')).toEqual(['12']);
    expect(await hits(page, 'b^$', 'ba b1 bc')).toEqual(['ba', 'bc']);
    expect(await hits(page, 'a^?c', 'abc a\rc')).toEqual(['abc']);
    expect(await hits(page, 'a^wb', 'a   b ab')).toEqual(['a   b']);
    expect(await hits(page, '^^', 'x^y')).toEqual(['^']);
    expect(await hits(page, '^l', 'a\vb\fc')).toEqual(['\v']);
    expect(await hits(page, '^m', 'a\vb\fc')).toEqual(['\f']);
    expect(await hits(page, '^0151', 'a—b')).toEqual(['—']);
    expect(await hits(page, '^8212', 'a—b')).toEqual(['—']);
    expect(await hits(page, '^u8212', 'a—b')).toEqual(['—']);
    expect(await hits(page, '^z', 'z')).toMatch(/^ERROR/);
    expect(await replaced(page, '^p^p', 'a\r\rb\r\r\rc', '^p')).toBe('a\rb\r\rc');
    expect(await replaced(page, 'x', 'axb', '[^&]^t')).toBe('a[x]\tb');
  });

  test('Word wildcards', async ({ page }) => {
    await boot(page);
    const w = { wildcards: true };
    expect(await hits(page, 'lo@t', 'lot loot lt', w)).toEqual(['lot', 'loot']);
    expect(await hits(page, 's?t', 'sat set st', w)).toEqual(['sat', 'set']);
    expect(await hits(page, 's*d', 'sad said sound', w)).toEqual(['sad', 'said', 'sound']);
    expect(await hits(page, '<(inter)*>', 'interesting winter intern', w)).toEqual(['interesting', 'intern']);
    expect(await hits(page, '[!c]at', 'cat bat hat', w)).toEqual(['bat', 'hat']);
    expect(await hits(page, '[0-9]{2,3}', '1 12 1234', w)).toEqual(['12', '123']);
    expect(await hits(page, '[0-9]{2;}', '1 12 1234', w)).toEqual(['12', '1234']);
    expect(await hits(page, '\\?', 'what?', w)).toEqual(['?']);
    // Always case-sensitive, as in Word.
    expect(await hits(page, 'Cat', 'cat Cat', w)).toEqual(['Cat']);
    expect(await hits(page, 'end^13', 'end\rx', w)).toEqual(['end\r']);
    expect(await replaced(page, '(<[A-Z][a-z]@>), (<[A-Z][a-z]@>)', 'Smith, John', '\\2 \\1', w)).toBe('John Smith');
    expect(await hits(page, '[abc', 'a', w)).toMatch(/^ERROR/);
  });

  test('regular expressions', async ({ page }) => {
    await boot(page);
    const r = { regex: true };
    expect(await hits(page, '\\d+', 'a 12 b 345', r)).toEqual(['12', '345']);
    expect(await hits(page, 'abc', 'ABC abc', r)).toEqual(['ABC', 'abc']);
    expect(await hits(page, 'abc', 'ABC abc', { ...r, matchCase: true })).toEqual(['abc']);
    expect(await hits(page, 'i)abc', 'ABC abc', { ...r, matchCase: true })).toEqual(['ABC', 'abc']);
    // ^ and $ are paragraph starts and ends; \n is a paragraph mark.
    expect(await hits(page, '^Chapter \\d', 'Chapter 1\rsee Chapter 2\rChapter 3', r)).toEqual(['Chapter 1', 'Chapter 3']);
    expect(await hits(page, 'end\\nnext', 'end\rnext', r)).toEqual(['end\rnext']);
    expect(await hits(page, '\\p{Lu}\\w+', 'hello World', { ...r, matchCase: true })).toEqual(['World']);
    expect(await replaced(page, '(\\w+) (\\w+)', 'hello world', '$2 $1', r)).toBe('world hello');
    expect(await replaced(page, '(?<a>x)', 'axb', '[$<a>$&$$]', r)).toBe('a[xx$]b');
    expect(await replaced(page, ',', 'a,b', '\\n', r)).toBe('a\rb');
    expect(await hits(page, '(', 'a', r)).toMatch(/^ERROR/);
  });
});

test.describe('find options in the editor', () => {
  test('^p^p -> ^p removes empty paragraphs, in one undo step', async ({ page }) => {
    await mount(page, ['one', '', 'two', '', '', 'three']);
    const count = await exec(page, 'replaceAll', { text: '^p^p', with: '^p' });
    expect(count).toBe(2);
    // the 3 empty paragraphs: ^p^p twice leaves one of the two after "two"
    expect(await texts(page)).toEqual(['one', 'two', '', 'three']);
    await exec(page, 'undo');
    await expect.poll(() => texts(page)).toEqual(['one', '', 'two', '', '', 'three']);
  });

  test('a paragraph mark replaced with a space joins lines; a comma with ^p splits them', async ({ page }) => {
    await mount(page, ['first line', 'second line', 'third']);
    expect(await exec(page, 'replaceAll', { text: '^p', with: ' ' })).toBe(2);
    expect(await texts(page)).toEqual(['first line second line third']);
    expect(await exec(page, 'replaceAll', { text: ' line', with: ',^p' })).toBe(2);
    expect(await texts(page)).toEqual(['first,', 'second,', 'third']);
  });

  test('find counts and selects a hit over a paragraph mark', async ({ page }) => {
    await mount(page, ['the end', 'start here']);
    expect(await exec(page, 'find', { text: 'end^pstart' })).toBe(1);
    // (Selection.toString() stops at the focused block; the range is what spans both.)
    const selected = await page.evaluate(() => {
      const all = Array.from(document.querySelectorAll('#fo-host [data-anchor][contenteditable="true"]'));
      const r = window.getSelection()!.getRangeAt(0);
      const index = (n: Node) => all.findIndex((b) => b.contains(n));
      return { from: index(r.startContainer), to: index(r.endContainer), text: r.toString().replace(/\s+/g, '') };
    });
    expect(selected).toEqual({ from: 0, to: 1, text: 'endstart' });
  });

  test('regex and wildcard replacements use the found groups', async ({ page }) => {
    await mount(page, ['Smith, John', 'call 555 1234']);
    expect(await exec(page, 'replaceAll', { text: '(<[A-Z][a-z]@>), (<[A-Z][a-z]@>)', with: '\\2 \\1', wildcards: true })).toBe(1);
    expect(await exec(page, 'replaceAll', { text: '(\\d{3}) (\\d{4})', with: '$1-$2', regex: true })).toBe(1);
    expect(await texts(page)).toEqual(['John Smith', 'call 555-1234']);
  });

  test('a tab in the replacement is a real tab', async ({ page }) => {
    await mount(page, ['a;b']);
    expect(await exec(page, 'replaceAll', { text: ';', with: '^t' })).toBe(1);
    await expect(page.locator('#fo-host [data-anchor][contenteditable="true"] [data-docx-tab]').first()).toBeAttached();
  });

  test('an invalid pattern is a command error', async ({ page }) => {
    await mount(page, ['x']);
    const error = await page.evaluate(() => { try { (window as any).__ribbon.exec('find', { text: '(', regex: true }); return ''; } catch (e) { return String(e); } });
    expect(error).toMatch(/find:/);
  });
});

test.describe('find bar menus', () => {
  test('Options: checking whole words narrows the hits, reports the options, and is recorded', async ({ page }) => {
    await mount(page, ['cat concat cats cat']);
    await page.keyboard.press('Control+h');
    await page.locator('#fo-host [data-dxr="findtext"]').fill('cat');
    await expect(page.locator('#fo-host [data-dxr="findcount"]')).toHaveText('1 of 4');
    await page.locator('#fo-host [data-dxr="findoptions"]').click();
    await page.locator('[data-findopt="wholeWord"]').check();
    await expect(page.locator('#fo-host [data-dxr="findcount"]')).toHaveText('1 of 2');
    expect(await page.evaluate(() => (window as any).__findOptionChanges)).toEqual([{ wholeWord: true }]);
    await expect(page.locator('#fo-host [data-dxr="findoptions"]')).toHaveAttribute('data-active', '');
    // Regular expressions grey out whole words; wildcards then take their place.
    await page.locator('[data-findopt="regex"]').check();
    await expect(page.locator('[data-findopt="wholeWord"]')).toBeDisabled();
    await page.locator('[data-findopt="wildcards"]').check();
    await expect(page.locator('[data-findopt="regex"]')).not.toBeChecked();
    await expect(page.locator('[data-findopt="matchCase"]')).toBeChecked();
    await page.locator('[data-findopt="wildcards"]').uncheck();
    // A click outside closes the menu (Escape in the find field would close the bar).
    await page.locator('#fo-host [data-dxr="replacetext"]').fill('dog');
    await page.locator('#fo-host [data-dxr="replaceall"]').click();
    expect(await texts(page)).toEqual(['dog concat cats dog']);
    const commands = await page.evaluate(() => (window as any).__commands);
    expect(commands).toContainEqual({ name: 'replaceAll', args: { text: 'cat', with: 'dog', wholeWord: true } });
  });

  test('an invalid pattern says so in the bar', async ({ page }) => {
    await mount(page, ['x'], { findOptions: { regex: true } });
    await page.keyboard.press('Control+f');
    await page.locator('#fo-host [data-dxr="findtext"]').fill('(x');
    await expect(page.locator('#fo-host [data-dxr="findcount"]')).toHaveText('Invalid');
    await expect(page.locator('#fo-host [data-dxr="findtext"]')).toHaveAttribute('aria-invalid', 'true');
  });

  test('Special inserts the code for the mode into the field last typed in', async ({ page }) => {
    await mount(page, ['a', 'b']);
    await page.keyboard.press('Control+h');
    const find = page.locator('#fo-host [data-dxr="findtext"]');
    await find.fill('a');
    await page.locator('#fo-host [data-dxr="findspecial"]').click();
    await page.locator('[data-dxr="findspecialmenu"] button[data-code="^p"]').click();
    await expect(find).toHaveValue('a^p');
    await expect(page.locator('#fo-host [data-dxr="findcount"]')).toHaveText('1 of 1');
    const replace = page.locator('#fo-host [data-dxr="replacetext"]');
    await replace.click();
    await page.locator('#fo-host [data-dxr="findspecial"]').click();
    await page.locator('[data-dxr="findspecialmenu"] button[data-code="^&"]').click();
    await expect(replace).toHaveValue('^&');
    await page.evaluate(() => (window as any).__ribbon.setFindOptions({ regex: true }));
    await find.click();
    await page.locator('#fo-host [data-dxr="findspecial"]').click();
    await expect(page.locator('[data-dxr="findspecialmenu"] button[data-code="\\\\n"]')).toBeVisible();
  });

  test('findOptions at mount checks the boxes', async ({ page }) => {
    await mount(page, ['x'], { findOptions: { matchCase: true, ignoreSpace: true } });
    expect(await page.evaluate(() => (window as any).__ribbon.getFindOptions())).toEqual({ matchCase: true, ignoreSpace: true });
    await expect(page.locator('[data-dxr="findcase"]')).toBeChecked();
    await expect(page.locator('[data-findopt="ignoreSpace"]')).toBeChecked();
  });
});
