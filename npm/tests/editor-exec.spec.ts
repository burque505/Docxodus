import { test, expect, Page } from '@playwright/test';

/**
 * DocxEditor.exec / DocxEditor.commands() (editor-commands.ts) and RibbonEditor.exec: editing
 * commands by name with plain-object arguments, for hosts that drive the editor from outside
 * (a scripting bridge, a command palette, a macro recorder).
 */

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/** A blank document in a bare DocxEditor, the caret in its paragraph. */
async function openBlank(page: Page) {
  await page.goto('/test-harness.html');
  await waitForDocxodus(page);
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'ex-editor';
    document.body.appendChild(container);
    const ed = D.DocxEditor.openBlank(container, D, {});
    (window as any).__ed = ed;
    container.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.focus();
  });
}

/** The saved document's text, read back through a fresh editor: what the session holds. */
const saved = (page: Page) => page.evaluate(() => {
  const D = (window as any).Docxodus;
  const probe = document.createElement('div');
  const reader = D.DocxEditor.open(probe, (window as any).__ed.save(), D, {});
  const text = reader.exec('text');
  reader.close();
  return text as string;
});

const exec = (page: Page, name: string, args?: object) =>
  page.evaluate(({ name, args }) => (window as any).__ed.exec(name, args), { name, args });

/** exec that is expected to throw: the error's name and code. */
const execError = (page: Page, name: string, args?: unknown) =>
  page.evaluate(({ name, args }) => {
    try {
      (window as any).__ed.exec(name, args);
      return null;
    } catch (e: any) {
      return { name: e.name, code: e.code, command: e.command, message: e.message };
    }
  }, { name, args });

test.describe('DocxEditor.exec', () => {
  test('commands() lists every command with its group, description and arguments', async ({ page }) => {
    await openBlank(page);
    const commands = await page.evaluate(() => (window as any).Docxodus.DocxEditor.commands());
    const names = commands.map((c: any) => c.name);
    for (const name of ['bold', 'fontSize', 'style', 'insertText', 'find', 'replaceAll', 'undo', 'text', 'selectedText'])
      expect(names).toContain(name);
    expect(new Set(names).size).toBe(names.length);
    for (const c of commands) {
      expect(c.group).toMatch(/^(font|paragraph|insert|layout|table|review|edit|view|query)$/);
      expect(c.description.length).toBeGreaterThan(5);
      expect(typeof c.args).toBe('object');
    }
    expect(commands.find((c: any) => c.name === 'fontSize').args).toEqual({ pts: 'number' });
  });

  test('a wrong name, a wrong argument and a closed editor fail with typed errors', async ({ page }) => {
    await openBlank(page);
    expect(await execError(page, 'noSuchCommand')).toMatchObject({ name: 'EditorCommandError', code: 'unknown_command' });
    expect(await execError(page, 'toString')).toMatchObject({ code: 'unknown_command' });
    expect(await execError(page, 'fontSize', { pts: 'large' })).toMatchObject({ code: 'bad_arguments', command: 'fontSize' });
    expect(await execError(page, 'fontSize', {})).toMatchObject({ code: 'bad_arguments' });
    expect(await execError(page, 'align', { alignment: 'middle' })).toMatchObject({ code: 'bad_arguments' });
    expect(await execError(page, 'bold', [1])).toMatchObject({ code: 'bad_arguments' });
    await page.evaluate(() => (window as any).__ed.close());
    expect(await execError(page, 'bold')).toMatchObject({ code: 'closed' });
  });

  test('insertText types at the caret through the engine, one undo step', async ({ page }) => {
    await openBlank(page);
    expect(await exec(page, 'insertText', { text: 'Hello world' })).toBeNull();
    expect(await exec(page, 'text')).toBe('Hello world');
    expect(await saved(page)).toBe('Hello world');
    // the caret is after the text: the next insert continues it
    await exec(page, 'insertText', { text: '!' });
    expect(await saved(page)).toBe('Hello world!');
    await exec(page, 'undo');
    expect(await saved(page)).toBe('Hello world');
    await exec(page, 'undo');
    expect(await saved(page)).toBe('');
  });

  test('insertText starts a new paragraph at each line break', async ({ page }) => {
    await openBlank(page);
    await exec(page, 'insertText', { text: 'Line one\nLine two\r\nLine three' });
    expect(await saved(page)).toBe('Line one\nLine two\nLine three');
    expect(await exec(page, 'paragraphText')).toBe('Line three');
    await exec(page, 'newParagraph');
    await exec(page, 'insertText', { text: 'Line four' });
    expect(await saved(page)).toBe('Line one\nLine two\nLine three\nLine four');
  });

  test('insertText replaces the selection, and the selection can be read', async ({ page }) => {
    await openBlank(page);
    await exec(page, 'insertText', { text: 'Hello world, hello again' });
    expect(await exec(page, 'find', { text: 'hello', matchCase: true })).toBe(1);
    expect(await exec(page, 'selectedText')).toBe('hello');
    await exec(page, 'insertText', { text: 'goodbye' });
    expect(await saved(page)).toBe('Hello world, goodbye again');
    expect(await exec(page, 'find', { text: 'HELLO' })).toBe(1);
    expect(await exec(page, 'replaceAll', { text: 'again', with: 'for now' })).toBe(1);
    expect(await saved(page)).toBe('Hello world, goodbye for now');
  });

  test('moveCaret, formatting and queries act where the user would', async ({ page }) => {
    await openBlank(page);
    await exec(page, 'insertText', { text: 'First\nSecond' });
    await exec(page, 'moveCaret', { to: 'documentStart' });
    await exec(page, 'insertText', { text: 'The ' });
    await exec(page, 'moveCaret', { to: 'documentEnd' });
    await exec(page, 'insertText', { text: ' end' });
    expect(await saved(page)).toBe('The First\nSecond end');

    expect(await exec(page, 'find', { text: 'Second' })).toBe(1);
    await exec(page, 'bold', { on: true });
    expect((await exec(page, 'formatState') as any).bold).toBe(true);
    await exec(page, 'style', { styleId: 'Heading1' });
    expect(await exec(page, 'styleAtCaret')).toBe('Heading1');
    expect(await exec(page, 'wordCount')).toBe(4);
    expect(typeof await exec(page, 'version')).toBe('number');
  });
});

test.describe('RibbonEditor.exec', () => {
  test('runs commands as the ribbon does, reports them to onCommand, and keeps its controls in step', async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
    const result = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const host = document.createElement('div');
      host.style.height = '700px';
      document.body.appendChild(host);
      const reported: Array<{ label: string; command: unknown }> = [];
      const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {
        onCommand: (label: string, _ms: number, command?: unknown) => reported.push({ label, command }),
      });
      ribbon.setExports(D);
      ribbon.openBlank('exec.docx');
      host.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.focus();
      ribbon.exec('insertText', { text: 'Ribbon text' });
      ribbon.exec('zoom', { scale: 1.5 });
      ribbon.exec('trackChanges', { on: true });
      ribbon.exec('pageWidth');
      const fits = ribbon.editor.fitsPageWidth;
      ribbon.exec('onePage');
      const whole = ribbon.editor.fitsOnePage;
      ribbon.exec('zoom', { scale: 1.5 });
      let error: unknown = null;
      try { ribbon.exec('fontSize', { pts: -1 }); } catch (e: any) { error = e.code; }
      return {
        text: ribbon.exec('text'),
        zoomShown: ribbon.control('zoom').value,
        tracking: ribbon.control('trackchanges').checked,
        rail: ribbon.control('railOp').textContent,
        reported,
        error,
        listed: ribbon.commands().length === D.DocxEditor.commands().length,
        fits, whole, pinned: !ribbon.editor.fitsPageWidth && !ribbon.editor.fitsOnePage,
      };
    });
    expect(result.text).toBe('Ribbon text');
    expect(result.zoomShown).toBe('1.5');
    expect(result.tracking).toBe(true);
    expect(result.rail).toMatch(/^fontSize \d/); // (the "text" query after it is no command)
    expect(result.error).toBe('bad_arguments');
    expect(result.listed).toBe(true);
    expect([result.fits, result.whole, result.pinned]).toEqual([true, true, true]);
    expect(result.reported.slice(0, 3)).toEqual([
      { label: 'insertText', command: { name: 'insertText', args: { text: 'Ribbon text' } } },
      { label: 'zoom', command: { name: 'zoom', args: { scale: 1.5 } } },
      { label: 'trackChanges', command: { name: 'trackChanges', args: { on: true } } },
    ]);
    // a command that failed is reported without its name and arguments (nothing to record)
    expect(result.reported.slice(3, 6).map((r: any) => r.command?.name)).toEqual(['pageWidth', 'onePage', 'zoom']);
    expect(result.reported[6]).toEqual({ label: 'fontSize', command: undefined });
    expect(result.reported).toHaveLength(7);
  });
});
