import { test, expect } from '@playwright/test';

/**
 * RibbonOptions.onCommand reports every command as `exec` takes it, whatever ran it: a button, a
 * list, the table grid, a shortcut (the editor's own Ctrl+I and Ctrl+Z among them). Running the
 * reported commands again through `exec` repeats what the user did: what a macro recorder keeps.
 */
test('the ribbon reports its commands as exec takes them, and replaying them repeats the edit', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const host = document.createElement('div');
    host.id = 'rec-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const reported: Array<{ label: string; command: unknown }> = [];
    (window as any).__reported = reported;
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {
      onCommand: (label: string, _ms: number, command?: unknown) => reported.push({ label, command }),
    });
    ribbon.setExports(D);
    ribbon.openBlank('record.docx');
    ribbon.loader.done();    // (the loading overlay, which a host takes down)
    (window as any).__ribbon = ribbon;
  });
  const block = page.locator('#rec-host [data-anchor][contenteditable="true"]').first();
  await block.click();
  await page.keyboard.type('Recorded heading');

  // a button, a list, the editor's own shortcuts, a shortcut the ribbon takes, the table grid
  await page.evaluate(() => {
    const host = document.querySelector('#rec-host')!;
    const control = (name: string) => host.querySelector<HTMLElement>(`[data-dxr="${name}"]`)!;
    host.querySelector<HTMLElement>('button[data-cmd="bold"]')!.click();
    const style = control('style') as HTMLSelectElement;
    style.value = 'Heading1';
    style.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await block.press('Control+i');
  await block.press('Control+z');
  await block.press('Control+e');
  await page.evaluate(() => {
    const host = document.querySelector('#rec-host')!;
    const control = (name: string) => host.querySelector<HTMLElement>(`[data-dxr="${name}"]`)!;
    control('zoomin').click();
    control('table').click();
    const cell = host.querySelector<HTMLElement>('[data-dxr="gridcells"] [data-r="1"][data-c="2"]')!;
    cell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
  });

  const result = await page.evaluate(() => {
    const reported = (window as any).__reported as Array<{ label: string; command?: { name: string; args: any } }>;
    const ribbon = (window as any).__ribbon;
    const commands = reported.map((r) => r.command).filter(Boolean) as Array<{ name: string; args: any }>;
    const state = (r: any, h: Element) => {
      r.exec('moveCaret', { to: 'documentStart' });
      return { text: r.exec('text'), style: r.exec('styleAtCaret'), tables: h.querySelectorAll('table').length };
    };
    const first = state(ribbon, document.querySelector('#rec-host')!);

    // Replay into a fresh document: the typing (what a recorder keeps from the keys) and the commands.
    const D = (window as any).Docxodus;
    const host2 = document.createElement('div');
    host2.style.height = '700px';
    document.body.appendChild(host2);
    const replay = (window as any).DocxodusEditor.mountRibbon(host2, {});
    replay.setExports(D);
    replay.openBlank('replay.docx');
    host2.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.focus();
    replay.exec('insertText', { text: 'Recorded heading' });
    replay.exec('moveCaret', { to: 'paragraphStart' });
    const errors: string[] = [];
    for (const c of commands) {
      // (the typed text was selected by nothing: bold and italic act on the whole paragraph in both)
      try { replay.exec(c.name, c.args); } catch (e: any) { errors.push(`${c.name}: ${e.message}`); }
    }
    const second = state(replay, host2);
    return { commands, labels: reported.map((r) => r.label), first, second, errors };
  });

  const names = result.commands.map((c) => c.name);
  expect(names).toEqual(['bold', 'style', 'italic', 'undo', 'align', 'zoom', 'table']);
  expect(result.commands[1]).toEqual({ name: 'style', args: { styleId: 'Heading1' } });
  expect(result.commands[4]).toEqual({ name: 'align', args: { alignment: 'center' } });
  expect(result.commands[5].args.scale).toBeGreaterThan(1);
  expect(result.commands[6]).toEqual({ name: 'table', args: { rows: 2, columns: 3, align: expect.any(String), borderless: false } });
  expect(result.errors).toEqual([]);
  expect(result.second.text).toBe(result.first.text);
  expect(result.second.tables).toBe(result.first.tables);
  expect(result.first.tables).toBe(1);
  expect(result.first.style).toBe('Heading1');
  expect(result.second.style).toBe(result.first.style);
});

test('every command a control reports is one exec knows, with arguments it takes', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  const result = await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const host = document.createElement('div');
    host.style.height = '700px';
    document.body.appendChild(host);
    const reported: Array<{ label: string; command?: { name: string; args: any } }> = [];
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {
      onCommand: (label: string, _ms: number, command?: any) => reported.push({ label, command }),
    });
    ribbon.setExports(D);
    ribbon.openBlank('all.docx');
    host.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!.focus();
    ribbon.exec('insertText', { text: 'Some text' });
    reported.length = 0;
    // Every plain button outside popovers, the find bar and the file menu, and one choice of every list.
    const skip = (el: Element) => !!el.closest('.dxr-pop, [data-dxr="findbar"], [data-dxr="loader"], .dxr-quick, .dxr-tab')
      || ['filebutton', 'save', 'new', 'table', 'link', 'comment', 'insertcomment', 'findtoggle', 'replacetoggle',
        'commentprev', 'commentnext', 'commentdelete', 'commentresolve', 'fontcolorbtn', 'shadebtn', 'gotoheader', 'gotofooter',
        'delblock'].includes((el as HTMLElement).dataset.dxr ?? '');
    for (const b of Array.from(host.querySelectorAll<HTMLButtonElement>('.dxr-panel button, .dxr-rail button'))) {
      if (skip(b) || b.disabled || b.closest('[data-panel="table"], [data-panel="headerfooter"], [data-panel="picture"], [data-panel="outlining"]')) continue;
      host.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')?.focus();
      try { b.click(); } catch { /* (a command that throws reports nothing) */ }
    }
    for (const s of Array.from(host.querySelectorAll<HTMLSelectElement>('.dxr-panel select'))) {
      if (skip(s) || s.closest('[data-panel="table"], [data-panel="headerfooter"], [data-panel="picture"], [data-panel="outlining"]')) continue;
      const option = Array.from(s.options).find((o) => o.value !== '' && o.value !== s.value);
      if (!option) continue;
      s.value = option.value;
      s.dispatchEvent(new Event('change', { bubbles: true }));
    }
    const known = new Map(ribbon.commands().map((c: any) => [c.name, c]));
    const problems: string[] = [];
    for (const r of reported) {
      if (!r.command) continue;
      const info: any = known.get(r.command.name);
      if (!info) { problems.push(`${r.label}: unknown command ${r.command.name}`); continue; }
      for (const key of Object.keys(r.command.args))
        if (!(key in info.args) && !(`${key}?` in info.args)) problems.push(`${r.label}: ${r.command.name} has no argument ${key}`);
    }
    return { problems, withCommand: reported.filter((r) => r.command).length, without: reported.filter((r) => !r.command).map((r) => r.label) };
  });
  expect(result.problems).toEqual([]);
  expect(result.withCommand).toBeGreaterThan(20);
  // (what has no exec equivalent: accepting or rejecting the change at the caret, stepping through changes)
  for (const label of result.without) expect(label).toMatch(/^revision (accept|reject|next|prev)$/);
});
