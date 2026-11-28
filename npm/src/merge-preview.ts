/**
 * Word's Preview Results: each merge field in the editor shows a recipient's value in place of
 * «Name», as the merge will fill it, without a change to the document (the value is drawn by CSS
 * over the field's own text, which stays as it is). Next Record and Next Record If move to the
 * next row, as in a sheet of labels; Merge Record # shows the row's number. Fields re-rendered by
 * an edit are previewed again.
 */
import { mergeFieldSample, parseMergeFieldCode } from "./ribbon-mergefields-pane.js";

const ATTR = "data-merge-preview";
/** On a label (table cell) whose merge fields all read a row past the last: shown blank, as the merge leaves it. */
const BLANK = "data-merge-preview-blank";

export const MERGE_PREVIEW_CSS = `
[${ATTR}], [${ATTR}] * { font-size: 0 !important; }
[${ATTR}]::after { content: attr(${ATTR}); font-size: var(--dxr-merge-preview-size, 11pt) !important; }
[${BLANK}] > * { visibility: hidden !important; }
`;

export class MergePreview {
  private observer: MutationObserver | null = null;
  private frame = 0;
  private columns = new Map<string, number>();
  private rows: string[][] = [];
  private index = 0;

  constructor(private readonly root: HTMLElement) {}

  get active(): boolean {
    return this.observer !== null;
  }

  /** Show row `index` (0-based) of `rows`, whose columns are `columns`. */
  show(columns: string[], rows: string[][], index: number): void {
    this.columns = new Map(columns.map((c, i) => [c.toLowerCase(), i]));
    this.rows = rows;
    this.index = index;
    if (!this.observer) {
      this.observer = new MutationObserver(() => this.schedule());
      this.observer.observe(this.root, { childList: true, subtree: true });
    }
    this.apply();
  }

  /** Back to the fields' own text. */
  hide(): void {
    this.observer?.disconnect();
    this.observer = null;
    cancelAnimationFrame(this.frame);
    for (const el of Array.from(this.root.querySelectorAll<HTMLElement>(`[${ATTR}]`))) {
      el.removeAttribute(ATTR);
      el.style.removeProperty("--dxr-merge-preview-size");
    }
    for (const el of Array.from(this.root.querySelectorAll<HTMLElement>(`[${BLANK}]`))) el.removeAttribute(BLANK);
  }

  private schedule(): void {
    cancelAnimationFrame(this.frame);
    this.frame = requestAnimationFrame(() => this.apply());
  }

  private value(name: string, offset: number): string {
    const row = this.rows[this.index + offset];
    const col = this.columns.get(name.trim().toLowerCase());
    return row && col !== undefined ? row[col] ?? "" : "";
  }

  /** Whether the record shown is one a Skip Record If leaves out of the merge. */
  skipped = false;

  private apply(): void {
    let offset = 0;
    // (the fields in document order; a field inside another one is its outer field's business)
    const fields = Array.from(this.root.querySelectorAll<HTMLElement>("[data-field-code]"))
      .filter((f) => !f.parentElement?.closest("[data-field-code]") && !f.closest("[data-hf-band]"));
    const codes = fields.map((f) => (f.dataset.fieldCode ?? "").trim());
    const ctx = (at: number): PreviewContext => ({ value: (name) => this.value(name, at), record: this.index + at + 1, asked });
    // ASK fields first, anywhere: their answers (the row's value of that name, else the default) for the REF fields
    const asked = new Map<string, string>();
    for (const code of codes) {
      if (fieldWord(code) !== "ASK") continue;
      const a = fieldArguments(code);
      if (a.length > 1) asked.set(a[1].toLowerCase(), this.value(a[1], 0) || (fieldSwitch(code, "d") ?? ""));
    }
    let skipped = false;
    const cells = new Map<HTMLElement, boolean>();
    fields.forEach((f, i) => {
      const code = codes[i];
      const word = fieldWord(code);
      // (a condition reads the row it is in; Next Record If then moves on when it holds)
      const here = ctx(offset);
      if (word === "NEXT") offset++;
      else if (word === "NEXTIF" && fieldCondition(code, here)) offset++;
      else if (word === "SKIPIF" && fieldCondition(code, here)) skipped = true;
      const cell = word === "MERGEFIELD" ? f.closest<HTMLElement>("td, th") : null;
      if (cell) cells.set(cell, (cells.get(cell) ?? true) && this.index + offset >= this.rows.length);
      const shown = fieldResult(code, here);
      if (shown === null) return;
      if (!f.hasAttribute(ATTR)) {
        const size = getComputedStyle(f).fontSize;
        if (size && size !== "0px") f.style.setProperty("--dxr-merge-preview-size", size);
      }
      if (f.getAttribute(ATTR) !== shown) f.setAttribute(ATTR, shown);
    });
    for (const el of Array.from(this.root.querySelectorAll<HTMLElement>(`[${BLANK}]`))) if (!cells.get(el)) el.removeAttribute(BLANK);
    for (const [cell, past] of cells) if (past && !cell.hasAttribute(BLANK)) cell.setAttribute(BLANK, "");
    if (skipped !== this.skipped) { this.skipped = skipped; this.onChange?.(); }
  }

  /** Called when what the preview says about the record changes (Skip Record If). */
  onChange: (() => void) | null = null;
}

// ─── Field codes evaluated as the host's merge does (host/Merge.cs) ──────────────────────────

/** What a field reads while it is evaluated: a column's value in its row, the row's number, ASK's answers. */
export interface PreviewContext {
  value(name: string): string;
  record: number;
  /** ASK's answers by bookmark name (lower case). */
  asked: Map<string, string>;
}

const OPERATORS = new Set(["=", "<>", "<", ">", "<=", ">="]);

/** The field's word: MERGEFIELD, IF, ... or "=" for a formula. */
export function fieldWord(code: string): string {
  const t = code.trimStart();
  if (t.startsWith("=")) return "=";
  return (/^[^\s\\{]+/.exec(t)?.[0] ?? "").toUpperCase();
}

/** The code's words and quoted texts, switches left out (MERGEFIELD Name \* Upper -> [MERGEFIELD, Name]). */
export function fieldArguments(code: string): string[] {
  const out: string[] = [];
  for (const m of code.matchAll(/\\\S\s*(?:"[^"]*"|“[^”]*”|\S+)?|"([^"]*)"|“([^”]*)”|(\S+)/g)) {
    if (m[0].startsWith("\\")) continue;
    out.push(m[1] ?? m[2] ?? m[3] ?? "");
  }
  return out;
}

/** A switch's argument (\d "default" -> default); null when the code has none. */
export function fieldSwitch(code: string, name: string): string | null {
  const m = new RegExp(`\\\\${name}\\s*(?:"([^"]*)"|“([^”]*)”|([^\\s\\\\]+))`, "i").exec(code);
  return m ? m[1] ?? m[2] ?? m[3] ?? "" : null;
}

/** A value as the code's switches show it (\# number, \@ date, \* case, \b \f): empty stays empty. */
function formatted(code: string, value: string): string {
  if (!value.trim()) return "";
  // (the switches: what is left from the first backslash outside quotes)
  // (not a nested field's switches)
  let outer = code;
  for (let guard = 0; guard < 200 && /\{[^{}]*\}/.test(outer); guard++) outer = outer.replace(/\{[^{}]*\}/, "");
  const at = outer.search(/\\(?=[*#@!bf])/i);
  const choice = parseMergeFieldCode(`MERGEFIELD X ${at >= 0 ? outer.slice(at) : ""}`);
  return mergeFieldSample(choice, value);
}

function numberOf(s: string): number | null {
  const t = s.replace(/[$%,]/g, "").trim();
  if (!t) return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

/** Word's comparison: numbers as numbers; = and <> with * or ? as a pattern; else text, case ignored. */
export function compareValues(a: string, op: string, b: string): boolean {
  let cmp: number;
  const x = numberOf(a), y = numberOf(b);
  if (x !== null && y !== null) cmp = x === y ? 0 : x < y ? -1 : 1;
  else if ((op === "=" || op === "<>") && /[*?]/.test(b)) {
    const like = new RegExp(`^${b.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".")}$`, "is").test(a);
    return op === "=" ? like : !like;
  } else {
    const p = a.toUpperCase(), q = b.toUpperCase();
    cmp = p === q ? 0 : p < q ? -1 : 1;
  }
  switch (op) {
    case "=": return cmp === 0;
    case "<>": return cmp !== 0;
    case "<": return cmp < 0;
    case ">": return cmp > 0;
    case "<=": return cmp <= 0;
    default: return cmp >= 0;
  }
}

const truthy = (v: string) => { const t = v.trim(); return t !== "" && t !== "0" && t.toLowerCase() !== "false"; };

/** The code with each nested { field } replaced by its result: in quotes (one value however many words it has),
 *  unless quotes are round it already or the field is a formula, whose numbers stay numbers. */
export function resolveNested(code: string, ctx: PreviewContext): string {
  const quote = fieldWord(code) !== "=";
  let s = code;
  for (let guard = 0; guard < 200; guard++) {
    const m = /\{([^{}]*)\}/.exec(s);
    if (!m) break;
    const value = (fieldResult(m[1].trim(), ctx) ?? "").replace(/["“”]/g, "");
    // (inside a quoted text already: an odd number of quotes before it)
    const quoted = (s.slice(0, m.index).match(/["“”]/g)?.length ?? 0) % 2 === 1;
    const put = quote && !quoted ? `"${value}"` : value;
    s = s.slice(0, m.index) + put + s.slice(m.index + m[0].length);
  }
  return s;
}

/** NEXTIF's and SKIPIF's condition: expression operator expression, or one expression (true unless empty, 0 or False). */
export function fieldCondition(code: string, ctx: PreviewContext): boolean {
  const a = fieldArguments(resolveNested(code, ctx));
  if (a.length >= 4 && OPERATORS.has(a[2])) return compareValues(a[1], a[2], a[3]);
  return truthy(a[1] ?? "");
}

/** IF's result: IF left op right "true text" "false text", or IF value "true text" "false text". */
export function ifResult(code: string): string {
  type Token = { text: string; quoted: boolean };
  const tokens: Token[] = [];
  const s = code;
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (/\s/.test(c)) { i++; continue; }
    const closers = c === '"' || c === "“" || c === "”" ? "\"”" : c === "‘" ? "’" : "";
    if (closers) {
      let j = i + 1;
      while (j < s.length && !closers.includes(s[j])) j++;
      tokens.push({ text: s.slice(i + 1, j), quoted: true });
      i = j + 1;
    } else if ("=<>".includes(c)) {
      let j = i;
      while (j < s.length && "=<>".includes(s[j])) j++;
      tokens.push({ text: s.slice(i, j), quoted: false });
      i = j;
    } else {
      let j = i;
      while (j < s.length && !/\s/.test(s[j]) && !"=<>".includes(s[j])) j++;
      tokens.push({ text: s.slice(i, j), quoted: false });
      i = j;
    }
  }
  // (the first token is IF; switches at the end are not texts)
  const args = tokens.slice(1).filter((t) => t.quoted || !t.text.startsWith("\\"));
  const [left, op] = args;
  if (op && !op.quoted && OPERATORS.has(op.text)) {
    const [, , right, yes, no] = args;
    return (compareValues(left?.text ?? "", op.text, right?.text ?? "") ? yes : no)?.text ?? "";
  }
  return (truthy(left?.text ?? "") ? args[1] : args[2])?.text ?? "";
}

/** A formula's value (numbers, + - * /, parentheses), or null when it is something else. */
export function formulaValue(expression: string): number | null {
  const s = expression.replace(/[$,]/g, "");
  let i = 0;
  const space = () => { while (s[i] === " ") i++; };
  const num = (): number | null => {
    space();
    if (s[i] === "(") { i++; const v = sum(); space(); if (s[i] !== ")") return null; i++; return v; }
    if (s[i] === "-") { i++; const v = num(); return v === null ? null : -v; }
    const m = /^\d+(?:\.\d+)?|^\.\d+/.exec(s.slice(i));
    if (!m) return null;
    i += m[0].length;
    return Number(m[0]);
  };
  const product = (): number | null => {
    let v = num();
    for (;;) {
      space();
      const op = s[i];
      if (v === null || (op !== "*" && op !== "/")) return v;
      i++;
      const w = num();
      if (w === null) return null;
      v = op === "*" ? v * w : v / w;
    }
  };
  const sum = (): number | null => {
    let v = product();
    for (;;) {
      space();
      const op = s[i];
      if (v === null || (op !== "+" && op !== "-")) return v;
      i++;
      const w = product();
      if (w === null) return null;
      v = op === "+" ? v + w : v - w;
    }
  };
  const v = sum();
  space();
  return v !== null && i === s.length && Number.isFinite(v) ? v : null;
}

/** What a field shows for the record (null: a field the preview leaves as it is). */
export function fieldResult(code: string, ctx: PreviewContext): string | null {
  const word = fieldWord(code);
  switch (word) {
    case "MERGEFIELD": {
      const a = fieldArguments(code);
      return formatted(code, a.length > 1 ? ctx.value(a[1]) : "");
    }
    case "FILLIN": {
      const a = fieldArguments(code);
      const v = a.length > 1 ? ctx.value(a[1]) : "";
      return formatted(code, v || (fieldSwitch(code, "d") ?? ""));
    }
    case "ASK": case "NEXT": case "NEXTIF": case "SKIPIF": return "";
    case "REF": {
      const a = fieldArguments(code);
      const v = a.length > 1 ? ctx.asked.get(a[1].toLowerCase()) : undefined;
      return v === undefined ? null : formatted(code, v);
    }
    case "MERGEREC": return formatted(code, String(ctx.record));
    case "MERGESEQ": return formatted(code, "1");
    case "IF": return ifResult(resolveNested(code, ctx));
    case "=": {
      const resolved = resolveNested(code, ctx).trim().slice(1);
      const at = resolved.search(/\\[#*@]/);
      const v = formulaValue(at >= 0 ? resolved.slice(0, at) : resolved);
      return v === null ? null : formatted(code, String(Math.round(v * 1e10) / 1e10));
    }
    default: return null;
  }
}
