/**
 * The ribbon's Insert Merge Field dialog — Mailings > Insert Merge Field…: a field from the data
 * source's columns or the document's own merge fields (or a name typed), shown in one of Word's
 * formats (`\* Upper`, `\# $,0.00`, `\@ "MMMM d, yyyy"`, a percent as `{= {MERGEFIELD X} * 100}`,
 * an If True/False...), with text before (`\b`) and after (`\f`). The field code and a sample of
 * its result are shown as they are chosen; the format chosen last for a field comes back for it.
 * Not modal: Insert puts the field at the document's caret and the dialog stays for the next one.
 */
import { makeDialogMovable } from "./ribbon-paragraph-dialog.js";

export const MERGE_FIELD_DIALOG_CSS = `
.dxr-mfdialog { width: min(460px, calc(100vw - 32px)); }
.dxr-mfdialog[open]:not(:modal) { position: fixed; z-index: 40; margin: 0; box-shadow: 0 8px 28px rgba(15, 23, 42, .22); }
.dxr-mfdialog .dxr-mfbody { display: grid; grid-template-columns: 1fr 1fr; gap: 8px 12px; padding: 10px 16px 4px; }
.dxr-mfdialog label.dxr-mfcol { display: flex; flex-direction: column; gap: 4px; }
.dxr-mfdialog .dxr-mfwide { grid-column: 1 / -1; }
.dxr-mfdialog input[type="text"], .dxr-mfdialog select { width: 100%; box-sizing: border-box; }
.dxr-mfdialog select[size] { min-height: 9em; }
.dxr-mfdialog .dxr-mfcode, .dxr-mfdialog .dxr-mfresult { font-family: Consolas, "Courier New", monospace; font-size: 12px;
  padding: 4px 6px; border: 1px solid var(--dxr-line, #d0d5dd); border-radius: 4px; min-height: 1.4em; word-break: break-all; }
.dxr-mfdialog .dxr-mfresult { font-family: inherit; font-size: inherit; }
.dxr-mfdialog .dxr-mfempty { color: var(--dxr-ink-soft, #667085); font-style: italic; }
.dxr-mfdialog [hidden] { display: none !important; }
`;

export const MERGE_FIELD_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-mfdialog" data-dxr="mergefielddialog">
  <form method="dialog" novalidate>
    <h2 data-mf-title>Insert Merge Field</h2>
    <div class="dxr-mfbody">
      <label class="dxr-mfcol">Field name <input type="text" autocomplete="off" spellcheck="false" data-mf="name" /></label>
      <label class="dxr-mfcol">Format
        <select data-mf="format">
          <option value="text">Text</option>
          <option value="upper">UPPERCASE</option>
          <option value="lower">lowercase</option>
          <option value="caps">Title Case</option>
          <option value="firstcap">First capital</option>
          <option value="number">Number (0)</option>
          <option value="currency">Currency ($,0.00)</option>
          <option value="percent">Percent (0.00%)</option>
          <option value="percenttext">Percent as text</option>
          <option value="cardtext">Number as words (CardText)</option>
          <option value="ordinal">Ordinal (1st)</option>
          <option value="ordtext">Ordinal as words (OrdText)</option>
          <option value="dollartext">Dollars as words (DollarText)</option>
          <option value="roman">Roman (Roman)</option>
          <option value="date1">Date: November 27, 2026</option>
          <option value="date2">Date: 11/27/2026</option>
          <option value="date3">Date: 27 November 2026</option>
          <option value="date4">Date: 2026-11-27</option>
          <option value="iftf">If True / False</option>
          <option value="custom">Custom picture</option>
        </select>
      </label>
      <label class="dxr-mfcol dxr-mfwide">Filter <input type="text" autocomplete="off" spellcheck="false" data-mf="filter" /></label>
      <select class="dxr-mfwide" size="8" data-mf="list" aria-label="Fields"></select>
      <label class="dxr-mfcol" data-mf-row="iftrue" hidden>Text if True <input type="text" data-mf="iftrue" /></label>
      <label class="dxr-mfcol" data-mf-row="iffalse" hidden>Text if False <input type="text" data-mf="iffalse" /></label>
      <label class="dxr-mfcol dxr-mfwide" data-mf-row="picture" hidden>Picture (# 0.00, or a date such as dd/MM/yyyy) <input type="text" spellcheck="false" data-mf="picture" /></label>
      <label class="dxr-mfcol" data-mf-row="before">Text before (\\b) <input type="text" data-mf="before" /></label>
      <label class="dxr-mfcol" data-mf-row="after">Text after (\\f) <input type="text" data-mf="after" /></label>
      <label class="dxr-mfcol">Sample value <input type="text" data-mf="sample" /></label>
      <span></span>
      <span class="dxr-mfwide">Field code</span>
      <div class="dxr-mfcode dxr-mfwide" data-mf="code"></div>
      <span class="dxr-mfwide">Result</span>
      <div class="dxr-mfresult dxr-mfwide" data-mf="result"></div>
    </div>
    <div class="dxr-derror" role="alert"><span data-mf-msg="name" hidden>Type or choose a field name.</span><span data-mf-msg="refused" hidden>The field cannot be put there.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="insert" data-mf="insert">Insert</button>
      <button type="submit" value="cancel" formnovalidate>Close</button>
    </div>
    <span hidden data-mf-text="none">(no fields: type a name)</span>
  </form>
</dialog>
`;

export type MergeFieldFormat = "text" | "upper" | "lower" | "caps" | "firstcap" | "number" | "currency" | "percent"
  | "percenttext" | "cardtext" | "ordinal" | "ordtext" | "dollartext" | "roman" | "date1" | "date2" | "date3" | "date4"
  | "iftf" | "custom";

export interface MergeFieldChoice {
  name: string;
  format: MergeFieldFormat;
  before?: string;
  after?: string;
  ifTrue?: string;
  ifFalse?: string;
  picture?: string;
}

const DATE_PICTURES: Partial<Record<MergeFieldFormat, string>> = {
  date1: "MMMM d, yyyy", date2: "M/d/yyyy", date3: "d MMMM yyyy", date4: "yyyy-MM-dd",
};
const CASE_SWITCHES: Partial<Record<MergeFieldFormat, string>> = {
  upper: "Upper", lower: "Lower", caps: "Caps", firstcap: "FirstCap",
  cardtext: "CardText", ordinal: "Ordinal", ordtext: "OrdText", dollartext: "DollarText", roman: "Roman",
};

const quote = (s: string) => `"${s.replace(/"/g, '\\"')}"`;
const fieldName = (name: string) => (/[\s"\\{}]/.test(name) ? quote(name) : name);

/** The field code for a choice (no outer braces; nested fields in `{ }`, as `insertField` takes it). */
export function mergeFieldCode(c: MergeFieldChoice): string {
  const name = fieldName(c.name.trim());
  const text = (c.before ? ` \\b ${quote(c.before)}` : "") + (c.after ? ` \\f ${quote(c.after)}` : "");
  switch (c.format) {
    case "percent": return `= { MERGEFIELD ${name} } * 100 \\# "0.00%"`;
    case "percenttext": return `= { MERGEFIELD ${name} } * 100 \\* CardText`;
    case "iftf": return `IF { MERGEFIELD ${name} } = True ${quote(c.ifTrue ?? "")} ${quote(c.ifFalse ?? "")}`;
    case "number": return `MERGEFIELD ${name} \\# 0${text}`;
    case "currency": return `MERGEFIELD ${name} \\# $,0.00${text}`;
    case "custom": {
      const pic = (c.picture ?? "").trim();
      if (!pic) return `MERGEFIELD ${name}${text}`;
      return `MERGEFIELD ${name} ${/[dMyHhms]/.test(pic) ? "\\@" : "\\#"} ${quote(pic)}${text}`;
    }
    default: {
      const date = DATE_PICTURES[c.format];
      if (date) return `MERGEFIELD ${name} \\@ ${quote(date)}${text}`;
      const sw = CASE_SWITCHES[c.format];
      return `MERGEFIELD ${name}${sw ? ` \\* ${sw}` : ""}${text}`;
    }
  }
}

/** The code as Word shows it with field codes on: braces round it. */
export function mergeFieldDisplay(code: string): string {
  return `{ ${code} }`;
}

// ─── A sample of the result (a guide only: the merge itself formats the real values) ────────

const ONES = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve",
  "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"];
const TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];

function cardText(n: number): string {
  if (n < 0) return `minus ${cardText(-n)}`;
  if (n < 20) return ONES[n];
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? `-${ONES[n % 10]}` : "");
  if (n < 1000) return `${ONES[Math.floor(n / 100)]} hundred${n % 100 ? ` ${cardText(n % 100)}` : ""}`;
  for (const [size, word] of [[1e9, "billion"], [1e6, "million"], [1e3, "thousand"]] as const)
    if (n >= size) return `${cardText(Math.floor(n / size))} ${word}${n % size ? ` ${cardText(n % size)}` : ""}`;
  return String(n);
}

function ordText(n: number): string {
  const words = cardText(n);
  const special: Record<string, string> = { one: "first", two: "second", three: "third", five: "fifth", eight: "eighth", nine: "ninth", twelve: "twelfth" };
  return words.replace(/(\w+)$/, (w) => special[w] ?? (w.endsWith("y") ? `${w.slice(0, -1)}ieth` : `${w}th`));
}

function ordinal(n: number): string {
  const t = n % 100;
  const s = t >= 11 && t <= 13 ? "th" : ["th", "st", "nd", "rd"][n % 10] ?? "th";
  return `${n}${s}`;
}

function roman(n: number): string {
  if (n <= 0 || n >= 4000) return String(n);
  let out = "";
  for (const [v, r] of [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]] as const)
    while (n >= v) { out += r; n -= v; }
  return out;
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function formatDate(d: Date, pic: string): string {
  return pic.replace(/MMMM|MMM|MM|M|dd|d|yyyy|yy/g, (t) => {
    switch (t) {
      case "MMMM": return MONTHS[d.getMonth()];
      case "MMM": return MONTHS[d.getMonth()].slice(0, 3);
      case "MM": return String(d.getMonth() + 1).padStart(2, "0");
      case "M": return String(d.getMonth() + 1);
      case "dd": return String(d.getDate()).padStart(2, "0");
      case "d": return String(d.getDate());
      case "yyyy": return String(d.getFullYear());
      default: return String(d.getFullYear()).slice(-2);
    }
  });
}

function formatNumber(n: number, pic: string): string {
  const pct = pic.includes("%");
  const body = pic.replace(/[$%]/g, "");
  const decimals = (body.split(".")[1] ?? "").replace(/[^0#]/g, "").length;
  let s = n.toFixed(decimals);
  if (body.includes(",")) { const [i, f] = s.split("."); s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (f ? `.${f}` : ""); }
  return (pic.includes("$") ? "$" : "") + s + (pct ? "%" : "");
}

/** What the field would show for `value` (`«Name»` without one). */
export function mergeFieldSample(c: MergeFieldChoice, value: string): string {
  const v = value.trim();
  if (!v) return c.format === "iftf" ? `${c.ifTrue ?? ""} / ${c.ifFalse ?? ""}` : `«${c.name.trim()}»`;
  const n = Number(v.replace(/[$,]/g, ""));
  const num = Number.isFinite(n);
  const wrap = (s: string) => (s ? `${c.before ?? ""}${s}${c.after ?? ""}` : s);
  switch (c.format) {
    case "upper": return wrap(v.toUpperCase());
    case "lower": return wrap(v.toLowerCase());
    case "caps": return wrap(v.toLowerCase().replace(/(^|\s)\S/g, (m) => m.toUpperCase()));
    case "firstcap": return wrap(v.charAt(0).toUpperCase() + v.slice(1));
    case "number": return wrap(num ? formatNumber(n, "0") : v);
    case "currency": return wrap(num ? formatNumber(n, "$,0.00") : v);
    case "percent": return num ? formatNumber(n * 100, "0.00%") : v;
    case "percenttext": return num ? cardText(Math.round(n * 100)) : v;
    case "cardtext": return wrap(num ? cardText(Math.round(n)) : v);
    case "ordinal": return wrap(num ? ordinal(Math.round(n)) : v);
    case "ordtext": return wrap(num ? ordText(Math.round(n)) : v);
    case "dollartext": {
      if (!num) return wrap(v);
      const whole = Math.floor(n), cents = Math.round((n - whole) * 100);
      return wrap(`${cardText(whole)} and ${String(cents).padStart(2, "0")}/100`);
    }
    case "roman": return wrap(num ? roman(Math.round(n)) : v);
    case "iftf": return /^(true|yes|1)$/i.test(v) ? c.ifTrue ?? "" : c.ifFalse ?? "";
    case "custom": {
      const pic = (c.picture ?? "").trim();
      if (!pic) return wrap(v);
      if (/[dMy]/.test(pic)) { const d = new Date(v); return wrap(isNaN(d.getTime()) ? v : formatDate(d, pic)); }
      return wrap(num ? formatNumber(n, pic) : v);
    }
    default: {
      const pic = DATE_PICTURES[c.format];
      if (pic) { const d = new Date(v); return wrap(isNaN(d.getTime()) ? v : formatDate(d, pic)); }
      return wrap(v);
    }
  }
}

export interface MergeFieldDialogHost {
  /** The names to list: the data source's columns, then the document's merge fields. */
  names(): string[];
  /** Insert the field `code` at the document's caret. */
  insert(code: string): { ok: boolean; code?: string };
  /** The dialog closed: the focus goes back to the document. */
  closed(): void;
}

/** The Insert Merge Field dialog over its markup (`MERGE_FIELD_DIALOG_HTML`, already in the page). */
export class MergeFieldDialog {
  /** The choice made last for each field (by its name, any case), for the next time it is chosen. */
  private readonly last = new Map<string, MergeFieldChoice>();

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: MergeFieldDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-mf-title]")!);
    const form = dialog.querySelector("form")!;
    const list = this.select("list");
    this.input("filter").addEventListener("input", () => this.fill());
    list.addEventListener("change", () => { if (list.value) this.choose(list.value); });
    list.addEventListener("dblclick", () => { if (list.value) this.insert(); });
    this.input("name").addEventListener("input", () => { this.showError(null); this.update(); });
    this.select("format").addEventListener("change", () => this.update());
    for (const key of ["before", "after", "iftrue", "iffalse", "picture", "sample"])
      this.input(key).addEventListener("input", () => this.update());
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter && submitter.value !== "insert") { this.close(); return; }
      this.insert();
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.close(); });
    dialog.addEventListener("keydown", (event) => {
      if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); this.close(); }
    });
    for (const button of Array.from(dialog.querySelectorAll<HTMLButtonElement>("button")))
      button.addEventListener("mousedown", (event) => event.preventDefault());
  }

  get open(): boolean {
    return this.dialog.open;
  }

  /** Open beside the document, on field `name` when given. */
  show(name = ""): void {
    if (!this.dialog.open) {
      this.input("filter").value = "";
      this.dialog.show();
      const view = this.dialog.ownerDocument.defaultView;
      const box = this.dialog.getBoundingClientRect();
      this.dialog.style.left = `${Math.max(8, ((view?.innerWidth ?? 1024) - box.width) / 2)}px`;
      this.dialog.style.top = `${Math.max(8, ((view?.innerHeight ?? 768) - box.height) / 3)}px`;
    }
    this.fill();
    if (name) this.choose(name);
    else this.update();
    this.input(name ? "sample" : "name").focus();
  }

  close(): void {
    if (!this.dialog.open) return;
    this.dialog.close();
    this.host.closed();
  }

  /** The choice as the dialog shows it now. */
  choice(): MergeFieldChoice {
    return {
      name: this.input("name").value.trim(),
      format: this.select("format").value as MergeFieldFormat,
      before: this.input("before").value,
      after: this.input("after").value,
      ifTrue: this.input("iftrue").value,
      ifFalse: this.input("iffalse").value,
      picture: this.input("picture").value,
    };
  }

  private fill(): void {
    const list = this.select("list");
    const filter = this.input("filter").value.trim().toLowerCase();
    const names = this.host.names().filter((n) => !filter || n.toLowerCase().includes(filter));
    const doc = this.dialog.ownerDocument;
    if (names.length === 0) {
      const empty = doc.createElement("option");
      empty.disabled = true;
      empty.className = "dxr-mfempty";
      empty.textContent = this.dialog.querySelector("[data-mf-text=none]")?.textContent ?? "";
      list.replaceChildren(empty);
      return;
    }
    const current = this.input("name").value.trim().toLowerCase();
    list.replaceChildren(...names.map((n) => {
      const o = doc.createElement("option");
      o.value = n;
      o.textContent = n;
      if (n.toLowerCase() === current) o.selected = true;
      return o;
    }));
  }

  private choose(name: string): void {
    this.input("name").value = name;
    const last = this.last.get(name.toLowerCase());
    this.select("format").value = last?.format ?? "text";
    this.input("before").value = last?.before ?? "";
    this.input("after").value = last?.after ?? "";
    this.input("iftrue").value = last?.ifTrue ?? "";
    this.input("iffalse").value = last?.ifFalse ?? "";
    this.input("picture").value = last?.picture ?? "";
    this.showError(null);
    this.update();
  }

  private update(): void {
    const c = this.choice();
    const row = (key: string, on: boolean) => { const el = this.dialog.querySelector<HTMLElement>(`[data-mf-row=${key}]`); if (el) el.hidden = !on; };
    row("iftrue", c.format === "iftf");
    row("iffalse", c.format === "iftf");
    row("picture", c.format === "custom");
    const plain = !["iftf", "percent", "percenttext"].includes(c.format);
    row("before", plain);
    row("after", plain);
    const code = this.dialog.querySelector<HTMLElement>("[data-mf=code]")!;
    const result = this.dialog.querySelector<HTMLElement>("[data-mf=result]")!;
    code.textContent = c.name ? mergeFieldDisplay(mergeFieldCode(c)) : "";
    result.textContent = c.name ? mergeFieldSample(c, this.input("sample").value) : "";
    (this.dialog.querySelector("[data-mf=insert]") as HTMLButtonElement).disabled = !c.name;
  }

  private insert(): void {
    const c = this.choice();
    if (!c.name) { this.showError("name"); this.input("name").focus(); return; }
    const r = this.host.insert(mergeFieldCode(c));
    if (!r.ok) { this.showError("refused"); return; }
    this.last.set(c.name.toLowerCase(), c);
    this.showError(null);
  }

  private showError(which: "name" | "refused" | null): void {
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-mf-msg]"))) el.hidden = el.dataset.mfMsg !== which;
  }

  private input(key: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`input[data-mf="${key}"]`)!;
  }

  private select(key: string): HTMLSelectElement {
    return this.dialog.querySelector<HTMLSelectElement>(`select[data-mf="${key}"]`)!;
  }
}
