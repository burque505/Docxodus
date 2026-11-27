/**
 * The ribbon's Merge Fields pane — Mailings > Insert Merge Field, kept open beside the document
 * as MergeFielder keeps its window beside Word: the data source (Browse…, Clear), a filter over
 * every field name (the source's columns, then the document's own merge fields), a Format box
 * (Word's switches: case, number, currency, percent, CardText, OrdText, Ordinal, DollarText,
 * Roman, dates, If True/False, a custom picture) with text before (`\b`) and after (`\f`), the code
 * and a sample of its result, and Insert (Enter, a double-click). The format chosen last for a field
 * comes back for it. The merge field at the caret is shown with its format, loaded into the Format
 * box, and "Apply to Field" changes it.
 */

export const MERGE_FIELDS_PANE_CSS = `
.dxr-mfpane { flex: 0 0 270px; display: flex; flex-direction: column; min-height: 0; overflow-x: hidden; overflow-y: auto;
  background: var(--dxr-chrome); border-left: 1px solid var(--dxr-rule); font-size: 12.5px; color: var(--dxr-ink); }
.dxr-mfpane[hidden], .dxr-mfpane [hidden] { display: none !important; }
.dxr-mfpane .dxr-mfsec { display: flex; flex-direction: column; gap: 5px; padding: 7px 10px; border-bottom: 1px solid var(--dxr-rule); }
.dxr-mfpane .dxr-mfrow { display: flex; align-items: center; gap: 6px; }
.dxr-mfpane .dxr-mfrow > input, .dxr-mfpane .dxr-mfrow > select { flex: 1 1 auto; min-width: 0; }
.dxr-mfpane input[type="text"], .dxr-mfpane select { font: inherit; font-size: 12px; box-sizing: border-box; }
.dxr-mfpane button { font: inherit; font-size: 12px; padding: 2px 9px; cursor: pointer; white-space: nowrap;
  border: 1px solid var(--dxr-rule-strong); border-radius: 4px; background: var(--dxr-chrome); color: var(--dxr-ink); }
.dxr-mfpane button:hover:not(:disabled) { background: var(--dxr-chrome-sunk); }
.dxr-mfpane button:disabled { color: var(--dxr-muted); cursor: default; }
.dxr-mfpane button.dxr-mfprimary { font-weight: 600; }
.dxr-mfpane .dxr-mfsource { flex: 1 1 auto; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-weight: 600; }
.dxr-mfpane .dxr-mfsource[data-none] { font-weight: 400; color: var(--dxr-muted); font-style: italic; }
.dxr-mfpane .dxr-mflabel { color: var(--dxr-muted); font-size: 11.5px; }
.dxr-mfpane .dxr-mflist { flex: 1 1 auto; min-height: 120px; overflow: auto; padding: 2px 0; border-bottom: 1px solid var(--dxr-rule); outline: none; }
.dxr-mfpane .dxr-mfitem { padding: 3px 10px; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; user-select: none; }
.dxr-mfpane .dxr-mfitem:hover { background: var(--dxr-chrome-sunk); }
.dxr-mfpane .dxr-mfitem[aria-selected="true"] { background: var(--dxr-wash, #e0f2fe); box-shadow: inset 3px 0 0 var(--dxr-accent); }
.dxr-mfpane .dxr-mfitem[data-doc]::after { content: attr(data-doc); color: var(--dxr-muted); font-size: 11px; margin-left: 6px; }
.dxr-mfpane .dxr-mfempty { padding: 8px 10px; color: var(--dxr-muted); font-style: italic; }
.dxr-mfpane .dxr-mfcode { font-family: Consolas, "Courier New", monospace; font-size: 11.5px; word-break: break-all;
  padding: 3px 5px; border: 1px solid var(--dxr-rule); border-radius: 4px; min-height: 1.3em; background: var(--dxr-surface, #fff); }
.dxr-mfpane textarea.dxr-mfcode { width: 100%; box-sizing: border-box; resize: vertical; min-height: 3.2em; color: var(--dxr-ink); }
.dxr-mfpane textarea.dxr-mfcode[data-edited] { border-color: var(--dxr-accent); }
.dxr-mfpane .dxr-mfresult { padding: 0 2px; min-height: 1.3em; word-break: break-word; }
.dxr-mfpane .dxr-mfat { background: var(--dxr-chrome-sunk); }
.dxr-mfpane .dxr-mferror { color: #b42318; }
.dxr[data-chrome="compact"] .dxr-mfpane { flex-basis: 230px; }
`;

/** The pane's markup: after the document in the ribbon's work area (see RIBBON_HTML). */
export const MERGE_FIELDS_PANE_HTML = `
<aside class="dxr-mfpane" data-dxr="mergefieldspane" hidden aria-label="Merge Fields">
  <div class="dxr-navhead"><span>Merge Fields</span><button type="button" data-mf="close" title="Close the Merge Fields pane" aria-label="Close">&times;</button></div>
  <div class="dxr-mfsec">
    <span class="dxr-mflabel">Data source</span>
    <div class="dxr-mfrow">
      <span class="dxr-mfsource" data-mf="source" data-none>(none)</span>
      <button type="button" data-mf="browse" title="Choose a data source (spreadsheet, CSV, JSON, Access), or a list of field names">Browse&#8230;</button>
      <button type="button" data-mf="clearsource" title="Forget this document's data source">Clear</button>
    </div>
  </div>
  <div class="dxr-mfsec">
    <div class="dxr-mfrow">
      <input type="text" data-mf="filter" placeholder="Filter" aria-label="Filter" autocomplete="off" spellcheck="false" />
      <button type="button" data-mf="clearfilter" title="Clear the filter">Clear</button>
    </div>
  </div>
  <div class="dxr-mflist" data-mf="list" role="listbox" tabindex="0" aria-label="Fields"></div>
  <div class="dxr-mfsec">
    <div class="dxr-mfrow"><span class="dxr-mflabel">Field</span><input type="text" data-mf="name" autocomplete="off" spellcheck="false" aria-label="Field name" /></div>
    <div class="dxr-mfrow"><span class="dxr-mflabel">Format</span>
      <select data-mf="format" aria-label="Format">
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
        <option value="other" hidden>Other (as in the document)</option>
      </select>
    </div>
    <div class="dxr-mfrow" data-mf-row="iftrue" hidden><span class="dxr-mflabel">If True</span><input type="text" data-mf="iftrue" aria-label="Text if True" /></div>
    <div class="dxr-mfrow" data-mf-row="iffalse" hidden><span class="dxr-mflabel">If False</span><input type="text" data-mf="iffalse" aria-label="Text if False" /></div>
    <div class="dxr-mfrow" data-mf-row="picture" hidden><span class="dxr-mflabel">Picture</span><input type="text" data-mf="picture" spellcheck="false" placeholder="0.00 or dd/MM/yyyy" aria-label="Picture" /></div>
    <div class="dxr-mfrow" data-mf-row="before"><span class="dxr-mflabel">Before</span><input type="text" data-mf="before" aria-label="Text before (\\b)" /></div>
    <div class="dxr-mfrow" data-mf-row="after"><span class="dxr-mflabel">After</span><input type="text" data-mf="after" aria-label="Text after (\\f)" /></div>
    <div class="dxr-mfrow"><span class="dxr-mflabel">Sample</span><input type="text" data-mf="sample" aria-label="Sample value" /></div>
    <span class="dxr-mflabel">Field code</span>
    <textarea class="dxr-mfcode" data-mf="code" rows="2" spellcheck="false" aria-label="Field code"
              title="The field code: change it, or type one of your own (nested fields in { }). Insert and Apply to Field use it as it is."></textarea>
    <div class="dxr-mfresult" data-mf="result" title="What the field shows for the sample"></div>
    <div class="dxr-mfrow">
      <button type="button" class="dxr-mfprimary" data-mf="insert" title="Insert the field at the caret (Enter, or double-click a name)">Insert</button>
      <span class="dxr-mferror" data-mf-msg="refused" hidden>The field cannot be put there.</span>
    </div>
  </div>
  <div class="dxr-mfsec dxr-mfat" data-mf="at" hidden>
    <span class="dxr-mflabel">Field at the caret</span>
    <div class="dxr-mfcode" data-mf="atcode"></div>
    <div class="dxr-mfrow">
      <button type="button" data-mf="apply" title="Give the field at the caret the name and format above">Apply to Field</button>
    </div>
  </div>
  <span hidden data-mf-text="none">(none)</span>
  <span hidden data-mf-text="empty">No fields. Browse for a data source, or type a name above.</span>
  <span hidden data-mf-text="indoc">in document</span>
</aside>
`;

export type MergeFieldFormat = "text" | "upper" | "lower" | "caps" | "firstcap" | "number" | "currency" | "percent"
  | "percenttext" | "cardtext" | "ordinal" | "ordtext" | "dollartext" | "roman" | "date1" | "date2" | "date3" | "date4"
  | "iftf" | "custom" | "other";

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


// ─── A field's code read back into the pane's terms ────────────────────────────────────────────

const NAME = String.raw`("(?:[^"\\]|\\.)*"|[^\s\\{}"]+)`;
const unquote = (s: string) => (s.startsWith('"') ? s.slice(1, -1).replace(/\\"/g, '"') : s);

/**
 * The pane's choice for a field's code (what {@link mergeFieldCode} makes, or Word's own of the
 * same shape); `format: "other"` for a code the pane has no format for (it is kept as it is).
 */
export function parseMergeFieldCode(code: string, fallbackName = ""): MergeFieldChoice {
  const c = code.trim().replace(/\s+/g, " ");
  const other: MergeFieldChoice = { name: fallbackName, format: "other" };
  let m = new RegExp(String.raw`^= ?\{ ?MERGEFIELD ${NAME} ?\} ?\* ?100 ?\\(#|\*) ?(.+)$`, "i").exec(c);
  if (m) {
    const name = unquote(m[1]);
    if (m[2] === "#" && unquote(m[3].trim()) === "0.00%") return { name, format: "percent" };
    if (m[2] === "*" && /^cardtext$/i.test(m[3].trim())) return { name, format: "percenttext" };
    return { ...other, name };
  }
  m = new RegExp(String.raw`^IF ?\{ ?MERGEFIELD ${NAME} ?\} ?= ?"?True"? "((?:[^"\\]|\\.)*)" "((?:[^"\\]|\\.)*)"$`, "i").exec(c);
  if (m) return { name: unquote(m[1]), format: "iftf", ifTrue: m[2].replace(/\\"/g, '"'), ifFalse: m[3].replace(/\\"/g, '"') };
  m = new RegExp(String.raw`^MERGEFIELD ${NAME}(.*)$`, "i").exec(c);
  if (!m) return other;
  const name = unquote(m[1]);
  const choice: MergeFieldChoice = { name, format: "text" };
  let rest = m[2];
  const switches = /^ ?\\([*#@bf]) ?("(?:[^"\\]|\\.)*"|[^\s\\]+)/;
  let format: string | null = null, picture: string | null = null, kind: "#" | "@" | null = null;
  for (let s = switches.exec(rest); s; s = switches.exec(rest)) {
    const value = unquote(s[2]);
    if (s[1] === "b") choice.before = value;
    else if (s[1] === "f") choice.after = value;
    else if (s[1] === "*") {
      if (value.toLowerCase() === "mergeformat") { /* (Word's own: keep the formatting; nothing to show) */ }
      else if (format) return { ...other, name };
      else format = value.toLowerCase();
    }
    else { if (picture !== null) return { ...other, name }; picture = value; kind = s[1] as "#" | "@"; }
    rest = rest.slice(s[0].length);
  }
  if (rest.trim() || (format && picture !== null)) return { ...other, name };
  if (format) {
    const byCase = (Object.keys(CASE_SWITCHES) as MergeFieldFormat[]).find((k) => CASE_SWITCHES[k]!.toLowerCase() === format);
    if (!byCase) return { ...other, name };
    choice.format = byCase;
  } else if (picture !== null) {
    if (kind === "#" && picture === "0") choice.format = "number";
    else if (kind === "#" && picture === "$,0.00") choice.format = "currency";
    else {
      const date = kind === "@" ? (Object.keys(DATE_PICTURES) as MergeFieldFormat[]).find((k) => DATE_PICTURES[k] === picture) : undefined;
      if (date) choice.format = date;
      else { choice.format = "custom"; choice.picture = picture; }
    }
  }
  return choice;
}

// ─── The pane ──────────────────────────────────────────────────────────────────────────────

export interface MergeFieldsPaneHost {
  /** The data source's columns (or a field list's names), in order. */
  columns(): string[];
  /** The data source's name to show; "" for none. */
  source(): string;
  /** The document's merge field names, in order. */
  documentFields(): string[];
  /** Insert field `code` at the document's caret. */
  insert(code: string): { ok: boolean; code?: string };
  /** The merge field at the document's caret, or null. */
  fieldAtCaret(): { code: string; name: string } | null;
  /** Change the merge field at the caret to `code`. */
  replace(code: string): { ok: boolean; code?: string };
  /** Browse… (choose a data source or a field list) and Clear: the host's. Null when it has none. */
  browse: (() => void) | null;
  clearSource: (() => void) | null;
  /** The focus goes back to the document (after an insert). */
  refocus(): void;
  close(): void;
}

export class MergeFieldsPane {
  /** The choice made last for each field (by its name, any case), for the next time it is chosen. */
  private readonly last = new Map<string, MergeFieldChoice>();
  private names: string[] = [];
  private selected = -1;
  /** The field at the caret as last shown (its code), so moving onto another loads that one. */
  private atCode: string | null = null;
  /** The code box was typed in: it is used as it is until a name or format is chosen again. */
  private codeEdited = false;
  /** The code typed for a field last (by its name, any case), when it was typed rather than chosen. */
  private readonly lastCode = new Map<string, string>();

  constructor(readonly element: HTMLElement, private readonly host: MergeFieldsPaneHost) {
    const list = this.part("list");
    this.button("close").addEventListener("click", () => host.close());
    const browse = this.button("browse"), clear = this.button("clearsource");
    browse.hidden = !host.browse;
    clear.hidden = !host.clearSource;
    browse.addEventListener("click", () => host.browse?.());
    clear.addEventListener("click", () => host.clearSource?.());
    const filter = this.input("filter");
    filter.addEventListener("input", () => { this.fill(); this.pick(0); });
    filter.addEventListener("keydown", (event) => {
      if (event.key === "ArrowDown") { event.preventDefault(); this.pick(Math.min(this.selected + 1, this.names.length - 1)); list.focus(); }
      else if (event.key === "Enter") { event.preventDefault(); this.insert(); }
      else if (event.key === "Escape" && filter.value) { event.preventDefault(); event.stopPropagation(); filter.value = ""; this.fill(); this.pick(0); }
    });
    this.button("clearfilter").addEventListener("click", () => { filter.value = ""; this.fill(); this.pick(0); filter.focus(); });
    list.addEventListener("mousedown", (event) => {
      const item = (event.target as HTMLElement).closest<HTMLElement>("[data-index]");
      if (item) this.pick(Number(item.dataset.index));
    });
    list.addEventListener("dblclick", (event) => { if ((event.target as HTMLElement).closest("[data-index]")) this.insert(); });
    list.addEventListener("keydown", (event) => {
      if (event.key === "ArrowDown") { event.preventDefault(); this.pick(Math.min(this.selected + 1, this.names.length - 1)); }
      else if (event.key === "ArrowUp") { event.preventDefault(); if (this.selected <= 0) filter.focus(); else this.pick(this.selected - 1); }
      else if (event.key === "Home") { event.preventDefault(); this.pick(0); }
      else if (event.key === "End") { event.preventDefault(); this.pick(this.names.length - 1); }
      else if (event.key === "Enter") { event.preventDefault(); this.insert(); }
    });
    this.input("name").addEventListener("input", () => { this.codeEdited = false; this.update(); });
    this.input("name").addEventListener("keydown", (event) => { if (event.key === "Enter") { event.preventDefault(); this.insert(); } });
    this.select("format").addEventListener("change", () => { this.codeEdited = false; this.update(); });
    for (const key of ["before", "after", "iftrue", "iffalse", "picture", "sample"]) {
      this.input(key).addEventListener("input", () => { if (key !== "sample") this.codeEdited = false; this.update(); });
      this.input(key).addEventListener("keydown", (event) => { if (event.key === "Enter") { event.preventDefault(); this.insert(); } });
    }
    const codeBox = this.codeBox();
    codeBox.addEventListener("input", () => { this.codeEdited = true; this.update(); });
    codeBox.addEventListener("keydown", (event) => {
      // (Enter inserts, as in the other boxes; Shift+Enter is a new line in a long code)
      if (event.key === "Enter" && !event.shiftKey) { event.preventDefault(); this.insert(); }
    });
    this.button("insert").addEventListener("click", () => this.insert());
    this.button("apply").addEventListener("click", () => this.apply());
    // (the buttons keep the document's selection: Insert and Apply act there)
    for (const button of Array.from(element.querySelectorAll<HTMLButtonElement>("button")))
      button.addEventListener("mousedown", (event) => event.preventDefault());
  }

  get visible(): boolean {
    return !this.element.hidden;
  }

  setVisible(on: boolean): void {
    this.element.hidden = !on;
    if (on) { this.refresh(); this.input("filter").focus(); }
  }

  /** The source or the document changed: list the names again. */
  refresh(): void {
    if (!this.visible) return;
    const source = this.part("source");
    const name = this.host.source();
    source.textContent = name || (this.element.querySelector("[data-mf-text=none]")?.textContent ?? "");
    source.toggleAttribute("data-none", !name);
    source.title = name;
    this.button("clearsource").disabled = !name;
    const keep = this.names[this.selected];
    this.fill();
    const again = keep ? this.names.findIndex((n) => n.toLowerCase() === keep.toLowerCase()) : -1;
    if (again >= 0) this.mark(again);
    else if (this.selected < 0 || this.selected >= this.names.length) this.pick(0);
    this.update();
    this.syncCaret();
  }

  /** The caret moved: show the merge field it is at, and load it into the Format box. */
  syncCaret(): void {
    if (!this.visible) return;
    const field = this.host.fieldAtCaret();
    const at = this.part("at");
    at.hidden = !field;
    if (!field) { this.atCode = null; return; }
    this.part("atcode").textContent = mergeFieldDisplay(field.code);
    if (field.code === this.atCode) return;
    this.atCode = field.code;
    this.load(parseMergeFieldCode(field.code, field.name), field.code);
    const index = this.names.findIndex((n) => n.toLowerCase() === field.name.toLowerCase());
    if (index >= 0) this.mark(index);
  }

  /** The choice as the pane shows it now. */
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
    const doc = this.element.ownerDocument;
    const list = this.part("list");
    const filter = this.input("filter").value.trim().toLowerCase();
    const columns = this.host.columns();
    const inDoc = this.host.documentFields().filter((n) => !columns.some((c) => c.toLowerCase() === n.toLowerCase()));
    const tag = this.element.querySelector("[data-mf-text=indoc]")?.textContent ?? "";
    const all = [...columns.map((n) => ({ n, doc: false })), ...inDoc.map((n) => ({ n, doc: true }))]
      .filter((e) => !filter || e.n.toLowerCase().includes(filter));
    this.names = all.map((e) => e.n);
    this.selected = -1;
    if (all.length === 0) {
      const empty = doc.createElement("div");
      empty.className = "dxr-mfempty";
      empty.textContent = this.element.querySelector("[data-mf-text=empty]")?.textContent ?? "";
      list.replaceChildren(empty);
      return;
    }
    list.replaceChildren(...all.map((e, i) => {
      const item = doc.createElement("div");
      item.className = "dxr-mfitem";
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", "false");
      item.dataset.index = String(i);
      if (e.doc) item.dataset.doc = tag;
      item.textContent = e.n;
      item.title = e.n;
      return item;
    }));
  }

  /** Mark row `index` without loading it. */
  private mark(index: number): void {
    const list = this.part("list");
    for (const item of Array.from(list.querySelectorAll<HTMLElement>("[data-index]")))
      item.setAttribute("aria-selected", String(Number(item.dataset.index) === index));
    this.selected = index;
    list.querySelector<HTMLElement>(`[data-index="${index}"]`)?.scrollIntoView({ block: "nearest" });
  }

  /** Choose row `index`: its name, with the format used for it last. */
  private pick(index: number): void {
    if (index < 0 || index >= this.names.length) { this.selected = -1; return; }
    this.mark(index);
    const name = this.names[index];
    this.load(this.last.get(name.toLowerCase()) ?? { name, format: "text" }, this.lastCode.get(name.toLowerCase()));
  }

  /** Show choice `c`; with `code`, the code box holds that code as it is (a field's own, or one typed before). */
  private load(c: MergeFieldChoice, code?: string): void {
    this.input("name").value = c.name;
    const other = this.select("format").querySelector<HTMLOptionElement>('option[value="other"]')!;
    other.hidden = c.format !== "other";
    this.select("format").value = c.format;
    this.input("before").value = c.before ?? "";
    this.input("after").value = c.after ?? "";
    this.input("iftrue").value = c.ifTrue ?? "";
    this.input("iffalse").value = c.ifFalse ?? "";
    this.input("picture").value = c.picture ?? "";
    this.codeEdited = code !== undefined;
    if (code !== undefined) this.codeBox().value = code;
    this.update();
  }

  private codeBox(): HTMLTextAreaElement {
    return this.element.querySelector<HTMLTextAreaElement>('textarea[data-mf="code"]')!;
  }

  /** The code box's code, without braces round the whole of it (`{ MERGEFIELD X }` is taken as typed). */
  code(): string {
    let code = this.codeBox().value.trim();
    if (code.startsWith("{") && code.endsWith("}")) {
      let depth = 0, whole = true;
      for (let i = 0; i < code.length; i++) {
        if (code[i] === "{") depth++;
        else if (code[i] === "}" && --depth === 0 && i < code.length - 1) { whole = false; break; }
      }
      if (whole) code = code.slice(1, -1).trim();
    }
    return code;
  }

  private update(): void {
    const c = this.choice();
    const row = (key: string, on: boolean) => { const el = this.element.querySelector<HTMLElement>(`[data-mf-row=${key}]`); if (el) el.hidden = !on; };
    row("iftrue", c.format === "iftf");
    row("iffalse", c.format === "iftf");
    row("picture", c.format === "custom");
    const plain = !["iftf", "percent", "percenttext", "other"].includes(c.format);
    row("before", plain);
    row("after", plain);
    const box = this.codeBox();
    if (!this.codeEdited) box.value = c.name && c.format !== "other" ? mergeFieldCode(c) : "";
    box.toggleAttribute("data-edited", this.codeEdited);
    // (the sample follows a typed code when it is one the Format box knows)
    const shown = this.codeEdited ? parseMergeFieldCode(this.code(), c.name) : c;
    this.part("result").textContent = shown.name && shown.format !== "other" ? mergeFieldSample(shown, this.input("sample").value) : "";
    const empty = !this.code();
    this.button("insert").disabled = empty;
    this.button("apply").disabled = empty;
    this.showError(false);
  }

  /** The format and code used for the field, for the next time it is chosen. */
  private remember(code: string): void {
    const c = this.choice();
    const name = (c.name || parseMergeFieldCode(code).name).toLowerCase();
    if (!name) return;
    if (c.format !== "other") this.last.set(name, c);
    if (this.codeEdited) this.lastCode.set(name, code);
    else this.lastCode.delete(name);
  }

  private insert(): void {
    const code = this.code();
    if (!code) return;
    const r = this.host.insert(code);
    if (!r.ok) { this.showError(true); return; }
    this.remember(code);
    this.host.refocus();
  }

  private apply(): void {
    const code = this.code();
    if (!code) return;
    const r = this.host.replace(code);
    if (!r.ok) { this.showError(true); return; }
    this.remember(code);
    this.atCode = null;
    this.host.refocus();
    this.syncCaret();
  }

  private showError(on: boolean): void {
    const el = this.element.querySelector<HTMLElement>("[data-mf-msg=refused]");
    if (el) el.hidden = !on;
  }

  private part(key: string): HTMLElement {
    return this.element.querySelector<HTMLElement>(`[data-mf="${key}"]`)!;
  }

  private button(key: string): HTMLButtonElement {
    return this.element.querySelector<HTMLButtonElement>(`button[data-mf="${key}"]`)!;
  }

  private input(key: string): HTMLInputElement {
    return this.element.querySelector<HTMLInputElement>(`input[data-mf="${key}"]`)!;
  }

  private select(key: string): HTMLSelectElement {
    return this.element.querySelector<HTMLSelectElement>(`select[data-mf="${key}"]`)!;
  }
}
