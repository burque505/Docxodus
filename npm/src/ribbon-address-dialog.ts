/**
 * The ribbon's Address Block and Greeting Line dialogs — Mailings > Address Block… / Greeting
 * Line…, as Word has them, but writing plain MERGEFIELDs (a field a line part) rather than Word's
 * ADDRESSBLOCK / GREETINGLINE fields, so the merge and any other program fill them the same way.
 * Word's address parts (Title, First Name, Last Name, Company, Address 1, City, Postal Code…) are
 * matched to the data source's columns by name, with the columns of one party taken together when
 * a document's columns carry a prefix (LeFull, LeStreet, LeCity for a lender; BoFull… for a
 * borrower). Match Fields shows and changes each match; the preview steps through the data
 * source's first rows, an empty line left out as the merge leaves it out.
 */
import { makeDialogMovable } from "./ribbon-paragraph-dialog.js";
import { mergeFieldCode } from "./ribbon-mergefields-pane.js";

/** Word's address parts, with the names a column may have for each (compared without case, spaces or punctuation). */
export const ADDRESS_PARTS = [
  { key: "title", label: "Courtesy Title", names: ["title", "courtesytitle", "honorific", "salutation", "mrmrs"] },
  { key: "first", label: "First Name", names: ["first", "firstname", "givenname", "given", "fname", "forename"] },
  { key: "middle", label: "Middle Name", names: ["middle", "middlename", "mi", "middleinitial"] },
  { key: "last", label: "Last Name", names: ["last", "lastname", "surname", "familyname", "lname"] },
  { key: "suffix", label: "Suffix", names: ["suffix", "namesuffix"] },
  { key: "full", label: "Full Name", names: ["full", "fullname", "name", "displayname", "legalname"] },
  { key: "short", label: "Short Name", names: ["short", "shortname", "nickname", "nick", "greetingname", "informalname"] },
  { key: "company", label: "Company", names: ["company", "companyname", "organization", "organisation", "org", "business", "firm", "entity"] },
  { key: "address1", label: "Address 1", names: ["address", "address1", "addressline1", "street", "streetaddress", "street1", "addr", "addr1"] },
  { key: "address2", label: "Address 2", names: ["address2", "addressline2", "street2", "addr2", "suite", "unit", "apt", "apartment"] },
  { key: "city", label: "City", names: ["city", "town", "locality"] },
  { key: "state", label: "State", names: ["state", "province", "region", "st"] },
  { key: "postal", label: "Postal Code", names: ["zip", "zipcode", "postal", "postalcode", "postcode", "zip5"] },
  { key: "country", label: "Country or Region", names: ["country", "countryregion", "nation"] },
] as const;

export type AddressPart = (typeof ADDRESS_PARTS)[number]["key"];
export type AddressMatch = Partial<Record<AddressPart, string>>;

const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");

/** The part a column names after `prefix` (a column "LeCity" with prefix "Le" names the city), else null. */
function partOf(column: string, prefix: string): AddressPart | null {
  if (prefix && !column.toLowerCase().startsWith(prefix.toLowerCase())) return null;
  const rest = norm(column.slice(prefix.length));
  if (!rest) return null;
  for (const part of ADDRESS_PARTS) if ((part.names as readonly string[]).includes(rest)) return part.key;
  return null;
}

/**
 * The prefixes a document's columns use for whole parties: the start of a column name that is
 * followed by an address part's name ("Le" for LeFull, LeStreet, LeCity), when two or more columns
 * share it. "" (no prefix) first when the columns have plain names too.
 */
export function addressPrefixes(columns: string[]): string[] {
  const counts = new Map<string, number>();
  for (const column of columns) {
    const n = column.length;
    for (let i = 0; i < n; i++) {
      const prefix = column.slice(0, i);
      if (i > 0 && /[^A-Za-z0-9]/.test(prefix)) break;
      if (partOf(column, prefix)) counts.set(prefix, (counts.get(prefix) ?? 0) + 1);
    }
  }
  const found = [...counts].filter(([, n]) => n >= 2).map(([p]) => p);
  return found.sort((a, b) => (a === "" ? -1 : b === "" ? 1 : columns.findIndex((c) => c.startsWith(a)) - columns.findIndex((c) => c.startsWith(b))));
}

/** Each address part's column, for the columns that start with `prefix` ("" any). */
export function matchAddress(columns: string[], prefix: string): AddressMatch {
  const match: AddressMatch = {};
  for (const column of columns) {
    const part = partOf(column, prefix);
    if (part && !match[part]) match[part] = column;
  }
  return match;
}

export type NameFormat = "full" | "firstlast" | "titlefirstlast" | "titlelast" | "first" | "short" | "all";
export const NAME_FORMATS: Array<{ value: NameFormat; label: string; parts: AddressPart[] }> = [
  { value: "full", label: "Full Name", parts: ["full"] },
  { value: "firstlast", label: "First Last", parts: ["first", "last"] },
  { value: "titlefirstlast", label: "Title First Last", parts: ["title", "first", "last"] },
  { value: "all", label: "Title First Middle Last Suffix", parts: ["title", "first", "middle", "last", "suffix"] },
  { value: "titlelast", label: "Title Last", parts: ["title", "last"] },
  { value: "first", label: "First", parts: ["first"] },
  { value: "short", label: "Short Name", parts: ["short"] },
];

/** The texts the dialog writes from code (the address parts and name formats), for a host's translation tables. */
export const ADDRESS_DIALOG_TEXTS: readonly string[] = [...ADDRESS_PARTS.map((p) => p.label), ...NAME_FORMATS.map((f) => f.label)];

/** A line of the block: fields (by column) and literal text between them. */
export type LinePiece = { field: string } | { text: string };
export type Line = LinePiece[];

/** The name's pieces in `format`: the parts matched (a format whose parts none match falls back to Full, then First Last). */
function namePieces(match: AddressMatch, format: NameFormat): Line {
  const pick = (f: NameFormat) => NAME_FORMATS.find((x) => x.value === f)!.parts.filter((p) => match[p]);
  let parts = pick(format);
  if (!parts.length) parts = pick("full");
  if (!parts.length) parts = pick("firstlast");
  if (!parts.length) parts = pick("short");
  const line: Line = [];
  parts.forEach((p, i) => { if (i) line.push({ text: " " }); line.push({ field: match[p]! }); });
  return line;
}

export interface AddressOptions {
  name: NameFormat;
  company: boolean;
  address: boolean;
  country: boolean;
}

/** The address block's lines, one a paragraph. */
export function addressLines(match: AddressMatch, o: AddressOptions): Line[] {
  const lines: Line[] = [];
  const name = namePieces(match, o.name);
  if (name.length) lines.push(name);
  if (o.company && match.company) lines.push([{ field: match.company }]);
  if (o.address) {
    if (match.address1) lines.push([{ field: match.address1 }]);
    if (match.address2) lines.push([{ field: match.address2 }]);
    const place: Line = [];
    if (match.city) place.push({ field: match.city });
    if (match.state) { if (place.length) place.push({ text: ", " }); place.push({ field: match.state }); }
    if (match.postal) { if (place.length) place.push({ text: match.state ? " " : ", " }); place.push({ field: match.postal }); }
    if (place.length) lines.push(place);
  }
  if (o.country && match.country) lines.push([{ field: match.country }]);
  return lines;
}

export interface GreetingOptions {
  salutation: string;
  name: NameFormat;
  punctuation: string;
}

/** The greeting line: "Dear «First»," and the like. */
export function greetingLine(match: AddressMatch, o: GreetingOptions): Line {
  const line: Line = [];
  const name = namePieces(match, o.name);
  if (o.salutation.trim()) line.push({ text: o.salutation.trim() + (name.length ? " " : "") });
  line.push(...name);
  if (o.punctuation) line.push({ text: o.punctuation });
  return line;
}

const escapeHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function lineHtml(line: Line): string {
  return line.map((piece) => "field" in piece
    ? `<span data-field-code="${escapeHtml(mergeFieldCode({ name: piece.field, format: "text" }))}">«${escapeHtml(piece.field)}»</span>`
    : escapeHtml(piece.text)).join("");
}

/** The HTML that pastes the lines as paragraphs with no space between them, the fields as fields (see paste-fields.ts). */
export function linesHtml(lines: Line[]): string {
  return `<html><body>${lines.map((l) => `<p style="margin:0">${lineHtml(l)}</p>`).join("")}</body></html>`;
}

/** The HTML that pastes one line into the paragraph at the caret. */
export function inlineHtml(line: Line): string {
  return `<html><body><span>${lineHtml(line)}</span></body></html>`;
}

/** A line as one record shows it ("" when all its fields are empty, as the merge leaves such a line out). */
export function lineText(line: Line, row: ((column: string) => string) | null): string {
  if (!row) return line.map((p) => ("field" in p ? `«${p.field}»` : p.text)).join("");
  const fields = line.filter((p): p is { field: string } => "field" in p);
  if (fields.length && fields.every((p) => !row(p.field).trim()) && line.every((p) => "field" in p || !p.text.trim() || /^[\s,]+$/.test(p.text))) return "";
  return line.map((p) => ("field" in p ? row(p.field) : p.text)).join("");
}

// ─── The dialog ────────────────────────────────────────────────────────────

export const ADDRESS_DIALOG_CSS = `
.dxr-addrdialog { width: min(620px, calc(100vw - 32px)); }
.dxr-addrdialog .dxr-addrbody { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 10px 16px; padding: 10px 16px 4px; }
.dxr-addrdialog .dxr-addropts { display: flex; flex-direction: column; gap: 7px; }
.dxr-addrdialog label { display: flex; flex-direction: column; gap: 3px; }
.dxr-addrdialog label.dxr-dcheck { flex-direction: row; align-items: center; gap: 6px; }
.dxr-addrdialog select, .dxr-addrdialog input[type="text"] { width: 100%; box-sizing: border-box; font: inherit; }
.dxr-addrdialog .dxr-addrrow { display: flex; gap: 8px; }
.dxr-addrdialog .dxr-addrrow > label { flex: 1 1 0; min-width: 0; }
.dxr-addrdialog .dxr-addrpreview { border: 1px solid var(--dxr-rule); border-radius: 4px; padding: 8px 10px; min-height: 7.5em;
  white-space: pre-wrap; font-family: Calibri, "Segoe UI", sans-serif; background: var(--dxr-surface, #fff); color: var(--dxr-ink); }
.dxr-addrdialog .dxr-addrnav { display: flex; align-items: center; gap: 6px; margin-top: 4px; }
.dxr-addrdialog .dxr-addrnav button { min-width: 2em; }
.dxr-addrdialog details { grid-column: 1 / -1; border-top: 1px solid var(--dxr-rule); padding-top: 6px; }
.dxr-addrdialog summary { cursor: pointer; font-weight: 600; }
.dxr-addrdialog .dxr-addrmatch { display: grid; grid-template-columns: max-content minmax(0, 1fr) max-content minmax(0, 1fr); gap: 4px 8px; align-items: center; margin-top: 6px; }
.dxr-addrdialog .dxr-addrnote { color: var(--dxr-muted); font-size: 11.5px; }
.dxr-addrdialog [hidden] { display: none !important; }
`;

export const ADDRESS_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-addrdialog" data-dxr="addressdialog">
  <form method="dialog" novalidate>
    <h2 data-ad-title><span data-ad-mode="address">Insert Address Block</span><span data-ad-mode="greeting">Insert Greeting Line</span></h2>
    <div class="dxr-addrbody">
      <div class="dxr-addropts">
        <label data-ad-row="prefix">Columns of
          <select data-ad="prefix" title="The party whose columns to use, when the columns carry a prefix (LeFull, LeCity... BoFull...)"></select>
        </label>
        <div data-ad-mode="greeting" class="dxr-addrrow">
          <label>Salutation <input type="text" data-ad="salutation" list="dxr-ad-salutations" value="Dear" /></label>
          <label>Punctuation <select data-ad="punctuation"><option value=",">,</option><option value=":">:</option><option value="">(none)</option></select></label>
        </div>
        <datalist id="dxr-ad-salutations"><option value="Dear"></option><option value="To"></option><option value="Hello"></option></datalist>
        <label>Recipient's name <select data-ad="name"></select></label>
        <div data-ad-mode="address" class="dxr-addropts">
          <label class="dxr-dcheck"><input type="checkbox" data-ad="company" checked /> Insert company name</label>
          <label class="dxr-dcheck"><input type="checkbox" data-ad="address" checked /> Insert postal address</label>
          <label class="dxr-dcheck"><input type="checkbox" data-ad="country" /> Insert country or region</label>
        </div>
        <span class="dxr-addrnote">Plain merge fields, a line each: a line whose fields are all empty is left out when merging.</span>
      </div>
      <div>
        <span>Preview</span>
        <div class="dxr-addrpreview" data-ad="preview" aria-live="polite"></div>
        <div class="dxr-addrnav" data-ad="nav">
          <button type="button" data-ad="prev" title="Previous recipient" aria-label="Previous recipient">&#9664;</button>
          <span data-ad="record"></span>
          <button type="button" data-ad="next" title="Next recipient" aria-label="Next recipient">&#9654;</button>
        </div>
      </div>
      <details data-ad="matchbox">
        <summary>Match Fields</summary>
        <div class="dxr-addrmatch" data-ad="match"></div>
      </details>
    </div>
    <div class="dxr-derror" role="alert"><span data-ad-msg="nothing" hidden>No column matches a name or an address part: choose the columns under Match Fields.</span><span data-ad-msg="refused" hidden>It cannot be put there.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok" data-ad="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
    <span hidden data-ad-text="none">(not matched)</span>
    <span hidden data-ad-text="allcolumns">(all columns)</span>
    <span hidden data-ad-text="of">{0} of {1}</span>
  </form>
</dialog>
`;

export interface AddressDialogHost {
  columns(): string[];
  /** The data source's first rows (cells in the columns' order), for the preview. */
  rows(): string[][];
  /** A text in the ribbon's language (see ADDRESS_DIALOG_TEXTS). */
  translate(text: string): string;
  /** Paste the block's HTML at the caret. */
  insert(html: string): Promise<boolean>;
  closed(): void;
}

export class AddressDialog {
  private mode: "address" | "greeting" = "address";
  private match: AddressMatch = {};
  private record = 0;
  /** The choices made last, for the next time (each mode its own). */
  private readonly kept: Record<"address" | "greeting", { prefix?: string; name?: NameFormat; match?: AddressMatch }> = { address: {}, greeting: {} };

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: AddressDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-ad-title]")!);
    this.select("prefix").addEventListener("change", () => { this.match = matchAddress(this.host.columns(), this.select("prefix").value); this.fillMatch(); this.update(); });
    for (const el of Array.from(dialog.querySelectorAll<HTMLElement>('[data-ad="name"], [data-ad="punctuation"], [data-ad="company"], [data-ad="address"], [data-ad="country"]')))
      el.addEventListener("change", () => this.update());
    this.input("salutation").addEventListener("input", () => this.update());
    this.button("prev").addEventListener("click", () => { this.record = Math.max(0, this.record - 1); this.update(); });
    this.button("next").addEventListener("click", () => { this.record = Math.min(this.host.rows().length - 1, this.record + 1); this.update(); });
    const form = dialog.querySelector("form")!;
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter && submitter.value !== "ok") { this.close(); return; }
      void this.ok();
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.close(); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  show(mode: "address" | "greeting"): void {
    this.mode = mode;
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-ad-mode]"))) el.hidden = el.dataset.adMode !== mode;
    // (written each time: the ribbon's language may have changed)
    this.select("name").replaceChildren(...NAME_FORMATS.map((f) => new Option(this.host.translate(f.label), f.value)));
    const columns = this.host.columns();
    const prefixes = addressPrefixes(columns);
    const prefix = this.select("prefix");
    const all = this.dialog.querySelector("[data-ad-text=allcolumns]")?.textContent ?? "";
    prefix.replaceChildren(...(prefixes.length ? prefixes : [""]).map((p) => new Option(p ? `${p}…` : all, p)));
    this.row("prefix", prefixes.length > 1 || (prefixes.length === 1 && prefixes[0] !== ""));
    const kept = this.kept[mode];
    prefix.value = kept.prefix !== undefined && Array.from(prefix.options).some((o) => o.value === kept.prefix) ? kept.prefix : prefix.options[0]?.value ?? "";
    const auto = matchAddress(columns, prefix.value);
    // (a match kept from last time, while its columns are still there)
    this.match = kept.match && kept.prefix === prefix.value && Object.values(kept.match).every((c) => !c || columns.includes(c)) ? { ...kept.match } : auto;
    this.select("name").value = kept.name ?? (mode === "greeting" ? (this.match.first ? "first" : this.match.short ? "short" : "full") : (this.match.full ? "full" : "titlefirstlast"));
    this.record = 0;
    this.fillMatch();
    this.update();
    if (!this.dialog.open) this.dialog.showModal();
    this.select("name").focus();
  }

  close(): void {
    if (!this.dialog.open) return;
    this.dialog.close();
    this.host.closed();
  }

  /** The lines the dialog would insert now. */
  lines(): Line[] {
    if (this.mode === "greeting") {
      const line = greetingLine(this.match, { salutation: this.input("salutation").value, name: this.select("name").value as NameFormat, punctuation: this.select("punctuation").value });
      return line.some((p) => "field" in p) ? [line] : [];
    }
    return addressLines(this.match, {
      name: this.select("name").value as NameFormat,
      company: this.input("company").checked, address: this.input("address").checked, country: this.input("country").checked,
    });
  }

  private fillMatch(): void {
    const box = this.part("match");
    const doc = this.dialog.ownerDocument;
    const none = this.dialog.querySelector("[data-ad-text=none]")?.textContent ?? "";
    const columns = this.host.columns();
    box.replaceChildren(...ADDRESS_PARTS.flatMap((part) => {
      const label = doc.createElement("span");
      label.textContent = this.host.translate(part.label);
      const select = doc.createElement("select");
      select.dataset.part = part.key;
      select.append(new Option(none, ""), ...columns.map((c) => new Option(c, c)));
      select.value = this.match[part.key] ?? "";
      select.addEventListener("change", () => {
        if (select.value) this.match[part.key] = select.value;
        else delete this.match[part.key];
        this.update();
      });
      return [label, select];
    }));
  }

  private update(): void {
    const rows = this.host.rows();
    const columns = this.host.columns();
    const row = rows.length ? (column: string) => rows[this.record]?.[columns.indexOf(column)] ?? "" : null;
    const lines = this.lines();
    this.part("preview").textContent = lines.map((l) => lineText(l, row)).filter((t, i) => !row || t !== "" || i === 0).join("\n");
    this.part("nav").hidden = rows.length === 0;
    const of = this.dialog.querySelector("[data-ad-text=of]")?.textContent ?? "{0} of {1}";
    this.part("record").textContent = of.replace("{0}", String(this.record + 1)).replace("{1}", String(rows.length));
    this.button("prev").disabled = this.record <= 0;
    this.button("next").disabled = this.record >= rows.length - 1;
    this.button("ok").disabled = lines.length === 0;
    this.showError(lines.length === 0 ? "nothing" : null);
  }

  private async ok(): Promise<void> {
    const lines = this.lines();
    if (!lines.length) return;
    const html = this.mode === "greeting" ? inlineHtml(lines[0]) : linesHtml(lines);
    this.kept[this.mode] = { prefix: this.select("prefix").value, name: this.select("name").value as NameFormat, match: { ...this.match } };
    this.dialog.close();
    const ok = await this.host.insert(html);
    if (!ok) { this.dialog.showModal(); this.showError("refused"); return; }
    this.host.closed();
  }

  private showError(which: "nothing" | "refused" | null): void {
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-ad-msg]"))) el.hidden = el.dataset.adMsg !== which;
  }

  private row(key: string, on: boolean): void {
    const el = this.dialog.querySelector<HTMLElement>(`[data-ad-row="${key}"]`);
    if (el) el.hidden = !on;
  }

  private part(key: string): HTMLElement {
    return this.dialog.querySelector<HTMLElement>(`[data-ad="${key}"]`)!;
  }

  private button(key: string): HTMLButtonElement {
    return this.dialog.querySelector<HTMLButtonElement>(`button[data-ad="${key}"]`)!;
  }

  private input(key: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`input[data-ad="${key}"]`)!;
  }

  private select(key: string): HTMLSelectElement {
    return this.dialog.querySelector<HTMLSelectElement>(`select[data-ad="${key}"]`)!;
  }
}
