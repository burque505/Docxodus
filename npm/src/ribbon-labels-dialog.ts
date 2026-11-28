/**
 * The ribbon's Labels dialog — Mailings > Labels…: a sheet of labels of a standard product (Avery
 * 5160, 5163, L7160) or a custom one, as Word lays them out: a table of fixed rows and columns, a
 * narrow empty column between labels across, each label the text typed (the Address Block's lines
 * for a party, «Name» a merge field). For a merge, every label after the first starts with Next
 * Record, so each label takes the next recipient; or every label is the same. The document becomes
 * the sheet when it is empty; else the sheet is put in front of it as a section of its own.
 */
import { makeDialogMovable } from "./ribbon-paragraph-dialog.js";
import { addressLines, addressPrefixes, matchAddress, type Line } from "./ribbon-address-dialog.js";
import { lineHtml, linesOf, textOf } from "./ribbon-envelope-dialog.js";

const IN = 1440;
const MM = 1440 / 25.4;

/** A label sheet: the page, a label's size, how many across and down, the sheet's top and side
 *  margins and the pitch (a label's start to the next one's), all twips. */
export interface LabelProduct {
  key: string;
  label: string;
  pageWidthTwips: number;
  pageHeightTwips: number;
  widthTwips: number;
  heightTwips: number;
  across: number;
  down: number;
  topTwips: number;
  sideTwips: number;
  hPitchTwips: number;
  vPitchTwips: number;
}

const LETTER = { pageWidthTwips: 12240, pageHeightTwips: 15840 };
const A4 = { pageWidthTwips: 11906, pageHeightTwips: 16838 };

/** The products Word lists for these numbers (their makers' published sizes). */
export const LABEL_PRODUCTS: LabelProduct[] = [
  { key: "5160", label: "Avery 5160 Address (1 x 2 5/8 in, 30 per sheet)", ...LETTER, widthTwips: Math.round(2.625 * IN), heightTwips: IN,
    across: 3, down: 10, topTwips: Math.round(0.5 * IN), sideTwips: Math.round(0.1875 * IN), hPitchTwips: Math.round(2.75 * IN), vPitchTwips: IN },
  { key: "5161", label: "Avery 5161 Address (1 x 4 in, 20 per sheet)", ...LETTER, widthTwips: 4 * IN, heightTwips: IN,
    across: 2, down: 10, topTwips: Math.round(0.5 * IN), sideTwips: Math.round(0.15625 * IN), hPitchTwips: Math.round(4.1875 * IN), vPitchTwips: IN },
  { key: "5162", label: "Avery 5162 Address (1 1/3 x 4 in, 14 per sheet)", ...LETTER, widthTwips: 4 * IN, heightTwips: Math.round(1.3333 * IN),
    across: 2, down: 7, topTwips: Math.round(0.8333 * IN), sideTwips: Math.round(0.15625 * IN), hPitchTwips: Math.round(4.1875 * IN), vPitchTwips: Math.round(1.3333 * IN) },
  { key: "5163", label: "Avery 5163 Shipping (2 x 4 in, 10 per sheet)", ...LETTER, widthTwips: 4 * IN, heightTwips: 2 * IN,
    across: 2, down: 5, topTwips: Math.round(0.5 * IN), sideTwips: Math.round(0.15625 * IN), hPitchTwips: Math.round(4.1875 * IN), vPitchTwips: 2 * IN },
  { key: "5167", label: "Avery 5167 Return Address (1/2 x 1 3/4 in, 80 per sheet)", ...LETTER, widthTwips: Math.round(1.75 * IN), heightTwips: Math.round(0.5 * IN),
    across: 4, down: 20, topTwips: Math.round(0.5 * IN), sideTwips: Math.round(0.3 * IN), hPitchTwips: Math.round(2.05 * IN), vPitchTwips: Math.round(0.5 * IN) },
  { key: "l7160", label: "Avery L7160 Address (63.5 x 38.1 mm, 21 per sheet)", ...A4, widthTwips: Math.round(63.5 * MM), heightTwips: Math.round(38.1 * MM),
    across: 3, down: 7, topTwips: Math.round(15.15 * MM), sideTwips: Math.round(7.21 * MM), hPitchTwips: Math.round(66.04 * MM), vPitchTwips: Math.round(38.1 * MM) },
  { key: "l7163", label: "Avery L7163 Address (99.1 x 38.1 mm, 14 per sheet)", ...A4, widthTwips: Math.round(99.1 * MM), heightTwips: Math.round(38.1 * MM),
    across: 2, down: 7, topTwips: Math.round(15.15 * MM), sideTwips: Math.round(4.65 * MM), hPitchTwips: Math.round(101.6 * MM), vPitchTwips: Math.round(38.1 * MM) },
];

/** The texts the dialog writes from code (the products and pages), for a host's translation tables. */
export const LABELS_DIALOG_TEXTS: readonly string[] = [...LABEL_PRODUCTS.map((p) => p.label), "Custom label", "Letter (8 1/2 x 11 in)", "A4 (210 x 297 mm)"];

/** A label's cell margins: a little in from the label's left and right edges, as Word sets them. */
const CELL_SIDE = Math.round(0.06 * IN);

/** Why a sheet cannot be laid out, or null when it can. */
export function labelProblem(p: LabelProduct): "size" | null {
  if (!(p.across >= 1 && p.across <= 10 && p.down >= 1 && p.down <= 40)) return "size";
  if (!(p.widthTwips >= 0.5 * IN && p.heightTwips >= 0.25 * IN)) return "size";
  if (p.hPitchTwips < p.widthTwips || p.vPitchTwips < p.heightTwips || p.topTwips < 0 || p.sideTwips < 0) return "size";
  // (every label on the page)
  if (p.sideTwips + (p.across - 1) * p.hPitchTwips + p.widthTwips > p.pageWidthTwips + 2) return "size";
  if (p.topTwips + (p.down - 1) * p.vPitchTwips + p.heightTwips > p.pageHeightTwips + 2) return "size";
  return null;
}

/** The sheet's layout: the table's grid columns (a label, then the gap to the next, ...), its rows' height, and
 *  the page margins (the bottom a little, so the paragraph after the table stays on the page). */
export function labelLayout(p: LabelProduct) {
  const gap = p.hPitchTwips - p.widthTwips;
  const columns: number[] = [];
  for (let i = 0; i < p.across; i++) {
    if (i > 0 && gap > 0) columns.push(gap);
    columns.push(p.widthTwips);
  }
  const tableWidth = columns.reduce((s, w) => s + w, 0);
  const below = p.pageHeightTwips - p.topTwips - p.down * p.vPitchTwips;
  return {
    columns, labelColumns: columns.map((w, i) => gap <= 0 || i % 2 === 0),
    rowHeight: p.vPitchTwips,
    marginTop: p.topTwips, marginLeft: p.sideTwips,
    marginRight: Math.max(0, p.pageWidthTwips - p.sideTwips - tableWidth),
    marginBottom: Math.max(0, Math.min(360, below - 360)),
  };
}

const pt = (twips: number) => `${(twips / 20).toFixed(1)}pt`;
const NEXT_HTML = `<span data-field-code="NEXT">«Next Record»</span>`;

/** The sheet as a table: each label its lines (the ones after the first with Next Record first, when `next`), the
 *  gap columns empty. */
export function labelsHtml(p: LabelProduct, lines: Line[], next: boolean): string {
  const { columns, labelColumns } = labelLayout(p);
  const body = lines.length ? lines : [[]];
  let n = 0;
  const rows: string[] = [];
  for (let r = 0; r < p.down; r++) {
    const cells = columns.map((w, c) => {
      if (!labelColumns[c]) return `<td style="width:${pt(w)}"><p>&#8203;</p></td>`;
      const first = n++ === 0;
      const ps = body.map((line, i) => `<p>${i === 0 && next && !first ? NEXT_HTML : ""}${lineHtml(line)}</p>`).join("");
      return `<td style="width:${pt(w)}">${ps}</td>`;
    });
    rows.push(`<tr>${cells.join("")}</tr>`);
  }
  return `<html><body><table style="border-collapse:collapse;width:${pt(columns.reduce((s, w) => s + w, 0))}">${rows.join("")}</table></body></html>`;
}

export const LABELS_DIALOG_CSS = `
.dxr-lbldialog { width: min(660px, calc(100vw - 32px)); }
.dxr-lbldialog .dxr-lblbody { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 10px 16px; padding: 10px 16px 4px; }
.dxr-lbldialog label { display: flex; flex-direction: column; gap: 3px; }
.dxr-lbldialog label.dxr-dcheck { flex-direction: row; align-items: center; gap: 6px; }
.dxr-lbldialog select, .dxr-lbldialog input[type="text"], .dxr-lbldialog textarea { width: 100%; box-sizing: border-box; font: inherit; }
.dxr-lbldialog textarea { resize: vertical; min-height: 6em; font-family: Calibri, "Segoe UI", sans-serif; }
.dxr-lbldialog .dxr-lblwide { grid-column: 1 / -1; }
.dxr-lbldialog .dxr-lblcustom { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 6px 8px; }
.dxr-lbldialog .dxr-lblpreview { position: relative; border: 1px solid var(--dxr-rule-strong); background: #fff; margin: 4px auto 0;
  box-shadow: 0 1px 3px rgba(0,0,0,.15); overflow: hidden; font-family: Calibri, "Segoe UI", sans-serif; color: #111; }
.dxr-lbldialog .dxr-lblpreview div { position: absolute; box-sizing: border-box; border: 1px dashed #9aa4b2; border-radius: 3px;
  overflow: hidden; white-space: pre; line-height: 1.15; padding: 1px 2px; }
.dxr-lbldialog .dxr-lblnote { color: var(--dxr-muted); font-size: 11.5px; }
.dxr-lbldialog [hidden] { display: none !important; }
`;

export const LABELS_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-lbldialog" data-dxr="labelsdialog">
  <form method="dialog" novalidate>
    <h2 data-lb-title>Labels</h2>
    <div class="dxr-lblbody">
      <label class="dxr-lblwide">Label <select data-lb="product"></select></label>
      <div class="dxr-lblwide dxr-lblcustom" data-lb-row="custom">
        <label>Page <select data-lb="page"></select></label>
        <label>Label width (in) <input type="text" data-lb="width" value="2.625" /></label>
        <label>Label height (in) <input type="text" data-lb="height" value="1" /></label>
        <label>Across <input type="text" data-lb="across" value="3" /></label>
        <label>Down <input type="text" data-lb="down" value="10" /></label>
        <label>Top margin (in) <input type="text" data-lb="top" value="0.5" /></label>
        <label>Side margin (in) <input type="text" data-lb="side" value="0.1875" /></label>
        <label>Horizontal pitch (in) <input type="text" data-lb="hpitch" value="2.75" title="From one label's left edge to the next one's" /></label>
        <label>Vertical pitch (in) <input type="text" data-lb="vpitch" value="1" title="From one label's top edge to the next one's" /></label>
      </div>
      <div>
        <label data-lb-row="prefix">Columns of <select data-lb="prefix"></select></label>
        <label>Label text <textarea data-lb="text" spellcheck="false" title="«Name» is a merge field; other text stays text"></textarea></label>
        <label class="dxr-dcheck" title="Each label after the first starts with Next Record: a sheet takes as many recipients as it has labels">
          <input type="checkbox" data-lb="next" checked /> A different recipient on each label</label>
      </div>
      <div>
        <span>Preview</span>
        <div class="dxr-lblpreview" data-lb="preview"></div>
      </div>
      <span class="dxr-lblnote dxr-lblwide" data-lb="where"></span>
    </div>
    <div class="dxr-derror" role="alert"><span data-lb-msg="size" hidden>Give the custom label's sizes in inches: every label must fit on the page, the pitch at least the label's size.</span><span data-lb-msg="empty" hidden>Type the label's text, or choose the columns of an address.</span><span data-lb-msg="refused" hidden>The labels could not be made here.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok" data-lb="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
    <span hidden data-lb-text="allcolumns">(all columns)</span>
    <span hidden data-lb-text="new">This document becomes the sheet of labels.</span>
    <span hidden data-lb-text="add">The labels go in front of the document, as a section of its own.</span>
  </form>
</dialog>
`;

export interface LabelsDialogHost {
  columns(): string[];
  /** Whether the document has anything in it yet (then the sheet goes in front of it). */
  hasContent(): boolean;
  /** Make the sheet: the page and its margins, the table `html`, then the table laid out as `grid` says. */
  make(page: { widthTwips: number; heightTwips: number; marginLeft: number; marginTop: number; marginRight: number; marginBottom: number }, html: string,
    grid: { columnWidthsTwips: number[]; rowHeightTwips: number; cellSideTwips: number }): Promise<boolean>;
  translate(text: string): string;
  closed(): void;
}

export class LabelsDialog {
  constructor(private readonly dialog: HTMLDialogElement, private readonly host: LabelsDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-lb-title]")!);
    for (const el of Array.from(dialog.querySelectorAll<HTMLElement>("input, select, textarea")))
      el.addEventListener(el.tagName === "SELECT" || (el as HTMLInputElement).type === "checkbox" ? "change" : "input", () => this.update());
    this.select("prefix").addEventListener("change", () => { this.fillText(); this.update(); });
    // (a custom label starts from the product chosen before)
    this.select("product").addEventListener("change", () => { if (this.select("product").value === "custom") this.fillCustom(); this.update(); });
    const form = dialog.querySelector("form")!;
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter && submitter.value !== "ok") { this.close(); return; }
      void this.ok();
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.close(); });
  }

  private lastProduct = "5160";

  get open(): boolean {
    return this.dialog.open;
  }

  show(): void {
    const product = this.select("product");
    const keep = product.value;
    product.replaceChildren(...LABEL_PRODUCTS.map((p) => new Option(this.host.translate(p.label), p.key)), new Option(this.host.translate("Custom label"), "custom"));
    product.value = keep || "5160";
    const page = this.select("page");
    const keepPage = page.value;
    page.replaceChildren(new Option(this.host.translate("Letter (8 1/2 x 11 in)"), "letter"), new Option(this.host.translate("A4 (210 x 297 mm)"), "a4"));
    page.value = keepPage || "letter";
    const columns = this.host.columns();
    const prefixes = addressPrefixes(columns);
    const prefix = this.select("prefix");
    const all = this.dialog.querySelector("[data-lb-text=allcolumns]")?.textContent ?? "";
    const keptPrefix = prefix.value;
    prefix.replaceChildren(...(prefixes.length ? prefixes : [""]).map((p) => new Option(p ? `${p}…` : all, p)));
    if (Array.from(prefix.options).some((o) => o.value === keptPrefix)) prefix.value = keptPrefix;
    const row = this.dialog.querySelector<HTMLElement>('[data-lb-row="prefix"]');
    if (row) row.hidden = !(prefixes.length > 1 || (prefixes.length === 1 && prefixes[0] !== ""));
    this.fillText();
    // (no recipients: a page of the same label)
    this.input("next").checked = columns.length > 0;
    this.part("where").textContent = this.dialog.querySelector(`[data-lb-text=${this.host.hasContent() ? "add" : "new"}]`)?.textContent ?? "";
    this.showError(null);
    this.update();
    if (!this.dialog.open) this.dialog.showModal();
    this.textarea("text").focus();
  }

  close(): void {
    if (!this.dialog.open) return;
    this.dialog.close();
    this.host.closed();
  }

  /** The sheet chosen, or null for a custom one that does not fit its page. */
  product(): LabelProduct | null {
    const key = this.select("product").value;
    const known = LABEL_PRODUCTS.find((p) => p.key === key);
    if (known) return known;
    const v = (k: string) => Number(this.input(k).value.replace(",", "."));
    const whole = (k: string) => (Number.isInteger(v(k)) ? v(k) : NaN);
    const page = this.select("page").value === "a4" ? A4 : LETTER;
    const p: LabelProduct = {
      key: "custom", label: "", ...page,
      widthTwips: Math.round(v("width") * IN), heightTwips: Math.round(v("height") * IN),
      across: whole("across"), down: whole("down"),
      topTwips: Math.round(v("top") * IN), sideTwips: Math.round(v("side") * IN),
      hPitchTwips: Math.round(v("hpitch") * IN), vPitchTwips: Math.round(v("vpitch") * IN),
    };
    return labelProblem(p) ? null : p;
  }

  private fillCustom(): void {
    const from = LABEL_PRODUCTS.find((p) => p.key === this.lastProduct) ?? LABEL_PRODUCTS[0];
    const inches = (t: number) => String(Math.round((t / IN) * 10000) / 10000);
    this.select("page").value = from.pageWidthTwips === A4.pageWidthTwips ? "a4" : "letter";
    this.input("width").value = inches(from.widthTwips);
    this.input("height").value = inches(from.heightTwips);
    this.input("across").value = String(from.across);
    this.input("down").value = String(from.down);
    this.input("top").value = inches(from.topTwips);
    this.input("side").value = inches(from.sideTwips);
    this.input("hpitch").value = inches(from.hPitchTwips);
    this.input("vpitch").value = inches(from.vPitchTwips);
  }

  private fillText(): void {
    const columns = this.host.columns();
    if (!columns.length) { this.textarea("text").value = ""; return; }
    const match = matchAddress(columns, this.select("prefix").value);
    this.textarea("text").value = textOf(addressLines(match, { name: match.full ? "full" : "titlefirstlast", company: true, address: true, country: false }));
  }

  private update(): void {
    const key = this.select("product").value;
    if (key !== "custom") this.lastProduct = key;
    const custom = this.dialog.querySelector<HTMLElement>('[data-lb-row="custom"]');
    if (custom) custom.hidden = key !== "custom";
    const p = this.product();
    const preview = this.part("preview");
    if (p) {
      const scale = 220 / p.pageWidthTwips;
      preview.style.width = `${Math.round(p.pageWidthTwips * scale)}px`;
      preview.style.height = `${Math.round(p.pageHeightTwips * scale)}px`;
      preview.style.fontSize = `${Math.max(3, 200 * scale)}px`;
      const doc = this.dialog.ownerDocument;
      const text = this.textarea("text").value;
      const next = this.input("next").checked;
      const labels: HTMLElement[] = [];
      for (let r = 0; r < p.down; r++)
        for (let c = 0; c < p.across; c++) {
          const box = doc.createElement("div");
          box.style.left = `${Math.round((p.sideTwips + c * p.hPitchTwips) * scale)}px`;
          box.style.top = `${Math.round((p.topTwips + r * p.vPitchTwips) * scale)}px`;
          box.style.width = `${Math.max(2, Math.round(p.widthTwips * scale))}px`;
          box.style.height = `${Math.max(2, Math.round(p.heightTwips * scale))}px`;
          box.textContent = r === 0 && c === 0 ? text : next ? "" : text;
          labels.push(box);
        }
      preview.replaceChildren(...labels);
    }
    this.button("ok").disabled = !p || !this.textarea("text").value.trim();
  }

  private async ok(): Promise<void> {
    const p = this.product();
    if (!p) { this.showError("size"); return; }
    const lines = linesOf(this.textarea("text").value);
    if (!lines.length) { this.showError("empty"); return; }
    const layout = labelLayout(p);
    const html = labelsHtml(p, lines, this.input("next").checked);
    this.dialog.close();
    if (!(await this.host.make({ widthTwips: p.pageWidthTwips, heightTwips: p.pageHeightTwips, marginLeft: layout.marginLeft, marginTop: layout.marginTop,
      marginRight: layout.marginRight, marginBottom: layout.marginBottom }, html,
      { columnWidthsTwips: layout.columns, rowHeightTwips: layout.rowHeight, cellSideTwips: CELL_SIDE }))) {
      this.dialog.showModal();
      this.showError("refused");
      return;
    }
    this.host.closed();
  }

  private showError(which: "size" | "empty" | "refused" | null): void {
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-lb-msg]"))) el.hidden = el.dataset.lbMsg !== which;
  }

  private part(key: string): HTMLElement {
    return this.dialog.querySelector<HTMLElement>(`[data-lb="${key}"]`)!;
  }

  private button(key: string): HTMLButtonElement {
    return this.dialog.querySelector<HTMLButtonElement>(`button[data-lb="${key}"]`)!;
  }

  private input(key: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`input[data-lb="${key}"]`)!;
  }

  private textarea(key: string): HTMLTextAreaElement {
    return this.dialog.querySelector<HTMLTextAreaElement>(`textarea[data-lb="${key}"]`)!;
  }

  private select(key: string): HTMLSelectElement {
    return this.dialog.querySelector<HTMLSelectElement>(`select[data-lb="${key}"]`)!;
  }
}
