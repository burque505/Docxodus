/**
 * The ribbon's Table Properties dialog — Word's, with its Table, Row, Column, Cell and Alt Text
 * tabs, for the table, row, column and cell the caret is in. OK writes only what was changed, as
 * one edit (one undo step).
 */
import type { CellPropertiesOp, TablePropertiesInfo, TablePropertiesOp, TableRowOptions } from "./types.js";
import { formatLength, makeDialogMovable, parseLength, restoreDialogPlace, type MeasurementUnit } from "./ribbon-paragraph-dialog.js";

export const TABLE_DIALOG_CSS = `
.dxr-tbldialog { width: min(460px, calc(100vw - 32px)); }
.dxr-tbldialog h3 { margin: 10px 0 4px; font-size: 12px; font-weight: 600; color: var(--dxr-ink); }
.dxr-tbldialog .dxr-tdrow { display: flex; flex-wrap: wrap; gap: 6px 10px; align-items: center; margin: 4px 0; }
.dxr-tbldialog .dxr-tdrow label { white-space: nowrap; }
.dxr-tbldialog input[type="text"] { width: 6em; }
.dxr-tbldialog .dxr-tdmar { display: grid; grid-template-columns: auto auto auto auto; gap: 6px 8px; align-items: center; margin: 4px 0; }
.dxr-tbldialog .dxr-tdwhich { color: var(--dxr-muted); margin: 6px 0 0; }
.dxr-tbldialog textarea { width: 100%; box-sizing: border-box; min-height: 5.5em; resize: vertical; font: inherit; }
.dxr-tbldialog input[data-td="alttitle"] { width: 100%; box-sizing: border-box; }
`;

export const TABLE_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-tbldialog" data-dxr="tabledialog">
  <form method="dialog" novalidate>
    <h2 data-td-title>Table Properties</h2>
    <div role="tablist">
      <button type="button" role="tab" data-td-tab="table" aria-selected="true">Table</button>
      <button type="button" role="tab" data-td-tab="row" aria-selected="false">Row</button>
      <button type="button" role="tab" data-td-tab="column" aria-selected="false">Column</button>
      <button type="button" role="tab" data-td-tab="cell" aria-selected="false">Cell</button>
      <button type="button" role="tab" data-td-tab="alt" aria-selected="false">Alt Text</button>
    </div>
    <div role="tabpanel" data-td-panel="table">
      <h3>Size</h3>
      <div class="dxr-tdrow">
        <label><input type="checkbox" data-td="haswidth" /> Preferred width</label>
        <input type="text" autocomplete="off" data-td="width" />
        <label>Measure in</label>
        <select data-td="widthtype"><option value="in">Inches</option><option value="cm">Centimeters</option><option value="percent">Percent</option></select>
      </div>
      <h3>Alignment</h3>
      <div class="dxr-tdrow">
        <label><input type="radio" name="tdalign" value="left" data-td="align-left" /> Left</label>
        <label><input type="radio" name="tdalign" value="center" data-td="align-center" /> Center</label>
        <label><input type="radio" name="tdalign" value="right" data-td="align-right" /> Right</label>
        <label>Indent from left</label><input type="text" autocomplete="off" data-td="indent" />
      </div>
      <h3>Default cell margins</h3>
      <div class="dxr-tdmar">
        <label>Top</label><input type="text" autocomplete="off" data-td="mtop" />
        <label>Left</label><input type="text" autocomplete="off" data-td="mleft" />
        <label>Bottom</label><input type="text" autocomplete="off" data-td="mbottom" />
        <label>Right</label><input type="text" autocomplete="off" data-td="mright" />
      </div>
    </div>
    <div role="tabpanel" data-td-panel="row" hidden>
      <p class="dxr-tdwhich" data-td-which="row"><span>Row</span> <span data-td-n></span></p>
      <h3>Size</h3>
      <div class="dxr-tdrow">
        <label><input type="checkbox" data-td="hasheight" /> Specify height</label>
        <input type="text" autocomplete="off" data-td="height" />
        <label>Row height is</label>
        <select data-td="heightrule"><option value="atLeast">At least</option><option value="exact">Exactly</option></select>
      </div>
      <h3>Options</h3>
      <label class="dxr-dcheck"><input type="checkbox" data-td="allowbreak" /> Allow row to break across pages</label>
      <label class="dxr-dcheck"><input type="checkbox" data-td="repeatheader" /> Repeat as header row at the top of each page</label>
    </div>
    <div role="tabpanel" data-td-panel="column" hidden>
      <p class="dxr-tdwhich" data-td-which="column"><span data-td-one>Column</span><span data-td-many hidden>Columns</span> <span data-td-n></span></p>
      <h3>Size</h3>
      <div class="dxr-tdrow">
        <label>Preferred width</label><input type="text" autocomplete="off" data-td="colwidth" />
      </div>
    </div>
    <div role="tabpanel" data-td-panel="cell" hidden>
      <h3>Vertical alignment</h3>
      <div class="dxr-tdrow">
        <label><input type="radio" name="tdvalign" value="top" data-td="valign-top" /> Top</label>
        <label><input type="radio" name="tdvalign" value="center" data-td="valign-center" /> Center</label>
        <label><input type="radio" name="tdvalign" value="bottom" data-td="valign-bottom" /> Bottom</label>
      </div>
    </div>
    <div role="tabpanel" data-td-panel="alt" hidden>
      <h3>Title</h3>
      <input type="text" autocomplete="off" data-td="alttitle" />
      <h3>Description</h3>
      <textarea data-td="altdesc" placeholder="How would you describe this table to someone who cannot see it?"></textarea>
    </div>
    <div class="dxr-derror" role="alert"><span data-td-msg="length" hidden>Type a measurement, such as 2" or 5 cm.</span><span data-td-msg="size" hidden>Type a size greater than zero.</span><span data-td-msg="percent" hidden>Type a percentage from 1 to 100.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

/** What OK asks for: only the parts that were changed. */
export interface TableDialogResult {
  table?: TablePropertiesOp;
  row?: TableRowOptions;
  columnWidthTwips?: number;
  cell?: CellPropertiesOp;
}

export interface TableDialogHost {
  /** The caret's table, or null when the caret is not in one. */
  read(): TablePropertiesInfo | null;
  apply(result: TableDialogResult): void;
}

export type TableDialogTab = "table" | "row" | "column" | "cell" | "alt";

/** Word's default cell margins: 0.08" left and right, none top and bottom. */
const DEFAULT_SIDE_MARGIN = 108;

/** The Table Properties dialog over its markup (`TABLE_DIALOG_HTML`, already in the page). */
export class TableDialog {
  private unit: MeasurementUnit;
  private info: TablePropertiesInfo | null = null;
  private initial: Record<string, unknown> = {};

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: TableDialogHost, unit: MeasurementUnit = "in") {
    this.unit = unit;
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-td-title]")!);
    const form = dialog.querySelector("form")!;
    for (const tab of Array.from(dialog.querySelectorAll<HTMLElement>("[data-td-tab]")))
      tab.addEventListener("click", () => this.selectTab(tab.dataset.tdTab as TableDialogTab));
    form.addEventListener("input", (event) => {
      const target = event.target as HTMLElement;
      target.removeAttribute("aria-invalid");
      this.showError(null);
      // Typing a width or a height turns its box on.
      if (target.dataset.td === "width") this.input("haswidth").checked = true;
      if (target.dataset.td === "height") this.input("hasheight").checked = true;
      this.enable();
    });
    form.addEventListener("change", (event) => {
      const name = (event.target as HTMLElement).dataset.td;
      if (name === "widthtype") this.convertWidth();
      this.enable();
    });
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.dialog.close("cancel"); return; }
      const result = this.collect();
      if (!result) return;
      this.dialog.close("ok");
      if (Object.keys(result).length) this.host.apply(result);
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.dialog.close("cancel"); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  setUnit(unit: MeasurementUnit): void {
    this.unit = unit;
  }

  /** Open on the caret's table at `tab`. False (nothing opened) outside a table. */
  show(tab: TableDialogTab = "table"): boolean {
    const info = this.host.read();
    if (!info) return false;
    this.info = info;
    const len = (twips: number | null | undefined, fallback = 0): string => formatLength(twips ?? fallback, this.unit);
    // Measure in: the unit in use, or Percent.
    for (const option of Array.from(this.select("widthtype").options))
      option.hidden = option.value !== "percent" && option.value !== this.unit;

    const hasWidth = info.widthType === "twips" || info.widthType === "percent";
    this.input("haswidth").checked = hasWidth;
    this.select("widthtype").value = info.widthType === "percent" ? "percent" : this.unit;
    const gridWidth = info.columnWidthsTwips.reduce((sum, w) => sum + w, 0);
    this.input("width").value = info.widthType === "percent"
      ? `${trim(info.width ?? 100)}%`
      : len(hasWidth ? info.width : gridWidth);
    const align = info.alignment === "center" || info.alignment === "right" ? info.alignment : "left";
    this.input(`align-${align}`).checked = true;
    this.input("indent").value = len(info.indentTwips);
    const m = info.cellMargins;
    this.input("mtop").value = len(m?.topTwips);
    this.input("mbottom").value = len(m?.bottomTwips);
    this.input("mleft").value = len(m?.leftTwips, DEFAULT_SIDE_MARGIN);
    this.input("mright").value = len(m?.rightTwips, DEFAULT_SIDE_MARGIN);

    this.which("row", `${info.rowIndex + 1}:`);
    this.input("hasheight").checked = info.rowHeightTwips !== null && info.rowHeightTwips > 0;
    this.input("height").value = len(info.rowHeightTwips);
    this.select("heightrule").value = info.rowHeightRule === "exact" ? "exact" : "atLeast";
    this.input("allowbreak").checked = info.allowBreakAcrossPages;
    this.input("repeatheader").checked = info.repeatHeader;

    const span = info.columnWidthsTwips.slice(info.columnIndex, info.columnIndex + info.columnSpan);
    this.which("column", info.columnSpan > 1
      ? `${info.columnIndex + 1}–${info.columnIndex + info.columnSpan}:`
      : `${info.columnIndex + 1}:`, info.columnSpan > 1);
    this.input("colwidth").value = len(span.reduce((sum, w) => sum + w, 0));

    const valign = info.verticalAlignment === "center" || info.verticalAlignment === "bottom" ? info.verticalAlignment : "top";
    this.input(`valign-${valign}`).checked = true;

    this.input("alttitle").value = info.altTitle ?? "";
    this.field<HTMLTextAreaElement>("altdesc").value = info.altDescription ?? "";

    this.enable();
    this.initial = this.snapshot();
    this.showError(null);
    this.selectTab(tab);
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    const first = { table: "width", row: "height", column: "colwidth", cell: "valign-top", alt: "alttitle" }[tab];
    const target = this.input(first);
    (target.disabled ? this.dialog.querySelector<HTMLElement>(`[data-td-tab="${tab}"]`)! : target).focus();
    return true;
  }

  selectTab(tab: TableDialogTab): void {
    for (const t of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-td-tab]")))
      t.setAttribute("aria-selected", String(t.dataset.tdTab === tab));
    for (const panel of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-td-panel]")))
      panel.hidden = panel.dataset.tdPanel !== tab;
  }

  /** The changes the fields ask for, or null (the bad field marked) when one is not valid. */
  collect(): TableDialogResult | null {
    const info = this.info;
    if (!info) return null;
    const now = this.snapshot();
    const was = this.initial;
    const changed = (...keys: string[]): boolean => keys.some((k) => JSON.stringify(now[k]) !== JSON.stringify(was[k]));
    const result: TableDialogResult = {};

    const table: TablePropertiesOp = {};
    if (changed("haswidth", "width", "widthtype")) {
      if (!this.input("haswidth").checked) table.widthType = "auto";
      else if (this.select("widthtype").value === "percent") {
        const pct = Number(this.input("width").value.replace("%", "").replace(",", ".").trim());
        if (!Number.isFinite(pct) || pct <= 0 || pct > 100) return this.invalid(this.input("width"), "percent");
        table.widthType = "percent";
        table.width = pct;
      } else {
        const twips = this.length("width");
        if (twips === undefined) return null;
        if (twips <= 0) return this.invalid(this.input("width"), "size");
        table.widthType = "twips";
        table.width = twips;
      }
    }
    if (changed("align")) table.alignment = now.align as "left" | "center" | "right";
    if (changed("indent")) {
      const twips = this.length("indent");
      if (twips === undefined) return null;
      table.indentTwips = twips;
    }
    const edges = [["mtop", "topTwips"], ["mleft", "leftTwips"], ["mbottom", "bottomTwips"], ["mright", "rightTwips"]] as const;
    if (changed(...edges.map(([k]) => k))) {
      const margins: Record<string, number> = {};
      for (const [field, key] of edges) {
        const twips = this.length(field, true);
        if (twips === undefined) return null;
        margins[key] = twips;
      }
      table.cellMargins = margins;
    }
    if (changed("alttitle")) table.altTitle = now.alttitle as string;
    if (changed("altdesc")) table.altDescription = now.altdesc as string;
    if (Object.keys(table).length) result.table = table;

    const row: TableRowOptions = {};
    if (changed("hasheight", "height", "heightrule")) {
      if (!this.input("hasheight").checked) row.heightTwips = 0;
      else {
        const twips = this.length("height");
        if (twips === undefined) return null;
        if (twips <= 0) return this.invalid(this.input("height"), "size");
        row.heightTwips = twips;
      }
      row.heightRule = this.select("heightrule").value as "atLeast" | "exact";
    }
    if (changed("allowbreak")) row.allowBreakAcrossPages = this.input("allowbreak").checked;
    if (changed("repeatheader")) row.repeatHeader = this.input("repeatheader").checked;
    if (Object.keys(row).length) result.row = row;

    if (changed("colwidth")) {
      const twips = this.length("colwidth");
      if (twips === undefined) return null;
      if (twips <= 0) return this.invalid(this.input("colwidth"), "size");
      result.columnWidthTwips = twips;
    }
    if (changed("valign")) result.cell = { verticalAlignment: now.valign as "top" | "center" | "bottom" };
    return result;
  }

  /** The fields as they stand, to compare with how they opened. */
  private snapshot(): Record<string, unknown> {
    const v = (name: string): string => this.field<HTMLInputElement>(name).value;
    const c = (name: string): boolean => this.input(name).checked;
    const radio = (group: string): string =>
      this.dialog.querySelector<HTMLInputElement>(`input[name="${group}"]:checked`)?.value ?? "";
    return {
      haswidth: c("haswidth"), width: v("width"), widthtype: v("widthtype"), align: radio("tdalign"), indent: v("indent"),
      mtop: v("mtop"), mleft: v("mleft"), mbottom: v("mbottom"), mright: v("mright"),
      hasheight: c("hasheight"), height: v("height"), heightrule: v("heightrule"),
      allowbreak: c("allowbreak"), repeatheader: c("repeatheader"),
      colwidth: v("colwidth"), valign: radio("tdvalign"),
      alttitle: v("alttitle"), altdesc: v("altdesc"),
    };
  }

  /** Switching Measure in between a length and a percent shows the same width the other way. */
  private convertWidth(): void {
    const info = this.info;
    if (!info) return;
    const field = this.input("width");
    const grid = info.columnWidthsTwips.reduce((sum, w) => sum + w, 0) || 1;
    if (this.select("widthtype").value === "percent") {
      const twips = parseLength(field.value, this.unit);
      // Of the text width the grid fills now: a table as wide as its columns reads as 100%.
      const pct = twips === null ? 100 : Math.min(100, Math.max(1, Math.round((twips / grid) * 100)));
      field.value = `${pct}%`;
    } else {
      const pct = Number(field.value.replace("%", "").trim());
      field.value = formatLength(Number.isFinite(pct) ? Math.round((grid * pct) / 100) : grid, this.unit);
    }
  }

  /** The width and height boxes apply only while their checkbox is on; the indent only for a
   *  left-aligned table, as in Word. */
  private enable(): void {
    this.input("width").disabled = !this.input("haswidth").checked;
    this.select("widthtype").disabled = !this.input("haswidth").checked;
    this.input("height").disabled = !this.input("hasheight").checked;
    this.select("heightrule").disabled = !this.input("hasheight").checked;
    this.input("indent").disabled = !this.input("align-left").checked;
  }

  /** "Row 2:", "Column 2:" or "Columns 2–3:": the words are the markup's (so they can be
   *  translated), the numbers are filled in. */
  private which(panel: string, numbers: string, many = false): void {
    const line = this.dialog.querySelector<HTMLElement>(`[data-td-which="${panel}"]`)!;
    line.querySelector<HTMLElement>("[data-td-n]")!.textContent = numbers;
    const one = line.querySelector<HTMLElement>("[data-td-one]");
    const several = line.querySelector<HTMLElement>("[data-td-many]");
    if (one && several) { one.hidden = many; several.hidden = !many; }
  }

  /** A length field in twips; undefined (the field marked) when it does not read as one. */
  private length(name: string, nonNegative = false): number | undefined {
    const field = this.input(name);
    const twips = parseLength(field.value, this.unit);
    if (twips === null || (nonNegative && twips < 0)) {
      this.invalid(field, "length");
      return undefined;
    }
    return twips;
  }

  private invalid(field: HTMLElement, message: string): null {
    const panel = field.closest<HTMLElement>("[data-td-panel]");
    if (panel?.hidden) this.selectTab(panel.dataset.tdPanel as TableDialogTab);
    field.setAttribute("aria-invalid", "true");
    this.showError(message);
    field.focus();
    return null;
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-td-msg]")))
      span.hidden = span.dataset.tdMsg !== message;
  }

  private field<T extends HTMLElement = HTMLElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-td="${name}"]`)!;
  }

  private input(name: string): HTMLInputElement {
    return this.field<HTMLInputElement>(name);
  }

  private select(name: string): HTMLSelectElement {
    return this.field<HTMLSelectElement>(name);
  }
}

function trim(n: number): string {
  return String(Math.round(n * 100) / 100);
}
