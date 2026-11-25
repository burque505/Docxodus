/**
 * The ribbon's Caption dialogs — Word's References > Insert Caption (the caption's text after its
 * label and number, the label, above or below the item, "Exclude label from caption", the number's
 * format, New Label / Delete Label) and Insert Table of Figures (which label it lists, with or
 * without the label and number, as links). Both are modal; OK is one engine call.
 */
import type { CaptionOptions, TableOfFiguresOptions } from "./types.js";
import { makeDialogMovable, restoreDialogPlace } from "./ribbon-paragraph-dialog.js";

export const CAPTION_DIALOG_CSS = `
.dxr-cpdialog { width: min(420px, calc(100vw - 32px)); }
.dxr-cpdialog .dxr-cpbody { display: grid; grid-template-columns: auto 1fr; gap: 8px 10px; align-items: center; padding: 10px 16px 4px; }
.dxr-cpdialog .dxr-cpwide { grid-column: 1 / -1; }
.dxr-cpdialog .dxr-cpcaption { display: flex; align-items: baseline; gap: 0; border: 1px solid var(--dxr-rule-strong, #cbd5e1);
  border-radius: 4px; padding: 0 6px; background: var(--dxr-paper, #fff); min-width: 0; }
.dxr-cpdialog .dxr-cpcaption span { white-space: pre; color: var(--dxr-ink-soft, #667085); }
.dxr-cpdialog .dxr-cpcaption input { flex: 1 1 auto; min-width: 0; border: 0; outline: none; padding: 4px 0; background: transparent; font: inherit; }
.dxr-cpdialog select { width: 100%; box-sizing: border-box; }
.dxr-cpdialog .dxr-cprow { display: flex; gap: 6px; flex-wrap: wrap; }
.dxr-cpdialog .dxr-cpnew { display: flex; gap: 6px; }
.dxr-cpdialog .dxr-cpnew[hidden] { display: none; }
.dxr-cpdialog .dxr-cpnew input { flex: 1 1 auto; min-width: 0; }
`;

const FORMATS: ReadonlyArray<[NonNullable<CaptionOptions["numberFormat"]>, string]> = [
  ["ARABIC", "1, 2, 3, …"], ["alphabetic", "a, b, c, …"], ["ALPHABETIC", "A, B, C, …"],
  ["roman", "i, ii, iii, …"], ["ROMAN", "I, II, III, …"],
];

export const CAPTION_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-cpdialog" data-dxr="captiondialog">
  <form method="dialog" novalidate>
    <h2 data-cp-title>Caption</h2>
    <div class="dxr-cpbody">
      <label class="dxr-cpwide" for="dxr-cp-text">Caption:</label>
      <div class="dxr-cpcaption dxr-cpwide"><span data-cp="prefix"></span><input id="dxr-cp-text" type="text" autocomplete="off" data-cp="text" /></div>
      <label>Label:</label><select data-cp="label"></select>
      <label>Position:</label><select data-cp="position">
        <option value="above">Above selected item</option>
        <option value="below">Below selected item</option>
      </select>
      <label>Numbering:</label><select data-cp="format">
        ${FORMATS.map(([value, label]) => `<option value="${value}">${label}</option>`).join("")}
      </select>
      <label class="dxr-dcheck dxr-cpwide"><input type="checkbox" data-cp="exclude" /> Exclude label from caption</label>
      <div class="dxr-cprow dxr-cpwide">
        <button type="button" data-cp="newlabel">New Label...</button>
        <button type="button" data-cp="deletelabel">Delete Label</button>
      </div>
      <div class="dxr-cpnew dxr-cpwide" data-cp="newrow" hidden>
        <input type="text" autocomplete="off" maxlength="40" aria-label="New label" data-cp="newname" />
        <button type="button" data-cp="addlabel">Add</button>
      </div>
    </div>
    <div class="dxr-derror" role="alert"><span data-cp-msg="label" hidden>A label has 1 to 40 characters, without quotation marks or backslashes.</span><span data-cp-msg="refused" hidden>The caption cannot be put there (are changes being tracked?).</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
<dialog class="dxr-dialog dxr-cpdialog" data-dxr="tofdialog">
  <form method="dialog" novalidate>
    <h2 data-tof-title>Table of Figures</h2>
    <div class="dxr-cpbody">
      <label>Caption label:</label><select data-tof="label"></select>
      <label class="dxr-dcheck dxr-cpwide"><input type="checkbox" data-tof="include" checked /> Include label and number</label>
      <label class="dxr-dcheck dxr-cpwide"><input type="checkbox" data-tof="links" checked /> Link each entry to its caption</label>
    </div>
    <div class="dxr-derror" role="alert"><span data-tof-msg="refused" hidden>The table cannot be put there.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

/** A caption label Word takes: 1 to 40 characters, no quotation mark or backslash (a SEQ field's
 *  identifier, spaces written as underscores). */
export function isCaptionLabel(label: string): boolean {
  const t = label.trim();
  return t.length > 0 && t.length <= 40 && !/["\\]/.test(t);
}

/** A caption's number in a SEQ format (as the engine writes it). */
export function formatCaptionNumber(n: number, format: NonNullable<CaptionOptions["numberFormat"]>): string {
  const lower = format[0] === format[0].toLowerCase();
  let text: string;
  if (/^alphabetic$/i.test(format)) text = String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
  else if (/^roman$/i.test(format)) {
    text = "";
    let rest = n;
    for (const [v, s] of [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
      [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]] as const)
      while (rest >= v) { text += s; rest -= v; }
  } else return String(n);
  return lower ? text.toLowerCase() : text;
}

export interface CaptionDialogHost {
  /** The labels to offer, built-in ones first, as shown (spaces, not underscores). */
  labels(): string[];
  /** Whether `label` is one the user made (New Label), so Delete Label may take it away. */
  isCustomLabel(label: string): boolean;
  addLabel(label: string): void;
  deleteLabel(label: string): void;
  /** The number the new caption would show for `label` at the caret. */
  nextNumber(label: string, position: "above" | "below"): number;
  /** What the caption is for: a table (above by default) or a paragraph (below); null when none. */
  target(): "table" | "paragraph" | null;
  insert(options: CaptionOptions): boolean;
  /** The dialog closed (OK or Cancel): the focus goes back to the document. */
  closed?(): void;
}

/** Word's Insert Caption dialog over its markup (`CAPTION_DIALOG_HTML`, already in the page). */
export class CaptionDialog {
  private lastLabel = "Figure";
  private lastFormat: NonNullable<CaptionOptions["numberFormat"]> = "ARABIC";

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: CaptionDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-cp-title]")!);
    const form = dialog.querySelector("form")!;
    for (const name of ["label", "format", "exclude", "position"])
      this.field(name).addEventListener("change", () => { this.showError(null); this.preview(); });
    this.field("newlabel").addEventListener("click", () => {
      const row = this.field("newrow");
      row.hidden = !row.hidden;
      if (!row.hidden) { this.input("newname").value = ""; this.input("newname").focus(); }
    });
    this.field("addlabel").addEventListener("click", () => this.addLabel());
    this.input("newname").addEventListener("keydown", (event) => {
      if (event.key === "Enter") { event.preventDefault(); this.addLabel(); }
    });
    this.field("deletelabel").addEventListener("click", () => {
      const label = this.select("label").value;
      if (!this.host.isCustomLabel(label)) return;
      this.host.deleteLabel(label);
      this.fillLabels("Figure");
    });
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.dialog.close("cancel"); return; }
      const label = this.select("label").value;
      const format = this.select("format").value as NonNullable<CaptionOptions["numberFormat"]>;
      const options: CaptionOptions = {
        label,
        position: this.select("position").value === "above" ? "above" : "below",
        numberFormat: format,
      };
      if (this.input("exclude").checked) options.excludeLabel = true;
      const text = this.input("text").value;
      if (text) options.text = text;
      if (!this.host.insert(options)) { this.showError("refused"); return; }
      this.lastLabel = label;
      this.lastFormat = format;
      this.dialog.close("ok");
      this.host.closed?.();
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.dialog.close("cancel"); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  /** Open on the caret's item; false (nothing opened) when there is nothing to caption. */
  show(): boolean {
    const target = this.host.target();
    if (!target) return false;
    this.fillLabels(this.lastLabel);
    this.select("format").value = this.lastFormat;
    this.select("position").value = target === "table" ? "above" : "below";
    this.input("exclude").checked = false;
    this.input("text").value = "";
    this.field("newrow").hidden = true;
    this.showError(null);
    this.preview();
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.input("text").focus();
    return true;
  }

  private addLabel(): void {
    const name = this.input("newname").value.trim();
    if (!isCaptionLabel(name)) { this.input("newname").setAttribute("aria-invalid", "true"); this.showError("label"); return; }
    this.input("newname").removeAttribute("aria-invalid");
    this.host.addLabel(name);
    this.field("newrow").hidden = true;
    this.fillLabels(name);
    this.preview();
    this.input("text").focus();
  }

  private fillLabels(selected: string): void {
    const select = this.select("label");
    select.replaceChildren();
    const labels = this.host.labels();
    for (const label of labels) {
      const option = select.ownerDocument.createElement("option");
      option.value = option.textContent = label;
      select.appendChild(option);
    }
    select.value = labels.includes(selected) ? selected : labels[0] ?? "";
    this.enable();
  }

  /** The label and number before the caption's text, as the caption will show them. */
  private preview(): void {
    const label = this.select("label").value;
    const format = this.select("format").value as NonNullable<CaptionOptions["numberFormat"]>;
    const number = formatCaptionNumber(this.host.nextNumber(label, this.select("position").value === "above" ? "above" : "below"), format);
    this.field("prefix").textContent = this.input("exclude").checked ? number : `${label} ${number}`;
    this.enable();
  }

  private enable(): void {
    (this.field("deletelabel") as HTMLButtonElement).disabled = !this.host.isCustomLabel(this.select("label").value);
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-cp-msg]")))
      span.hidden = span.dataset.cpMsg !== message;
  }

  private field(name: string): HTMLElement {
    return this.dialog.querySelector<HTMLElement>(`[data-cp="${name}"]`)!;
  }

  private input(name: string): HTMLInputElement {
    return this.field(name) as HTMLInputElement;
  }

  private select(name: string): HTMLSelectElement {
    return this.field(name) as HTMLSelectElement;
  }
}

export interface TableOfFiguresDialogHost {
  labels(): string[];
  insert(options: TableOfFiguresOptions): boolean;
  closed?(): void;
}

/** Word's Table of Figures dialog over its markup (`CAPTION_DIALOG_HTML`'s second dialog). */
export class TableOfFiguresDialog {
  private lastLabel = "Figure";

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: TableOfFiguresDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-tof-title]")!);
    const form = dialog.querySelector("form")!;
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.dialog.close("cancel"); return; }
      const label = this.field<HTMLSelectElement>("label").value;
      const options: TableOfFiguresOptions = {
        // (the SEQ identifier: Word writes a label's spaces as underscores)
        captionLabel: label.replace(/ /g, "_"),
        hyperlinks: this.field<HTMLInputElement>("links").checked,
      };
      if (!this.field<HTMLInputElement>("include").checked) options.includeLabelAndNumber = false;
      if (!this.host.insert(options)) { this.showError("refused"); return; }
      this.lastLabel = label;
      this.dialog.close("ok");
      this.host.closed?.();
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.dialog.close("cancel"); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  show(): void {
    const select = this.field<HTMLSelectElement>("label");
    select.replaceChildren();
    const labels = this.host.labels();
    for (const label of labels) {
      const option = select.ownerDocument.createElement("option");
      option.value = option.textContent = label;
      select.appendChild(option);
    }
    select.value = labels.includes(this.lastLabel) ? this.lastLabel : labels[0] ?? "";
    this.showError(null);
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    select.focus();
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-tof-msg]")))
      span.hidden = span.dataset.tofMsg !== message;
  }

  private field<T extends HTMLElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-tof="${name}"]`)!;
  }
}
