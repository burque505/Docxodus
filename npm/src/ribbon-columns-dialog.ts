/**
 * The ribbon's Columns dialog — Word's Layout > Columns > More Columns: how many equal columns,
 * the space between them, a line between, and Word's "Apply to" (this section, this point
 * forward, the whole document, or the selected text, which gets a section of its own). OK is
 * one `setColumns`.
 */
import type { ColumnsApplyTo, ColumnsOp, SectionInfo } from "./types.js";
import { formatLength, makeDialogMovable, parseLength, restoreDialogPlace, type MeasurementUnit } from "./ribbon-paragraph-dialog.js";

export const COLUMNS_DIALOG_CSS = `
.dxr-coldialog { width: min(380px, calc(100vw - 32px)); }
.dxr-coldialog .dxr-cdbody { padding: 10px 16px 4px; }
.dxr-coldialog .dxr-dgrid input[type="number"] { width: 5em; }
.dxr-coldialog .dxr-cdpreview {
  display: flex; gap: 0; height: 64px; margin-top: 10px; padding: 8px 10px; border: 1px solid var(--dxr-rule);
  border-radius: 6px; background: var(--dxr-sheet);
}
.dxr-coldialog .dxr-cdpreview span {
  flex: 1; background: repeating-linear-gradient(var(--dxr-rule-strong) 0 3px, transparent 3px 8px);
}
.dxr-coldialog .dxr-cdpreview i { flex: 0 0 auto; display: block; }
.dxr-coldialog .dxr-cdpreview i[data-line] { background: linear-gradient(var(--dxr-muted), var(--dxr-muted)) center / 1px 100% no-repeat; }
`;

export const COLUMNS_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-coldialog" data-dxr="columnsdialog">
  <form method="dialog" novalidate>
    <h2 data-cd-title>Columns</h2>
    <div class="dxr-cdbody">
      <div class="dxr-dgrid">
        <label>Number of columns</label><input type="number" min="1" max="45" step="1" data-cd="count" />
        <label>Spacing</label><input type="text" autocomplete="off" data-cd="space" />
        <label>Apply to</label>
        <select data-cd="applyTo">
          <option value="section">This section</option>
          <option value="forward">This point forward</option>
          <option value="document">Whole document</option>
          <option value="selection" data-cd-selection>Selected text</option>
        </select>
      </div>
      <label class="dxr-dcheck"><input type="checkbox" data-cd="separator" /> Line between</label>
      <div class="dxr-cdpreview" aria-hidden="true" data-cd-preview></div>
    </div>
    <div class="dxr-derror" role="alert"><span data-cd-msg="count" hidden>Type a number of columns from 1 to 45.</span><span data-cd-msg="space" hidden>Type a measurement, such as 0.5" or 1.27 cm.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

export interface ColumnsDialogHost {
  /** The caret's section, or null when there is none (no document, or the caret in a band). */
  read(): SectionInfo | null;
  /** Whether text is selected (offers "Selected text", and starts on it). */
  hasSelection(): boolean;
  /** Write the columns (OK). */
  apply(op: ColumnsOp, applyTo: ColumnsApplyTo): void;
}

/** The Columns dialog over its markup (`COLUMNS_DIALOG_HTML`, already in the page). */
export class ColumnsDialog {
  private unit: MeasurementUnit;
  /** What the section had when the dialog opened: only what the user changed is written. */
  private initial: { count: number; space: number; separator: boolean } = { count: 1, space: 720, separator: false };

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: ColumnsDialogHost, unit: MeasurementUnit = "in") {
    this.unit = unit;
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-cd-title]")!);
    const form = dialog.querySelector("form")!;
    form.addEventListener("input", (event) => {
      (event.target as HTMLElement).removeAttribute("aria-invalid");
      this.showError(null);
      this.preview();
    });
    form.addEventListener("change", () => this.preview());
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.dialog.close("cancel"); return; }
      const result = this.collect();
      if (!result) return;
      this.dialog.close("ok");
      this.host.apply(result.op, result.applyTo);
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.dialog.close("cancel"); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  setUnit(unit: MeasurementUnit): void {
    this.unit = unit;
  }

  /** Open on the caret's section. False (nothing opened) when there is no section to read. */
  show(): boolean {
    const info = this.host.read();
    if (!info) return false;
    this.initial = {
      count: info.columns,
      space: info.columnSpaceTwips ?? 720,
      separator: !!info.columnSeparator,
    };
    this.field<HTMLInputElement>("count").value = String(this.initial.count);
    this.field<HTMLInputElement>("space").value = formatLength(this.initial.space, this.unit);
    this.field<HTMLInputElement>("separator").checked = this.initial.separator;
    const selected = this.host.hasSelection();
    const applyTo = this.field<HTMLSelectElement>("applyTo");
    const selectionOption = applyTo.querySelector<HTMLOptionElement>("[data-cd-selection]")!;
    selectionOption.hidden = selectionOption.disabled = !selected;
    applyTo.value = selected ? "selection" : "section";
    this.showError(null);
    this.preview();
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.field<HTMLInputElement>("count").select();
    return true;
  }

  /** The op the fields say, or null (the bad field marked) when one is not valid. Fields left as
   *  the section had them are left out of the op, except that a new section gets them all. */
  collect(): { op: ColumnsOp; applyTo: ColumnsApplyTo } | null {
    const countField = this.field<HTMLInputElement>("count");
    const count = Number(countField.value.trim());
    if (!Number.isInteger(count) || count < 1 || count > 45) return this.invalid(countField, "count");
    const spaceField = this.field<HTMLInputElement>("space");
    const space = parseLength(spaceField.value, this.unit);
    if (space === null || space < 0) return this.invalid(spaceField, "space");
    const separator = this.field<HTMLInputElement>("separator").checked;
    const applyTo = this.field<HTMLSelectElement>("applyTo").value as ColumnsApplyTo;
    const op: ColumnsOp = {};
    // A count always goes in (it also makes unequal columns equal, as OK does in Word); the
    // space and the line only when changed, so "Whole document" keeps each section's own.
    op.count = count;
    if (space !== this.initial.space || applyTo === "forward" || applyTo === "selection") op.spaceTwips = space;
    if (separator !== this.initial.separator || applyTo === "forward" || applyTo === "selection") op.separator = separator;
    return { op, applyTo };
  }

  private invalid(field: HTMLElement, message: string): null {
    field.setAttribute("aria-invalid", "true");
    this.showError(message);
    field.focus();
    return null;
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-cd-msg]")))
      span.hidden = span.dataset.cdMsg !== message;
  }

  /** A sketch of the page: the columns as text lines, the gaps to scale, the line between. */
  private preview(): void {
    const box = this.dialog.querySelector<HTMLElement>("[data-cd-preview]")!;
    const count = Math.min(Math.max(Math.round(Number(this.field<HTMLInputElement>("count").value) || 1), 1), 12);
    const space = parseLength(this.field<HTMLInputElement>("space").value, this.unit) ?? 720;
    const line = this.field<HTMLInputElement>("separator").checked;
    // 6.5" of text across the sketch's width.
    const gap = `${Math.min(40, (space / 9360) * 100)}%`;
    box.replaceChildren();
    for (let i = 0; i < count; i++) {
      if (i > 0) {
        const g = document.createElement("i");
        g.style.width = gap;
        if (line) g.dataset.line = "";
        box.append(g);
      }
      box.append(document.createElement("span"));
    }
  }

  private field<T extends HTMLElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-cd="${name}"]`)!;
  }
}
