/**
 * The ribbon's Line Numbers dialog — Word's Line Numbers options (Page Setup > Layout > Line
 * Numbers): add line numbering, start at, from text, count by, when numbering restarts, and which
 * sections it applies to. OK is one `setLineNumbering`.
 */
import type { LineNumbering, SectionInfo } from "./types.js";
import { formatLength, makeDialogMovable, parseLength, restoreDialogPlace, type MeasurementUnit } from "./ribbon-paragraph-dialog.js";

export const LINE_NUMBERS_DIALOG_CSS = `
.dxr-lndialog { width: min(380px, calc(100vw - 32px)); }
.dxr-lndialog .dxr-lnbody { padding: 10px 16px 4px; }
.dxr-lndialog .dxr-dgrid input[type="number"] { width: 5em; }
.dxr-lndialog .dxr-dgrid input[type="text"] { width: 6em; }
.dxr-lndialog fieldset { border: 0; margin: 0; padding: 0; min-width: 0; }
.dxr-lndialog fieldset:disabled { opacity: .55; }
.dxr-lndialog .dxr-lnrestart { display: flex; flex-direction: column; gap: 4px; margin: 6px 0; }
`;

export const LINE_NUMBERS_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-lndialog" data-dxr="linenumbersdialog">
  <form method="dialog" novalidate>
    <h2 data-ln-title>Line Numbers</h2>
    <div class="dxr-lnbody">
      <label class="dxr-dcheck"><input type="checkbox" data-ln="on" /> Add line numbering</label>
      <fieldset data-ln-fields>
        <div class="dxr-dgrid">
          <label>Start at</label><input type="number" min="1" max="32767" step="1" data-ln="start" />
          <label>From text</label><input type="text" autocomplete="off" data-ln="distance" placeholder="Auto" />
          <label>Count by</label><input type="number" min="1" max="100" step="1" data-ln="countBy" />
        </div>
        <div class="dxr-lnrestart" role="radiogroup">
          <span>Numbering</span>
          <label><input type="radio" name="lnrestart" value="newPage" data-ln="restart-newPage" /> Restart each page</label>
          <label><input type="radio" name="lnrestart" value="newSection" data-ln="restart-newSection" /> Restart each section</label>
          <label><input type="radio" name="lnrestart" value="continuous" data-ln="restart-continuous" /> Continuous</label>
        </div>
      </fieldset>
      <div class="dxr-dgrid">
        <label>Apply to</label>
        <select data-ln="applyTo">
          <option value="document">Whole document</option>
          <option value="section">This section</option>
        </select>
      </div>
    </div>
    <div class="dxr-derror" role="alert"><span data-ln-msg="start" hidden>Type a starting number from 1 to 32767.</span><span data-ln-msg="countBy" hidden>Type a number to count by from 1 to 100.</span><span data-ln-msg="distance" hidden>Type a measurement, such as 0.25", or leave it empty for Auto.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

export interface LineNumbersDialogHost {
  /** The caret's section, or null when there is none. */
  read(): SectionInfo | null;
  apply(lineNumbering: LineNumbering, applyTo: "section" | "document"): void;
}

/** The Line Numbers dialog over its markup (`LINE_NUMBERS_DIALOG_HTML`, already in the page). */
export class LineNumbersDialog {
  private unit: MeasurementUnit;

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: LineNumbersDialogHost, unit: MeasurementUnit = "in") {
    this.unit = unit;
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-ln-title]")!);
    const form = dialog.querySelector("form")!;
    form.addEventListener("input", (event) => {
      (event.target as HTMLElement).removeAttribute("aria-invalid");
      this.showError(null);
    });
    form.addEventListener("change", () => this.enable());
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.dialog.close("cancel"); return; }
      const result = this.collect();
      if (!result) return;
      this.dialog.close("ok");
      this.host.apply(result, this.field<HTMLSelectElement>("applyTo").value === "document" ? "document" : "section");
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.dialog.close("cancel"); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  setUnit(unit: MeasurementUnit): void {
    this.unit = unit;
  }

  /** Open on the caret's section. False (nothing opened) when there is none. */
  show(): boolean {
    const info = this.host.read();
    if (!info) return false;
    const ln = info.lineNumbering;
    this.input("on").checked = !!ln;
    this.input("start").value = String(ln?.start ?? 1);
    this.input("countBy").value = String(ln?.countBy ?? 1);
    this.input("distance").value = ln?.distanceTwips !== undefined ? formatLength(ln.distanceTwips, this.unit) : "";
    this.input(`restart-${ln?.restart ?? "newPage"}`).checked = true;
    this.field<HTMLSelectElement>("applyTo").value = "document";
    this.enable();
    this.showError(null);
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.input("on").focus();
    return true;
  }

  /** What the fields ask for, or null (the bad field marked) when one is not valid. */
  collect(): LineNumbering | null {
    if (!this.input("on").checked) return { countBy: 0 };
    const start = Number(this.input("start").value);
    if (!Number.isInteger(start) || start < 1 || start > 32767) return this.invalid("start");
    const countBy = Number(this.input("countBy").value);
    if (!Number.isInteger(countBy) || countBy < 1 || countBy > 100) return this.invalid("countBy");
    const text = this.input("distance").value.trim();
    let distanceTwips: number | undefined;
    if (text && !/^auto$/i.test(text)) {
      const twips = parseLength(text, this.unit);
      if (twips === null || twips < 0) return this.invalid("distance");
      distanceTwips = twips;
    }
    const restart = (this.dialog.querySelector<HTMLInputElement>('input[name="lnrestart"]:checked')?.value
      ?? "newPage") as LineNumbering["restart"];
    return { countBy, start, restart, ...(distanceTwips !== undefined ? { distanceTwips } : {}) };
  }

  private enable(): void {
    this.dialog.querySelector<HTMLFieldSetElement>("[data-ln-fields]")!.disabled = !this.input("on").checked;
  }

  private invalid(name: string): null {
    const field = this.input(name);
    field.setAttribute("aria-invalid", "true");
    this.showError(name);
    field.focus();
    return null;
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-ln-msg]")))
      span.hidden = span.dataset.lnMsg !== message;
  }

  private field<T extends HTMLElement = HTMLElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-ln="${name}"]`)!;
  }

  private input(name: string): HTMLInputElement {
    return this.field<HTMLInputElement>(name);
  }
}
