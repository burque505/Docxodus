/**
 * The ribbon's Paragraph dialog — Word's, with its two tabs (Indents and Spacing; Line and Page
 * Breaks) and a preview. It opens on the paragraphs the selection covers: a field where they
 * differ is blank (a check box is mixed), and OK sets only the fields the user changed, so a
 * blank field leaves each paragraph's own value alone. The whole change is one
 * `DocxEditor.setParagraphFormat` call: one undo step, one `paragraphFormat` command. Tabs… (when
 * the host gives `openTabs`) does what OK does, then opens the Tabs dialog, as in Word.
 *
 * Lengths show in inches or centimetres (`MeasurementUnit`); any field takes a typed unit too
 * ("2 cm", "1 in", "0.5\"", "12 pt", "15 mm", "1 pi"), with a decimal point or comma.
 */
import type { EditorAlignment, EditorParagraphFormat } from "./editor.js";

/** How the dialog shows indents: inches (`"in"`) or centimetres (`"cm"`). */
export type MeasurementUnit = "in" | "cm";

const TWIPS_PER: Record<string, number> = {
  in: 1440, inch: 1440, inches: 1440, '"': 1440, "″": 1440,
  cm: 1440 / 2.54, mm: 144 / 2.54, pt: 20, pi: 240,
};

/** Twips from a typed length; a bare number is in `unit`. Null when it is not a length. */
export function parseLength(text: string, unit: string): number | null {
  const m = /^\s*([-+]?(?:\d+(?:[.,]\d*)?|[.,]\d+))\s*(in|inch|inches|"|″|cm|mm|pt|pi)?\s*$/i.exec(text);
  if (!m) return null;
  const value = Number(m[1].replace(",", "."));
  const per = TWIPS_PER[(m[2] ?? unit).toLowerCase()];
  return Number.isFinite(value) && per ? Math.round(value * per) : null;
}

function trimNumber(n: number, places: number): string {
  return String(Number(n.toFixed(places)));
}

/** A length as the dialog shows it: `0.5"` or `1.27 cm`. */
export function formatLength(twips: number, unit: MeasurementUnit): string {
  return unit === "cm" ? `${trimNumber(twips / (1440 / 2.54), 2)} cm` : `${trimNumber(twips / 1440, 2)}"`;
}

/** Drag a ribbon dialog by `handle` (its title), off the text it is changing. */
export function makeDialogMovable(dialog: HTMLDialogElement, handle: HTMLElement): void {
  handle.style.cursor = "move";
  handle.style.userSelect = "none";
  handle.addEventListener("pointerdown", (down) => {
    if (down.button !== 0) return;
    const box = dialog.getBoundingClientRect();
    const dx = down.clientX - box.left, dy = down.clientY - box.top;
    handle.setPointerCapture(down.pointerId);
    const move = (e: PointerEvent) => placeDialogAt(dialog, e.clientX - dx, e.clientY - dy);
    const up = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", up);
      handle.removeEventListener("pointercancel", up);
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", up);
    handle.addEventListener("pointercancel", up);
    down.preventDefault();
  });
}

/** Put a ribbon dialog at a viewport position, its title kept reachable on screen. */
export function placeDialogAt(dialog: HTMLDialogElement, left: number, top: number): void {
  const view = dialog.ownerDocument.defaultView ?? window;
  const box = dialog.getBoundingClientRect();
  const x = Math.min(Math.max(left, 40 - box.width), view.innerWidth - 40);
  const y = Math.min(Math.max(top, 0), view.innerHeight - 40);
  dialog.style.margin = "0";
  dialog.style.left = `${x}px`;
  dialog.style.top = `${y}px`;
}

/** Where the user left a moved dialog, put back on screen after it opens again. */
export function restoreDialogPlace(dialog: HTMLDialogElement): void {
  if (dialog.style.left) placeDialogAt(dialog, parseFloat(dialog.style.left), parseFloat(dialog.style.top));
}

export function formatPoints(twips: number): string {
  return `${trimNumber(twips / 20, 1)} pt`;
}

type LineChoice = "single" | "onehalf" | "double" | "atLeast" | "exact" | "multiple";

function lineChoiceOf(f: EditorParagraphFormat): { choice: LineChoice; at: string } {
  const line = f.lineSpacing ?? 240;
  if (f.lineSpacingRule === "exact") return { choice: "exact", at: formatPoints(line) };
  if (f.lineSpacingRule === "atLeast") return { choice: "atLeast", at: formatPoints(line) };
  if (line === 240) return { choice: "single", at: "" };
  if (line === 360) return { choice: "onehalf", at: "" };
  if (line === 480) return { choice: "double", at: "" };
  return { choice: "multiple", at: trimNumber(line / 240, 2) };
}

/** One value when every paragraph agrees, else undefined (a blank, mixed field). */
function common<T>(values: T[]): T | undefined {
  return values.length > 0 && values.every((v) => v === values[0]) ? values[0] : undefined;
}

export const PARAGRAPH_DIALOG_CSS = `
.dxr-dialog {
  width: min(470px, calc(100vw - 32px));
  max-height: calc(100vh - 32px);
  padding: 0;
  border: 1px solid var(--dxr-rule);
  border-radius: 12px;
  background: var(--dxr-chrome);
  color: var(--dxr-ink);
  font: 12.5px/1.4 var(--dxr-ui);
  box-shadow: var(--dxr-shadow-xl);
}
.dxr-dialog::backdrop { background: rgba(15, 23, 42, .18); }
.dxr-dialog form { display: flex; flex-direction: column; }
.dxr-dialog h2 { margin: 0; padding: 12px 16px 0; font-size: 14px; font-weight: 600; touch-action: none; }
.dxr-dialog [role="tablist"] { display: flex; gap: 2px; padding: 8px 12px 0; border-bottom: 1px solid var(--dxr-rule); }
.dxr-dialog [role="tab"] {
  min-height: 30px; padding: 4px 10px; border: none; border-bottom: 2px solid transparent;
  border-radius: 0; background: none; color: var(--dxr-muted); font: inherit; cursor: pointer;
}
.dxr-dialog [role="tab"][aria-selected="true"] { border-bottom-color: var(--dxr-accent); color: var(--dxr-ink); }
.dxr-dialog [role="tabpanel"] { padding: 10px 16px 4px; }
.dxr-dialog [role="tabpanel"][hidden] { display: none; }
.dxr-dialog fieldset { margin: 0 0 10px; padding: 0; border: none; }
.dxr-dialog legend {
  margin-bottom: 5px; padding: 0; color: var(--dxr-faint);
  font-size: 9.5px; font-weight: 600; letter-spacing: .11em; text-transform: uppercase;
}
.dxr-dialog .dxr-dgrid { display: grid; grid-template-columns: auto 1fr auto 1fr; gap: 6px 10px; align-items: center; }
.dxr-dialog .dxr-dgrid label { color: var(--dxr-muted); }
.dxr-dialog input[type="text"], .dxr-dialog select {
  min-width: 0; height: 28px; padding: 2px 6px; border: 1px solid var(--dxr-rule-strong);
  border-radius: 6px; background: var(--dxr-sheet); color: var(--dxr-ink); font: inherit;
}
.dxr-dialog input[aria-invalid="true"] { border-color: var(--dxr-danger); background: var(--dxr-danger-wash); }
.dxr-dialog .dxr-dcheck { display: flex; align-items: center; gap: 7px; margin: 5px 0; }
.dxr-dialog .dxr-dpreview {
  height: 112px; overflow: hidden; padding: 8px 10px; border: 1px solid var(--dxr-rule);
  border-radius: 6px; background: var(--dxr-sheet); font: 6px/1.2 var(--dxr-serif); color: var(--dxr-faint);
}
.dxr-dialog .dxr-dpreview p { margin: 0; }
.dxr-dialog .dxr-dpreview [data-sample] { color: var(--dxr-ink); }
.dxr-dialog .dxr-derror { min-height: 16px; padding: 0 16px; color: var(--dxr-danger); font-size: 11.5px; }
.dxr-dialog .dxr-dfoot { display: flex; justify-content: flex-end; gap: 8px; padding: 6px 16px 14px; }
.dxr-dialog .dxr-dfoot button {
  min-width: 76px; min-height: 30px; border: 1px solid var(--dxr-rule-strong); border-radius: 6px;
  background: var(--dxr-sheet); color: var(--dxr-ink); font: inherit; cursor: pointer;
}
.dxr-dialog .dxr-dfoot [data-pd-tabs] { margin-right: auto; }
.dxr-dialog .dxr-dfoot [data-pd-tabs][hidden] { display: none; }
.dxr-dialog .dxr-dfoot button[value="ok"] { border-color: var(--dxr-accent); background: var(--dxr-accent); color: #fff; }
.dxr-dialog .dxr-dfoot button[value="ok"]:hover { background: var(--dxr-accent-hover); }
.dxr-glabel .dxr-launch {
  min-height: 0; margin-left: 6px; padding: 0 2px; vertical-align: -1px; color: var(--dxr-faint); line-height: 1;
}
.dxr-glabel .dxr-launch:hover { color: var(--dxr-ink); }
`;

const SAMPLE = "Sample Text ".repeat(28);
const OTHER = "Previous Paragraph ".repeat(14);
const NEXT = "Following Paragraph ".repeat(14);

/** The dialog's markup; the ribbon puts it after its own (ids come from `data-dxr`). */
export const PARAGRAPH_DIALOG_HTML = `
<dialog class="dxr-dialog" data-dxr="paradialog" aria-labelledby="dxr-paradialog-title">
  <form method="dialog" data-dxr="paraform">
    <h2 data-pd-title>Paragraph</h2>
    <div role="tablist">
      <button type="button" role="tab" data-pd-tab="indents" aria-selected="true">Indents and Spacing</button>
      <button type="button" role="tab" data-pd-tab="breaks" aria-selected="false">Line and Page Breaks</button>
    </div>
    <div role="tabpanel" data-pd-panel="indents">
      <fieldset>
        <legend>General</legend>
        <div class="dxr-dgrid">
          <label>Alignment</label>
          <select data-pd="alignment">
            <option value="" hidden></option>
            <option value="left">Left</option>
            <option value="center">Centered</option>
            <option value="right">Right</option>
            <option value="justify">Justified</option>
          </select>
        </div>
      </fieldset>
      <fieldset>
        <legend>Indentation</legend>
        <div class="dxr-dgrid">
          <label>Left</label><input type="text" autocomplete="off" data-pd="left" />
          <label>Special</label>
          <select data-pd="special">
            <option value="" hidden></option>
            <option value="none">(none)</option>
            <option value="firstLine">First line</option>
            <option value="hanging">Hanging</option>
          </select>
          <label>Right</label><input type="text" autocomplete="off" data-pd="right" />
          <label>By</label><input type="text" autocomplete="off" data-pd="by" />
        </div>
      </fieldset>
      <fieldset>
        <legend>Spacing</legend>
        <div class="dxr-dgrid">
          <label>Before</label><input type="text" autocomplete="off" data-pd="before" />
          <label>Line spacing</label>
          <select data-pd="line">
            <option value="" hidden></option>
            <option value="single">Single</option>
            <option value="onehalf">1.5 lines</option>
            <option value="double">Double</option>
            <option value="atLeast">At least</option>
            <option value="exact">Exactly</option>
            <option value="multiple">Multiple</option>
          </select>
          <label>After</label><input type="text" autocomplete="off" data-pd="after" />
          <label>At</label><input type="text" autocomplete="off" data-pd="at" />
        </div>
        <label class="dxr-dcheck"><input type="checkbox" data-pd="contextualSpacing" /> Don't add space between paragraphs of the same style</label>
      </fieldset>
      <fieldset>
        <legend>Preview</legend>
        <div class="dxr-dpreview" aria-hidden="true">
          <p>${OTHER}</p>
          <p data-sample>${SAMPLE}</p>
          <p data-next>${NEXT}</p>
        </div>
      </fieldset>
    </div>
    <div role="tabpanel" data-pd-panel="breaks" hidden>
      <fieldset>
        <legend>Pagination</legend>
        <label class="dxr-dcheck"><input type="checkbox" data-pd="widowControl" /> Widow/Orphan control</label>
        <label class="dxr-dcheck"><input type="checkbox" data-pd="keepNext" /> Keep with next</label>
        <label class="dxr-dcheck"><input type="checkbox" data-pd="keepLines" /> Keep lines together</label>
        <label class="dxr-dcheck"><input type="checkbox" data-pd="pageBreakBefore" /> Page break before</label>
      </fieldset>
    </div>
    <div class="dxr-derror" data-pd-error role="alert"><span data-pd-msg="length" hidden>Type a measurement, such as 1.5" or 2 cm.</span><span data-pd-msg="points" hidden>Type a size in points, such as 6 pt, or 0.</span><span data-pd-msg="multiple" hidden>Type how many lines, such as 1.15.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="tabs" data-pd-tabs hidden>Tabs…</button>
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

const TOGGLES = ["contextualSpacing", "widowControl", "keepNext", "keepLines", "pageBreakBefore"] as const;
type Toggle = (typeof TOGGLES)[number];

export interface ParagraphDialogHost {
  /** The selected paragraphs' formats (DocxEditor.paragraphFormats). */
  read(): EditorParagraphFormat[];
  /** Apply what the user changed; called after the dialog has closed. */
  apply(op: EditorParagraphFormat): void;
  /** Called when the dialog closes without a change (Cancel, Esc, or OK with nothing changed). */
  cancelled?(): void;
  /** Word's Tabs… button: called after the dialog has closed and applied what changed (as OK
   *  does), to open the Tabs dialog. Without it the dialog has no Tabs… button. */
  openTabs?(): void;
}

/** The Paragraph dialog over its markup (`PARAGRAPH_DIALOG_HTML`, already in the page). */
export class ParagraphDialog {
  private unit: MeasurementUnit;
  private initial = new Map<string, string>();

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: ParagraphDialogHost, unit: MeasurementUnit = "in") {
    this.unit = unit;
    const form = dialog.querySelector("form")!;
    for (const tab of Array.from(dialog.querySelectorAll<HTMLElement>("[data-pd-tab]"))) {
      tab.addEventListener("click", () => this.showTab(tab.dataset.pdTab!));
    }
    const base = dialog.id || "dxr-paradialog";
    dialog.querySelector("[data-pd-title]")!.id = `${base}-title`;
    dialog.setAttribute("aria-labelledby", `${base}-title`);
    // Each grid label names the field after it.
    for (const label of Array.from(dialog.querySelectorAll<HTMLLabelElement>(".dxr-dgrid label"))) {
      const field = label.nextElementSibling as HTMLElement | null;
      if (!field?.dataset.pd) continue;
      field.id = `${base}-${field.dataset.pd}`;
      label.htmlFor = field.id;
    }
    this.makeMovable(dialog.querySelector<HTMLElement>("[data-pd-title]")!);
    this.field("special").addEventListener("change", () => this.specialChanged());
    this.field("line").addEventListener("change", () => this.lineChanged());
    form.addEventListener("input", (event) => {
      (event.target as HTMLElement).removeAttribute("aria-invalid");
      this.showError(null);
      this.preview();
    });
    form.addEventListener("change", () => this.preview());
    // OK and Cancel are handled here, at once (the dialog's own close event comes a task later):
    // a bad entry keeps the dialog open on that field.
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.finish(null); return; }
      const op = this.collect();
      if (op === null) return;
      this.finish(op);
      // Tabs… is OK, then the Tabs dialog (Word): its stops are set on the paragraphs as now formatted.
      if (submitter?.value === "tabs") this.host.openTabs?.();
    });
    this.dialog.querySelector<HTMLElement>("[data-pd-tabs]")!.hidden = !host.openTabs;
    // Esc closes the dialog by itself.
    dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      this.finish(null);
    });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  setUnit(unit: MeasurementUnit): void {
    this.unit = unit;
  }

  /** Open on the selection's paragraphs; false when there is none. */
  show(): boolean {
    const formats = this.host.read();
    if (formats.length === 0) return false;
    const pick = <K extends keyof EditorParagraphFormat>(key: K) => common(formats.map((f) => f[key]));
    const len = (v: number | undefined) => (v === undefined ? "" : formatLength(v, this.unit));
    const pts = (v: number | undefined) => (v === undefined ? "" : formatPoints(v));

    this.setValue("alignment", pick("alignment") ?? "");
    this.setValue("left", len(pick("leftIndent") ?? (formats.every((f) => f.leftIndent === undefined) ? 0 : undefined)));
    this.setValue("right", len(pick("rightIndent") ?? (formats.every((f) => f.rightIndent === undefined) ? 0 : undefined)));
    const specials = formats.map((f) =>
      f.hangingIndent ? `hanging:${f.hangingIndent}` : f.firstLineIndent ? `firstLine:${f.firstLineIndent}` : "none:0");
    const special = common(specials);
    if (special === undefined) {
      this.setValue("special", "");
      this.setValue("by", "");
    } else {
      const [kind, twips] = special.split(":");
      this.setValue("special", kind);
      this.setValue("by", kind === "none" ? "" : len(Number(twips)));
    }
    this.setValue("before", pts(pick("spacingBefore") ?? (formats.every((f) => f.spacingBefore === undefined) ? 0 : undefined)));
    this.setValue("after", pts(pick("spacingAfter") ?? (formats.every((f) => f.spacingAfter === undefined) ? 0 : undefined)));
    const lines = formats.map(lineChoiceOf);
    const choice = common(lines.map((l) => `${l.choice}|${l.at}`));
    this.setValue("line", choice === undefined ? "" : choice.split("|")[0]);
    this.setValue("at", choice === undefined ? "" : choice.split("|")[1]);
    for (const key of TOGGLES) {
      const box = this.field<HTMLInputElement>(key);
      const value = common(formats.map((f) => !!f[key]));
      box.indeterminate = value === undefined;
      box.checked = value === true;
    }
    this.initial = this.snapshot();
    for (const el of Array.from(this.dialog.querySelectorAll("[aria-invalid]"))) el.removeAttribute("aria-invalid");
    this.showError(null);
    this.showTab("indents");
    this.preview();
    this.dialog.returnValue = "";
    this.dialog.showModal();
    // Where the user left it, still on screen if the window has shrunk since.
    if (this.dialog.style.left) this.placeAt(parseFloat(this.dialog.style.left), parseFloat(this.dialog.style.top));
    this.field("alignment").focus();
    return true;
  }

  /** What OK would apply, or null (with the bad field marked) when an entry is not valid. */
  collect(): EditorParagraphFormat | null {
    const now = this.snapshot();
    const changed = (...keys: string[]) => keys.some((k) => now.get(k) !== this.initial.get(k));
    const op: EditorParagraphFormat = {};
    const bad = (key: string, message: "length" | "points" | "multiple"): null => {
      const el = this.field(key);
      el.setAttribute("aria-invalid", "true");
      this.showTab(el.closest<HTMLElement>("[data-pd-panel]")!.dataset.pdPanel!);
      el.focus();
      this.showError(message);
      return null;
    };
    const length = (key: string, min: number, message: "length" | "points"): number | null | undefined => {
      const text = now.get(key)!;
      if (text.trim() === "") return undefined; // a blank (mixed) field stays as it is
      const twips = parseLength(text, message === "points" ? "pt" : this.unit);
      // Word's limits: indents within ±22", spacing up to 1584 pt (both 31680 twips).
      if (twips === null || twips < min || Math.abs(twips) > 31680) return null;
      return twips;
    };

    if (changed("alignment") && now.get("alignment")) op.alignment = now.get("alignment") as EditorAlignment;
    for (const [key, field] of [["leftIndent", "left"], ["rightIndent", "right"]] as const) {
      if (!changed(field)) continue;
      const v = length(field, -31680, "length");
      if (v === null) return bad(field, "length");
      if (v !== undefined) op[key] = v;
    }
    if (changed("special", "by")) {
      const kind = now.get("special");
      if (kind === "none") op.firstLineIndent = 0;
      else if (kind === "firstLine" || kind === "hanging") {
        const v = length("by", 0, "length");
        if (v === null) return bad("by", "length");
        if (v !== undefined) op[kind === "firstLine" ? "firstLineIndent" : "hangingIndent"] = v;
      }
    }
    for (const [key, field] of [["spacingBefore", "before"], ["spacingAfter", "after"]] as const) {
      if (!changed(field)) continue;
      const v = length(field, 0, "points");
      if (v === null) return bad(field, "points");
      if (v !== undefined) op[key] = v;
    }
    if (changed("line", "at")) {
      const choice = now.get("line") as LineChoice | "";
      const fixed = { single: 240, onehalf: 360, double: 480 } as Record<string, number>;
      if (choice && fixed[choice]) {
        op.lineSpacing = fixed[choice];
        op.lineSpacingRule = "auto";
      } else if (choice === "atLeast" || choice === "exact") {
        const v = length("at", 1, "points");
        if (v === null || v === undefined) return bad("at", "points");
        op.lineSpacing = v;
        op.lineSpacingRule = choice;
      } else if (choice === "multiple") {
        const n = Number(now.get("at")!.trim().replace(",", "."));
        if (!(n > 0 && n <= 132)) return bad("at", "multiple");
        op.lineSpacing = Math.round(n * 240);
        op.lineSpacingRule = "auto";
      }
    }
    for (const key of TOGGLES) {
      if (!changed(key)) continue;
      const box = this.field<HTMLInputElement>(key);
      if (!box.indeterminate) op[key as Toggle] = box.checked;
    }
    return op;
  }

  // ── internals ──

  private makeMovable(handle: HTMLElement): void {
    makeDialogMovable(this.dialog, handle);
  }

  private placeAt(left: number, top: number): void {
    placeDialogAt(this.dialog, left, top);
  }

  private finish(op: EditorParagraphFormat | null): void {
    this.dialog.close(op ? "ok" : "cancel");
    if (op && Object.keys(op).length > 0) this.host.apply(op);
    else this.host.cancelled?.();
  }

  private field<T extends HTMLElement = HTMLInputElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-pd="${name}"]`)!;
  }

  private setValue(name: string, value: string): void {
    (this.field<HTMLInputElement | HTMLSelectElement>(name)).value = value;
  }

  private snapshot(): Map<string, string> {
    const out = new Map<string, string>();
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLInputElement | HTMLSelectElement>("[data-pd]"))) {
      const key = el.dataset.pd!;
      out.set(key, el instanceof HTMLInputElement && el.type === "checkbox"
        ? (el.indeterminate ? "mixed" : String(el.checked))
        : el.value);
    }
    return out;
  }

  private showTab(name: string): void {
    for (const tab of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-pd-tab]"))) {
      tab.setAttribute("aria-selected", String(tab.dataset.pdTab === name));
    }
    for (const panel of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-pd-panel]"))) {
      panel.hidden = panel.dataset.pdPanel !== name;
    }
  }

  private showError(which: string | null): void {
    for (const msg of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-pd-msg]"))) {
      msg.hidden = msg.dataset.pdMsg !== which;
    }
  }

  /** Word's defaults when Special changes: First line and Hanging start at half an inch. */
  private specialChanged(): void {
    const kind = this.field<HTMLSelectElement>("special").value;
    const by = this.field("by");
    if (kind === "none") by.value = "";
    else if (by.value.trim() === "") by.value = formatLength(this.unit === "cm" ? 709 : 720, this.unit); // 1.25 cm or 0.5"
  }

  /** Word's defaults when Line spacing changes: 12 pt for At least / Exactly, 3 for Multiple. */
  private lineChanged(): void {
    const choice = this.field<HTMLSelectElement>("line").value;
    const at = this.field("at");
    if (choice === "atLeast" || choice === "exact") {
      if (parseLength(at.value, "pt") === null || !/pt\s*$/i.test(at.value)) at.value = "12 pt";
    } else if (choice === "multiple") {
      if (!(Number(at.value.replace(",", ".")) > 0)) at.value = "3";
    } else at.value = "";
  }

  /** The sample paragraph drawn as the fields say, at about a fortieth of the page. */
  private preview(): void {
    const sample = this.dialog.querySelector<HTMLElement>("[data-sample]")!;
    const px = (key: string, unit: string, scale: number) => {
      const t = parseLength(this.field(key).value, unit);
      return t === null ? 0 : t * scale;
    };
    const indentScale = 36 / 1440; // 36 px an inch
    const pointScale = 0.5 / 20; // half a pixel a point
    const special = this.field<HTMLSelectElement>("special").value;
    const by = special === "none" ? 0 : px("by", this.unit, indentScale);
    const left = px("left", this.unit, indentScale);
    sample.style.marginLeft = `${left + (special === "hanging" ? by : 0)}px`;
    sample.style.marginRight = `${px("right", this.unit, indentScale)}px`;
    sample.style.textIndent = `${special === "hanging" ? -by : special === "firstLine" ? by : 0}px`;
    sample.style.marginTop = `${px("before", "pt", pointScale)}px`;
    sample.style.marginBottom = `${px("after", "pt", pointScale)}px`;
    const align = this.field<HTMLSelectElement>("alignment").value;
    sample.style.textAlign = align === "justify" ? "justify" : align || "left";
    const choice = this.field<HTMLSelectElement>("line").value;
    const at = this.field("at").value;
    const multiple = { single: 1, onehalf: 1.5, double: 2 }[choice as "single"]
      ?? (choice === "multiple" ? Number(at.replace(",", ".")) || 1 : undefined);
    if (multiple !== undefined) sample.style.lineHeight = String(1.2 * multiple);
    else if (choice === "atLeast" || choice === "exact") {
      const height = px("at", "pt", 0.5 / 20) || 7.2;
      sample.style.lineHeight = `${choice === "atLeast" ? Math.max(height, 7.2) : height}px`;
    } else sample.style.lineHeight = "";
  }
}
