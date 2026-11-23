/**
 * The ribbon's Font dialog — Word's Font (Home > Font launcher, Ctrl+D): font, style, size, color,
 * underline and effects, with a preview. On the Home tab OK formats the selection; from Modify
 * Style's Format menu it formats the style. Only what the user changed is applied.
 */
import type { FormatOp, RunFormattingInfo } from "./types.js";
import { makeDialogMovable, restoreDialogPlace } from "./ribbon-paragraph-dialog.js";

export const FONT_DIALOG_CSS = `
.dxr-fontdialog { width: min(440px, calc(100vw - 32px)); }
.dxr-fontdialog .dxr-fontbody { padding: 10px 16px 4px; }
.dxr-fontdialog .dxr-fontrow { display: grid; grid-template-columns: 1fr 9em 5em; gap: 4px 10px; }
.dxr-fontdialog .dxr-fontrow label { color: var(--dxr-muted); }
.dxr-fontdialog .dxr-fontrow input, .dxr-fontdialog .dxr-fontrow select { width: 100%; box-sizing: border-box; }
.dxr-fontdialog .dxr-fontmid { display: grid; grid-template-columns: 1fr 1fr; gap: 4px 10px; margin-top: 10px; }
.dxr-fontdialog .dxr-fontmid label { color: var(--dxr-muted); }
.dxr-fontdialog .dxr-fontmid input[type="color"] { width: 100%; height: 28px; padding: 1px; box-sizing: border-box; }
.dxr-fontdialog .dxr-fonteffects { display: grid; grid-template-columns: 1fr 1fr; gap: 0 10px; margin-top: 10px; }
.dxr-fontdialog .dxr-fontpreview {
  display: flex; align-items: center; justify-content: center; height: 56px; margin-top: 6px; overflow: hidden;
  border: 1px solid var(--dxr-rule); border-radius: 6px; background: var(--dxr-sheet); color: #000; white-space: nowrap;
}
`;

export const FONT_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-fontdialog" data-dxr="fontdialog">
  <form method="dialog" novalidate>
    <h2 data-fd-title>Font</h2>
    <div class="dxr-fontbody">
      <div class="dxr-fontrow">
        <label for="dxr-fd-font">Font</label><label for="dxr-fd-style">Font style</label><label for="dxr-fd-size">Size</label>
        <input id="dxr-fd-font" type="text" autocomplete="off" list="dxr-fd-fonts" data-fd="font" />
        <select id="dxr-fd-style" data-fd="style">
          <option value="" hidden></option>
          <option value="regular">Regular</option>
          <option value="italic">Italic</option>
          <option value="bold">Bold</option>
          <option value="bolditalic">Bold Italic</option>
        </select>
        <input id="dxr-fd-size" type="text" autocomplete="off" list="dxr-fd-sizes" data-fd="size" />
      </div>
      <datalist id="dxr-fd-fonts" data-fd-fonts></datalist>
      <datalist id="dxr-fd-sizes">
        <option>8</option><option>9</option><option>10</option><option>11</option><option>12</option><option>14</option>
        <option>16</option><option>18</option><option>20</option><option>24</option><option>28</option><option>36</option>
        <option>48</option><option>72</option>
      </datalist>
      <div class="dxr-fontmid">
        <label for="dxr-fd-color">Font color</label><label for="dxr-fd-underline">Underline style</label>
        <input id="dxr-fd-color" type="color" data-fd="color" />
        <select id="dxr-fd-underline" data-fd="underline">
          <option value="" hidden></option>
          <option value="none">(none)</option>
          <option value="single">Single</option>
        </select>
      </div>
      <fieldset class="dxr-fonteffects">
        <legend>Effects</legend>
        <label class="dxr-dcheck"><input type="checkbox" data-fd="strike" /> Strikethrough</label>
        <label class="dxr-dcheck"><input type="checkbox" data-fd="smallCaps" /> Small caps</label>
        <label class="dxr-dcheck"><input type="checkbox" data-fd="superscript" /> Superscript</label>
        <label class="dxr-dcheck"><input type="checkbox" data-fd="caps" /> All caps</label>
        <label class="dxr-dcheck"><input type="checkbox" data-fd="subscript" /> Subscript</label>
      </fieldset>
      <fieldset>
        <legend>Preview</legend>
        <div class="dxr-fontpreview" aria-hidden="true"><span data-fd-sample></span></div>
      </fieldset>
    </div>
    <div class="dxr-derror" role="alert"><span data-fd-msg="size" hidden>Type a size from 1 to 1638 points.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

export interface FontDialogHost {
  /** The formatting to show (the selection's, or a style's), or null when there is nothing. */
  read(): RunFormattingInfo | null;
  /** The fonts to offer. */
  fonts(): string[];
  /** Apply what the user changed; called after the dialog has closed. */
  apply(op: FormatOp): void;
  cancelled?(): void;
}

const CHECKS = ["strike", "smallCaps", "caps", "superscript", "subscript"] as const;

/** `#rrggbb` for a w:color value ("auto" and unknown values read as black). */
function colorInput(color: string | undefined): string {
  return color && /^[0-9a-f]{6}$/i.test(color) ? `#${color.toLowerCase()}` : "#000000";
}

/** The Font dialog over its markup (`FONT_DIALOG_HTML`, already in the page). */
export class FontDialog {
  private initial = new Map<string, string>();
  private sampleText = "";
  /** The host of the current opening (the Home tab's, or Modify Style's for a style). */
  private host: FontDialogHost;

  constructor(private readonly dialog: HTMLDialogElement, host: FontDialogHost) {
    this.host = host;
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-fd-title]")!);
    const form = dialog.querySelector("form")!;
    form.addEventListener("input", (event) => {
      (event.target as HTMLElement).removeAttribute("aria-invalid");
      this.showError(null);
      this.preview();
    });
    form.addEventListener("change", (event) => {
      // Superscript and subscript exclude each other, as small caps and all caps do.
      const target = event.target as HTMLInputElement;
      const other = { superscript: "subscript", subscript: "superscript", caps: "smallCaps", smallCaps: "caps" }[target.dataset.fd ?? ""];
      if (other && target.checked) {
        const box = this.field(other);
        box.checked = false;
        box.indeterminate = false;
      }
      this.preview();
    });
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.finish(null); return; }
      const op = this.collect();
      if (op === null) return;
      this.finish(op);
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.finish(null); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  /** Open on the host's formatting (`host` replaces the default one for this opening: Modify
   *  Style edits a style through it). False when there is nothing to format. */
  show(host?: FontDialogHost, sample?: string): boolean {
    if (host) this.host = host;
    const f = this.host.read();
    if (!f) return false;
    const list = this.dialog.querySelector<HTMLDataListElement>("[data-fd-fonts]")!;
    list.replaceChildren(...this.host.fonts().map((name) => Object.assign(document.createElement("option"), { value: name })));
    this.value("font").value = f.fontFamily ?? "";
    (this.field("style") as unknown as HTMLSelectElement).value = f.bold ? (f.italic ? "bolditalic" : "bold") : f.italic ? "italic" : "regular";
    this.value("size").value = f.fontSizePts !== undefined ? String(f.fontSizePts) : "";
    this.value("color").value = colorInput(f.color);
    (this.field("underline") as unknown as HTMLSelectElement).value = f.underline ? "single" : "none";
    const checks: Record<(typeof CHECKS)[number], boolean> = {
      strike: !!f.strike, smallCaps: !!f.smallCaps, caps: !!f.caps,
      superscript: f.vertAlign === "superscript", subscript: f.vertAlign === "subscript",
    };
    for (const key of CHECKS) {
      const box = this.field(key);
      box.indeterminate = false;
      box.checked = checks[key];
    }
    this.sampleText = sample?.trim() || f.fontFamily || "AaBbYyZz";
    this.initial = this.snapshot();
    this.showError(null);
    this.preview();
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.value("font").focus();
    this.value("font").select();
    return true;
  }

  /** What OK would apply, or null (with the bad field marked) when an entry is not valid. */
  collect(): FormatOp | null {
    const now = this.snapshot();
    const changed = (key: string) => now.get(key) !== this.initial.get(key);
    const op: FormatOp = {};
    if (changed("font") && now.get("font")!.trim()) op.fontFamily = now.get("font")!.trim();
    if (changed("style")) {
      const style = now.get("style")!;
      op.bold = style === "bold" || style === "bolditalic";
      op.italic = style === "italic" || style === "bolditalic";
    }
    if (changed("size")) {
      const pts = Number(now.get("size")!.trim().replace(",", "."));
      if (!(pts >= 1 && pts <= 1638)) {
        const el = this.value("size");
        el.setAttribute("aria-invalid", "true");
        el.focus();
        this.showError("size");
        return null;
      }
      op.fontSizePts = Math.round(pts * 2) / 2;
    }
    if (changed("color")) op.color = now.get("color")!.slice(1).toUpperCase();
    if (changed("underline")) op.underline = now.get("underline") === "single";
    if (changed("strike")) op.strike = this.field("strike").checked;
    if (changed("smallCaps")) op.smallCaps = this.field("smallCaps").checked;
    if (changed("caps")) op.caps = this.field("caps").checked;
    if (changed("superscript") || changed("subscript")) {
      op.vertAlign = this.field("superscript").checked ? "superscript" : this.field("subscript").checked ? "subscript" : "baseline";
    }
    return op;
  }

  // ── internals ──

  private finish(op: FormatOp | null): void {
    this.dialog.close(op ? "ok" : "cancel");
    if (op && Object.keys(op).length > 0) this.host.apply(op);
    else this.host.cancelled?.();
  }

  private field(name: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`[data-fd="${name}"]`)!;
  }

  private value(name: string): HTMLInputElement {
    return this.field(name);
  }

  private snapshot(): Map<string, string> {
    const out = new Map<string, string>();
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLInputElement | HTMLSelectElement>("[data-fd]"))) {
      out.set(el.dataset.fd!, el instanceof HTMLInputElement && el.type === "checkbox" ? String(el.checked) : el.value);
    }
    return out;
  }

  private showError(which: string | null): void {
    for (const msg of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-fd-msg]"))) msg.hidden = msg.dataset.fdMsg !== which;
  }

  private preview(): void {
    const sample = this.dialog.querySelector<HTMLElement>("[data-fd-sample]")!;
    sample.textContent = this.sampleText;
    const style = (this.field("style") as unknown as HTMLSelectElement).value;
    const size = Number(this.value("size").value.replace(",", "."));
    sample.style.fontFamily = this.value("font").value ? `"${this.value("font").value}"` : "";
    sample.style.fontWeight = style === "bold" || style === "bolditalic" ? "700" : "400";
    sample.style.fontStyle = style === "italic" || style === "bolditalic" ? "italic" : "normal";
    sample.style.fontSize = size >= 1 ? `${Math.min(size, 36)}pt` : "";
    sample.style.color = this.value("color").value;
    sample.style.textDecorationLine = [
      (this.field("underline") as unknown as HTMLSelectElement).value === "single" ? "underline" : "",
      this.field("strike").checked ? "line-through" : "",
    ].filter(Boolean).join(" ") || "none";
    sample.style.fontVariant = this.field("smallCaps").checked ? "small-caps" : "normal";
    sample.style.textTransform = this.field("caps").checked ? "uppercase" : "none";
    sample.style.verticalAlign = this.field("superscript").checked ? "super" : this.field("subscript").checked ? "sub" : "baseline";
  }
}
