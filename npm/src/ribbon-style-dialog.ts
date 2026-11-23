/**
 * The ribbon's style dialog — Word's Modify Style and Create New Style from Formatting: name, type,
 * based on, style for the following paragraph, a row of common formatting, a preview with a
 * description, "Add to the Styles gallery", and a Format menu that opens the Font and Paragraph
 * dialogs onto the style. OK is one `setStyle`.
 */
import type { EditorParagraphFormat } from "./editor.js";
import type { FormatOp, ParagraphFormatting, RunFormattingInfo, StyleDefinitionOp, StyleInfo } from "./types.js";
import type { FontDialogHost } from "./ribbon-font-dialog.js";
import { formatPoints, makeDialogMovable, restoreDialogPlace, type ParagraphDialogHost } from "./ribbon-paragraph-dialog.js";

export const STYLE_DIALOG_CSS = `
.dxr-styledialog { width: min(520px, calc(100vw - 32px)); }
.dxr-styledialog .dxr-stylebody { padding: 10px 16px 4px; }
.dxr-styledialog .dxr-dgrid { grid-template-columns: auto 1fr; }
.dxr-styledialog .dxr-dgrid input, .dxr-styledialog .dxr-dgrid select { width: 100%; box-sizing: border-box; }
.dxr-styledialog .dxr-stylebar { display: flex; flex-wrap: wrap; align-items: center; gap: 4px; margin: 4px 0; }
.dxr-styledialog .dxr-stylebar input[type="text"] { height: 28px; }
.dxr-styledialog .dxr-stylebar [data-sd="font"] { width: 11em; }
.dxr-styledialog .dxr-stylebar [data-sd="size"] { width: 4em; }
.dxr-styledialog .dxr-stylebar button {
  min-width: 28px; min-height: 28px; padding: 0 6px; border: 1px solid var(--dxr-rule-strong); border-radius: 6px;
  background: var(--dxr-sheet); color: var(--dxr-ink); font: inherit; cursor: pointer;
}
.dxr-styledialog .dxr-stylebar button[aria-pressed="true"] { border-color: var(--dxr-accent); background: var(--dxr-accent-wash, #ccfbf1); }
.dxr-styledialog .dxr-stylebar input[type="color"] { width: 32px; height: 28px; padding: 1px; }
.dxr-styledialog .dxr-stylebar[hidden] { display: none; }
.dxr-styledialog .dxr-dpreview { height: 96px; font-size: 6px; }
.dxr-styledialog .dxr-dpreview [data-sample] { font-size: 9px; }
.dxr-styledialog .dxr-styledesc { margin: 6px 0 0; color: var(--dxr-muted); font-size: 11.5px; }
.dxr-styledialog .dxr-dfoot [data-sd-format] { margin-right: auto; }
.dxr-styledialog .dxr-sdmenu {
  position: fixed; z-index: 3; display: flex; flex-direction: column; min-width: 140px; padding: 4px;
  border: 1px solid var(--dxr-rule); border-radius: 8px; background: var(--dxr-chrome); box-shadow: var(--dxr-shadow-xl);
}
.dxr-styledialog .dxr-sdmenu[hidden] { display: none; }
.dxr-styledialog .dxr-sdmenu button { min-height: 28px; padding: 2px 10px; border: 0; border-radius: 4px; background: none; color: var(--dxr-ink); font: inherit; text-align: left; cursor: pointer; }
.dxr-styledialog .dxr-sdmenu button:hover, .dxr-styledialog .dxr-sdmenu button:focus { background: var(--dxr-hover, rgba(15, 23, 42, .06)); }
`;

const OTHER = "Previous Paragraph ".repeat(12);
const NEXT = "Following Paragraph ".repeat(12);

export const STYLE_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-styledialog" data-dxr="styledialog">
  <form method="dialog" novalidate>
    <h2 data-sd-title>Modify Style</h2>
    <div class="dxr-stylebody">
      <fieldset>
        <legend>Properties</legend>
        <div class="dxr-dgrid">
          <label for="dxr-sd-name">Name</label><input id="dxr-sd-name" type="text" autocomplete="off" data-sd="name" />
          <label for="dxr-sd-type">Style type</label>
          <select id="dxr-sd-type" data-sd="type">
            <option value="paragraph">Paragraph</option>
            <option value="character">Character</option>
          </select>
          <label for="dxr-sd-basedon">Style based on</label><select id="dxr-sd-basedon" data-sd="basedOn"></select>
          <label for="dxr-sd-next" data-sd-nextlabel>Style for following paragraph</label><select id="dxr-sd-next" data-sd="next"></select>
        </div>
      </fieldset>
      <fieldset>
        <legend>Formatting</legend>
        <div class="dxr-stylebar">
          <input type="text" autocomplete="off" list="dxr-sd-fonts" data-sd="font" aria-label="Font" />
          <datalist id="dxr-sd-fonts" data-sd-fonts></datalist>
          <input type="text" autocomplete="off" list="dxr-fd-sizes" data-sd="size" aria-label="Font size" />
          <button type="button" data-sd-toggle="bold" aria-pressed="false" title="Bold"><b>B</b></button>
          <button type="button" data-sd-toggle="italic" aria-pressed="false" title="Italic"><i>I</i></button>
          <button type="button" data-sd-toggle="underline" aria-pressed="false" title="Underline"><u>U</u></button>
          <input type="color" data-sd="color" aria-label="Font color" title="Font color" />
        </div>
        <div class="dxr-stylebar" data-sd-parabar>
          <button type="button" data-sd-align="left" aria-pressed="false" title="Align left">&#8676;</button>
          <button type="button" data-sd-align="center" aria-pressed="false" title="Center">&#8596;</button>
          <button type="button" data-sd-align="right" aria-pressed="false" title="Align right">&#8677;</button>
          <button type="button" data-sd-align="justify" aria-pressed="false" title="Justify">&#8801;</button>
          <select data-sd="line" aria-label="Line spacing" title="Line spacing">
            <option value="" hidden></option>
            <option value="240">Single</option>
            <option value="276">1.15</option>
            <option value="360">1.5 lines</option>
            <option value="480">Double</option>
          </select>
          <button type="button" data-sd-space="up" title="Increase paragraph spacing">&#8645;+</button>
          <button type="button" data-sd-space="down" title="Decrease paragraph spacing">&#8645;&#8722;</button>
          <button type="button" data-sd-indent="-360" title="Decrease indent">&#8677;&#8722;</button>
          <button type="button" data-sd-indent="360" title="Increase indent">&#8677;+</button>
        </div>
      </fieldset>
      <fieldset>
        <legend>Preview</legend>
        <div class="dxr-dpreview" aria-hidden="true">
          <p>${OTHER}</p>
          <p data-sample></p>
          <p>${NEXT}</p>
        </div>
        <p class="dxr-styledesc" data-sd-desc></p>
      </fieldset>
      <label class="dxr-dcheck"><input type="checkbox" data-sd="gallery" /> Add to the Styles gallery</label>
    </div>
    <div class="dxr-derror" role="alert"><span data-sd-msg="name" hidden>Type a name for the style.</span><span data-sd-msg="size" hidden>Type a size from 1 to 1638 points.</span><span data-sd-msg="engine" hidden></span></div>
    <div class="dxr-dfoot">
      <button type="button" data-sd-format aria-haspopup="menu" aria-expanded="false">Format &#9662;</button>
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
    <div class="dxr-sdmenu" role="menu" data-sd-menu hidden>
      <button type="button" role="menuitem" data-sd-open="font">Font&#8230;</button>
      <button type="button" role="menuitem" data-sd-open="paragraph" data-sd-paraonly>Paragraph&#8230;</button>
    </div>
  </form>
</dialog>
`;

/** The caret's formatting, for a new style ("Create New Style from Formatting"). */
export interface StyleSource {
  /** The caret paragraph's style id. */
  styleId: string | null;
  paragraph: EditorParagraphFormat;
  run: RunFormattingInfo;
  /** A few words of the paragraph, for the preview. */
  text: string;
}

export interface StyleDialogHost {
  styles(): StyleInfo[];
  fonts(): string[];
  /** The caret's formatting (New Style starts from it), or null when there is no paragraph. */
  source(): StyleSource | null;
  /** Create or change the style; the engine's message when it refuses, else null. Called before
   *  the dialog closes, so a refusal keeps it open. */
  apply(op: StyleDefinitionOp, create: boolean): string | null;
  /** Open the Font / Paragraph dialog over this one with `host`. */
  openFont(host: FontDialogHost, sample: string): void;
  openParagraph(host: ParagraphDialogHost): void;
}

const RUN_KEYS = ["bold", "italic", "underline", "strike", "color", "fontSizePts", "fontFamily", "caps", "smallCaps", "vertAlign"] as const;
const PARA_KEYS = ["alignment", "leftIndent", "rightIndent", "firstLineIndent", "hangingIndent", "spacingBefore", "spacingAfter",
  "lineSpacing", "lineSpacingRule", "keepNext", "keepLines", "widowControl", "contextualSpacing", "pageBreakBefore"] as const;

/** An engine paragraph snapshot in the Paragraph dialog's names. */
export function paragraphOf(f: ParagraphFormatting | undefined): EditorParagraphFormat {
  if (!f) return {};
  const out: EditorParagraphFormat = {
    alignment: f.alignment, leftIndent: f.leftIndentTwips, rightIndent: f.rightIndentTwips,
    spacingBefore: f.spacingBeforeTwips, spacingAfter: f.spacingAfterTwips,
    lineSpacing: f.lineSpacing, lineSpacingRule: f.lineSpacingRule,
    keepNext: f.keepNext, keepLines: f.keepLines, widowControl: f.widowControl,
    contextualSpacing: f.contextualSpacing, pageBreakBefore: f.pageBreakBefore,
  };
  if (f.hangingIndentTwips) out.hangingIndent = f.hangingIndentTwips;
  else if (f.firstLineIndentTwips !== undefined) out.firstLineIndent = f.firstLineIndentTwips;
  return out;
}

/** The fields of `target` that `base` does not already give (a style's own formatting, made from
 *  what a paragraph shows over what the style it is based on gives). */
export function runDifference(target: RunFormattingInfo, base: RunFormattingInfo): FormatOp {
  const op: FormatOp = {};
  const norm = (key: string, v: unknown) =>
    key === "vertAlign" ? (v === "superscript" || v === "subscript" ? v : "baseline")
      : key === "color" ? (typeof v === "string" && /^[0-9a-f]{6}$/i.test(v) ? v.toUpperCase() : "")
        : typeof v === "boolean" || v === undefined ? !!v : v;
  for (const key of RUN_KEYS) {
    const t = norm(key, target[key]);
    if (t !== norm(key, base[key]) && !(key === "fontSizePts" && t === undefined)) {
      if (key === "color" && t === "") continue;
      (op as Record<string, unknown>)[key] = t;
    }
  }
  return op;
}

export function paragraphDifference(target: EditorParagraphFormat, base: EditorParagraphFormat): EditorParagraphFormat {
  const op: EditorParagraphFormat = {};
  for (const key of PARA_KEYS) {
    const t = target[key], b = base[key];
    const same = typeof t === "boolean" || typeof b === "boolean" ? !!t === !!b : (t ?? 0) === (b ?? 0);
    if (!same && t !== undefined) (op as Record<string, unknown>)[key] = t;
  }
  // Word's either/or indents: a first-line indent replaces a hanging one and back.
  if (op.hangingIndent !== undefined) delete op.firstLineIndent;
  if (op.lineSpacing !== undefined && op.lineSpacingRule === undefined) op.lineSpacingRule = target.lineSpacingRule ?? "auto";
  return op;
}

function colorInput(color: string | undefined): string {
  return color && /^[0-9a-f]{6}$/i.test(color) ? `#${color.toLowerCase()}` : "#000000";
}

/** The style dialog over its markup (`STYLE_DIALOG_HTML`, already in the page). */
export class StyleDialog {
  private create = false;
  private styleId: string | null = null;
  /** What the style gives now (Modify) or the caret shows (New), before this opening's changes. */
  private baseRun: RunFormattingInfo = {};
  private basePara: EditorParagraphFormat = {};
  /** This opening's changes, from the formatting row and the Font and Paragraph dialogs. */
  private run: FormatOp = {};
  private para: EditorParagraphFormat = {};
  private initialFields = new Map<string, string>();
  private sample = "";
  private builtIn = false;

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: StyleDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-sd-title]")!);
    const form = dialog.querySelector("form")!;
    form.addEventListener("input", (event) => {
      (event.target as HTMLElement).removeAttribute("aria-invalid");
      this.showError(null);
      this.formatRowChanged(event.target as HTMLElement);
      this.refresh();
    });
    this.field("type").addEventListener("change", () => { this.fillLists(); this.refresh(); });
    for (const b of Array.from(dialog.querySelectorAll<HTMLButtonElement>("[data-sd-toggle]"))) {
      b.addEventListener("click", () => {
        const key = b.dataset.sdToggle as "bold" | "italic" | "underline";
        this.run[key] = !this.effectiveRun()[key];
        this.refresh();
      });
    }
    for (const b of Array.from(dialog.querySelectorAll<HTMLButtonElement>("[data-sd-align]"))) {
      b.addEventListener("click", () => { this.para.alignment = b.dataset.sdAlign as EditorParagraphFormat["alignment"]; this.refresh(); });
    }
    for (const b of Array.from(dialog.querySelectorAll<HTMLButtonElement>("[data-sd-space]"))) {
      // Word's buttons: 6 points more or less before and after.
      b.addEventListener("click", () => {
        const step = b.dataset.sdSpace === "up" ? 120 : -120;
        const p = this.effectivePara();
        this.para.spacingBefore = Math.max(0, (p.spacingBefore ?? 0) + step);
        this.para.spacingAfter = Math.max(0, (p.spacingAfter ?? 0) + step);
        this.refresh();
      });
    }
    for (const b of Array.from(dialog.querySelectorAll<HTMLButtonElement>("[data-sd-indent]"))) {
      b.addEventListener("click", () => {
        this.para.leftIndent = Math.max(0, (this.effectivePara().leftIndent ?? 0) + Number(b.dataset.sdIndent));
        this.refresh();
      });
    }
    const formatButton = dialog.querySelector<HTMLButtonElement>("[data-sd-format]")!;
    const menu = dialog.querySelector<HTMLElement>("[data-sd-menu]")!;
    formatButton.addEventListener("click", () => {
      if (!menu.hidden) { this.closeMenu(); return; }
      menu.querySelector<HTMLElement>("[data-sd-paraonly]")!.hidden = this.type() !== "paragraph";
      menu.hidden = false;
      const r = formatButton.getBoundingClientRect();
      menu.style.left = `${r.left}px`;
      menu.style.top = `${Math.max(4, r.top - menu.offsetHeight - 4)}px`;
      formatButton.setAttribute("aria-expanded", "true");
      menu.querySelector<HTMLElement>("button:not([hidden])")?.focus();
    });
    menu.addEventListener("click", (event) => {
      const which = (event.target as HTMLElement).closest<HTMLElement>("[data-sd-open]")?.dataset.sdOpen;
      if (!which) return;
      this.closeMenu();
      if (which === "font") this.openFont(); else this.openParagraph();
    });
    menu.addEventListener("keydown", (event) => {
      if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); this.closeMenu(); formatButton.focus(); }
    });
    dialog.addEventListener("click", (event) => {
      if (!menu.hidden && !(event.target as HTMLElement).closest("[data-sd-menu], [data-sd-format]")) this.closeMenu();
    });
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.dialog.close("cancel"); return; }
      this.ok();
    });
    dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      if (!menu.hidden) { this.closeMenu(); return; }
      this.dialog.close("cancel");
    });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  /** Modify Style on `styleId`; false when there is no such paragraph or character style. */
  showModify(styleId: string): boolean {
    const style = this.host.styles().find((s) => s.id === styleId);
    if (!style || (style.type !== "paragraph" && style.type !== "character")) return false;
    this.create = false;
    this.styleId = style.id;
    this.builtIn = !style.isCustom;
    this.baseRun = { ...(style.resolvedRun ?? {}) };
    this.basePara = paragraphOf(style.resolvedParagraph);
    this.sample = this.host.source()?.text || style.name;
    return this.start("Modify Style", style.name, style.type, style.basedOn ?? "", style.next ?? "", !!style.quickFormat);
  }

  /** Create New Style from Formatting: starts from the caret's formatting and style. */
  showNew(): boolean {
    const source = this.host.source();
    if (!source) return false;
    const styles = this.host.styles();
    this.create = true;
    this.styleId = null;
    this.builtIn = false;
    this.baseRun = { ...source.run };
    this.basePara = { ...source.paragraph };
    this.sample = source.text;
    const taken = new Set(styles.map((s) => s.name.toLowerCase()));
    let n = 1;
    while (taken.has(`style${n}`)) n++;
    const basedOn = source.styleId ?? styles.find((s) => s.type === "paragraph" && s.isDefault)?.id ?? "";
    return this.start("Create New Style from Formatting", `Style${n}`, "paragraph", basedOn, "", true);
  }

  // ── internals ──

  private start(title: string, name: string, type: string, basedOn: string, next: string, gallery: boolean): boolean {
    this.run = {};
    this.para = {};
    this.dialog.querySelector<HTMLElement>("[data-sd-title]")!.textContent = title;
    const nameField = this.field("name");
    nameField.value = name;
    // Word keeps a built-in style's name.
    nameField.readOnly = this.builtIn;
    const typeField = this.select("type");
    typeField.value = type;
    typeField.disabled = !this.create;
    this.fillLists();
    this.select("basedOn").value = basedOn;
    this.select("next").value = this.create ? "" : next;
    this.field("gallery").checked = gallery;
    const fonts = this.dialog.querySelector<HTMLDataListElement>("[data-sd-fonts]")!;
    fonts.replaceChildren(...this.host.fonts().map((f) => Object.assign(document.createElement("option"), { value: f })));
    this.initialFields = this.snapshot();
    this.closeMenu();
    this.showError(null);
    this.refresh();
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    (this.builtIn ? this.select("basedOn") : nameField).focus();
    if (!this.builtIn) nameField.select();
    return true;
  }

  private type(): "paragraph" | "character" {
    return this.select("type").value === "character" ? "character" : "paragraph";
  }

  /** Based on: the styles of the dialog's type but this one; following: the paragraph styles. */
  private fillLists(): void {
    const styles = this.host.styles().filter((s) => s.id !== this.styleId);
    const byName = (a: StyleInfo, b: StyleInfo) => a.name.localeCompare(b.name);
    const option = (value: string, text: string) => Object.assign(document.createElement("option"), { value, textContent: text });
    const basedOn = this.select("basedOn");
    const keep = basedOn.value;
    basedOn.replaceChildren(option("", "(no style)"),
      ...styles.filter((s) => s.type === this.type()).sort(byName).map((s) => option(s.id, s.name)));
    basedOn.value = Array.from(basedOn.options).some((o) => o.value === keep) ? keep : "";
    const next = this.select("next");
    const keepNext = next.value;
    next.replaceChildren(option("", this.create ? "(this style)" : "(same style)"),
      ...this.host.styles().filter((s) => s.type === "paragraph" && s.id !== this.styleId).sort(byName).map((s) => option(s.id, s.name)));
    next.value = Array.from(next.options).some((o) => o.value === keepNext) ? keepNext : "";
    const paragraph = this.type() === "paragraph";
    next.hidden = !paragraph;
    this.dialog.querySelector<HTMLElement>("[data-sd-nextlabel]")!.hidden = !paragraph;
    this.dialog.querySelector<HTMLElement>("[data-sd-parabar]")!.hidden = !paragraph;
  }

  private effectiveRun(): RunFormattingInfo {
    const r: RunFormattingInfo = { ...this.baseRun };
    for (const [k, v] of Object.entries(this.run)) (r as Record<string, unknown>)[k] = v;
    return r;
  }

  private effectivePara(): EditorParagraphFormat {
    const p: EditorParagraphFormat = { ...this.basePara, ...this.para };
    if (this.para.hangingIndent !== undefined) delete p.firstLineIndent;
    if (this.para.firstLineIndent !== undefined) delete p.hangingIndent;
    return p;
  }

  /** The font, size and color fields write into this opening's changes as they are typed. */
  private formatRowChanged(target: HTMLElement): void {
    const key = target.dataset.sd;
    if (key === "font") {
      const name = this.field("font").value.trim();
      if (name) this.run.fontFamily = name; else delete this.run.fontFamily;
    } else if (key === "size") {
      const pts = Number(this.field("size").value.trim().replace(",", "."));
      if (pts >= 1 && pts <= 1638) this.run.fontSizePts = Math.round(pts * 2) / 2;
    } else if (key === "color") {
      this.run.color = this.field("color").value.slice(1).toUpperCase();
    } else if (key === "line") {
      const line = Number(this.select("line").value);
      if (line > 0) { this.para.lineSpacing = line; this.para.lineSpacingRule = "auto"; }
    }
  }

  /** The formatting row, the preview and the description, from the effective formatting. */
  private refresh(): void {
    const run = this.effectiveRun();
    const para = this.effectivePara();
    const active = this.dialog.ownerDocument.activeElement;
    if (active !== this.field("font")) this.field("font").value = run.fontFamily ?? "";
    if (active !== this.field("size")) this.field("size").value = run.fontSizePts !== undefined ? String(run.fontSizePts) : "";
    this.field("color").value = colorInput(run.color);
    for (const b of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-sd-toggle]"))) {
      b.setAttribute("aria-pressed", String(!!run[b.dataset.sdToggle as "bold"]));
    }
    for (const b of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-sd-align]"))) {
      b.setAttribute("aria-pressed", String((para.alignment ?? "left") === b.dataset.sdAlign));
    }
    const line = this.select("line");
    const lineValue = para.lineSpacingRule && para.lineSpacingRule !== "auto" ? "" : String(para.lineSpacing ?? 240);
    line.value = Array.from(line.options).some((o) => o.value === lineValue) ? lineValue : "";

    const sample = this.dialog.querySelector<HTMLElement>("[data-sample]")!;
    sample.textContent = (this.sample || "Sample Text").repeat(1) + " " + "Sample Text ".repeat(10);
    sample.style.fontFamily = run.fontFamily ? `"${run.fontFamily}"` : "";
    sample.style.fontWeight = run.bold ? "700" : "400";
    sample.style.fontStyle = run.italic ? "italic" : "normal";
    sample.style.textDecorationLine = [run.underline ? "underline" : "", run.strike ? "line-through" : ""].filter(Boolean).join(" ") || "none";
    sample.style.color = colorInput(run.color);
    sample.style.fontVariant = run.smallCaps ? "small-caps" : "normal";
    sample.style.textTransform = run.caps ? "uppercase" : "none";
    sample.style.fontSize = `${Math.max(5, Math.min(14, (run.fontSizePts ?? 11) * 0.75))}px`;
    const paragraph = this.type() === "paragraph";
    sample.style.textAlign = paragraph ? (para.alignment === "justify" ? "justify" : para.alignment ?? "left") : "";
    sample.style.marginLeft = paragraph ? `${(para.leftIndent ?? 0) / 40}px` : "";
    sample.style.marginTop = paragraph ? `${(para.spacingBefore ?? 0) / 40}px` : "";
    sample.style.marginBottom = paragraph ? `${(para.spacingAfter ?? 0) / 40}px` : "";
    sample.style.lineHeight = paragraph && (!para.lineSpacingRule || para.lineSpacingRule === "auto") ? String(1.2 * (para.lineSpacing ?? 240) / 240) : "";
    this.dialog.querySelector<HTMLElement>("[data-sd-desc]")!.textContent = this.describe(run, para);
  }

  /** Word's description under the preview: the style's formatting in words. */
  private describe(run: RunFormattingInfo, para: EditorParagraphFormat): string {
    const parts: string[] = [];
    const font = [run.fontFamily ? `Font: ${run.fontFamily}` : "", run.fontSizePts ? `${run.fontSizePts} pt` : "",
      run.bold ? "Bold" : "", run.italic ? "Italic" : "", run.underline ? "Underline" : "", run.strike ? "Strikethrough" : "",
      run.smallCaps ? "Small caps" : "", run.caps ? "All caps" : "",
      run.vertAlign === "superscript" ? "Superscript" : run.vertAlign === "subscript" ? "Subscript" : "",
      run.color && /^[0-9a-f]{6}$/i.test(run.color) ? `Font color: #${run.color.toUpperCase()}` : ""].filter(Boolean);
    if (font.length) parts.push(font.join(", "));
    if (this.type() === "paragraph") {
      const align = { left: "", center: "Centered", right: "Right", justify: "Justified" }[para.alignment ?? "left"];
      const lines = !para.lineSpacingRule || para.lineSpacingRule === "auto"
        ? ({ 240: "Single", 276: "Multiple 1.15 li", 360: "1.5 lines", 480: "Double" } as Record<number, string>)[para.lineSpacing ?? 240]
          ?? `Multiple ${((para.lineSpacing ?? 240) / 240).toFixed(2)} li`
        : `${para.lineSpacingRule === "exact" ? "Exactly" : "At least"} ${formatPoints(para.lineSpacing ?? 0)}`;
      const p = [align, para.leftIndent ? `Indent: Left ${formatPoints(para.leftIndent)}` : "", `Line spacing: ${lines}`,
        para.spacingBefore ? `Space Before: ${formatPoints(para.spacingBefore)}` : "",
        para.spacingAfter ? `Space After: ${formatPoints(para.spacingAfter)}` : "",
        para.keepNext ? "Keep with next" : "", para.keepLines ? "Keep lines together" : "",
        para.pageBreakBefore ? "Page break before" : ""].filter(Boolean);
      parts.push(p.join(", "));
    }
    const basedOn = this.select("basedOn").selectedOptions[0]?.textContent;
    if (this.select("basedOn").value && basedOn) parts.push(`Based on: ${basedOn}`);
    if (this.type() === "paragraph" && this.select("next").value) parts.push(`Following style: ${this.select("next").selectedOptions[0]?.textContent}`);
    if (this.field("gallery").checked) parts.push("Show in the Styles gallery");
    return parts.filter(Boolean).join("; ");
  }

  private openFont(): void {
    this.host.openFont({
      read: () => this.effectiveRun(),
      fonts: () => this.host.fonts(),
      apply: (op) => { Object.assign(this.run, op); this.refresh(); this.focusBack(); },
      cancelled: () => this.focusBack(),
    }, this.sample);
  }

  private openParagraph(): void {
    this.host.openParagraph({
      read: () => [this.effectivePara()],
      apply: (op) => {
        Object.assign(this.para, op);
        if (op.hangingIndent !== undefined) delete this.para.firstLineIndent;
        if (op.firstLineIndent !== undefined) delete this.para.hangingIndent;
        this.refresh();
        this.focusBack();
      },
      cancelled: () => this.focusBack(),
    });
  }

  private focusBack(): void {
    this.dialog.querySelector<HTMLElement>("[data-sd-format]")?.focus();
  }

  private closeMenu(): void {
    this.dialog.querySelector<HTMLElement>("[data-sd-menu]")!.hidden = true;
    this.dialog.querySelector<HTMLElement>("[data-sd-format]")!.setAttribute("aria-expanded", "false");
  }

  private ok(): void {
    const name = this.field("name").value.trim();
    if (!name) { this.invalid("name", "name"); return; }
    const size = this.field("size").value.trim();
    if (size && !(Number(size.replace(",", ".")) >= 1 && Number(size.replace(",", ".")) <= 1638)) { this.invalid("size", "size"); return; }
    const now = this.snapshot();
    const changed = (key: string) => now.get(key) !== this.initialFields.get(key);
    const type = this.type();
    const op: StyleDefinitionOp = {};
    if (this.create) {
      op.name = name;
      op.type = type;
      op.basedOn = this.select("basedOn").value;
      if (type === "paragraph" && this.select("next").value) op.next = this.select("next").value;
      op.quickFormat = this.field("gallery").checked;
      // The new style holds what the caret shows over what its based-on style gives, and the
      // changes made here.
      const basis = this.host.styles().find((s) => s.id === op.basedOn);
      const run = runDifference(this.effectiveRun(), basis?.resolvedRun ?? {});
      if (Object.keys(run).length) op.run = run;
      if (type === "paragraph") {
        const para = paragraphDifference(this.effectivePara(), paragraphOf(basis?.resolvedParagraph));
        if (Object.keys(para).length) op.paragraph = para;
      }
    } else {
      op.styleId = this.styleId!;
      if (changed("name") && !this.builtIn) op.name = name;
      if (changed("basedOn")) op.basedOn = this.select("basedOn").value;
      if (changed("next")) op.next = this.select("next").value;
      if (changed("gallery")) op.quickFormat = this.field("gallery").checked;
      if (Object.keys(this.run).length) op.run = { ...this.run };
      if (type === "paragraph" && Object.keys(this.para).length) op.paragraph = { ...this.para };
    }
    const nothing = !this.create && Object.keys(op).length === 1;
    if (nothing) { this.dialog.close("cancel"); return; }
    const refused = this.host.apply(op, this.create);
    if (refused !== null) {
      const msg = this.dialog.querySelector<HTMLElement>('[data-sd-msg="engine"]')!;
      this.showError("engine");
      msg.textContent = refused;
      return;
    }
    this.dialog.close("ok");
  }

  private invalid(field: string, message: string): void {
    const el = this.field(field);
    el.setAttribute("aria-invalid", "true");
    el.focus();
    this.showError(message);
  }

  private field(name: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`[data-sd="${name}"]`)!;
  }

  private select(name: string): HTMLSelectElement {
    return this.dialog.querySelector<HTMLSelectElement>(`[data-sd="${name}"]`)!;
  }

  private snapshot(): Map<string, string> {
    const out = new Map<string, string>();
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLInputElement | HTMLSelectElement>("[data-sd]"))) {
      out.set(el.dataset.sd!, el instanceof HTMLInputElement && el.type === "checkbox" ? String(el.checked) : el.value);
    }
    return out;
  }

  private showError(which: string | null): void {
    for (const msg of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-sd-msg]"))) msg.hidden = msg.dataset.sdMsg !== which;
  }
}
