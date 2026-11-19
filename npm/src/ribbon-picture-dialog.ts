/**
 * The ribbon's picture dialog — Word's Layout dialog for a picture (Size, Text Wrapping and
 * Position in one) with its Alt Text pane as a third tab. OK writes only what was changed:
 * the wrapping and position first, then the size, then the alt text.
 */
import type {
  FloatingImageLayout,
  ImageHorizontalAlignment,
  ImageHorizontalReference,
  ImageVerticalAlignment,
  ImageVerticalReference,
} from "./types.js";
import type { PictureInfo, PictureWrap } from "./editor.js";
import { formatLength, makeDialogMovable, parseLength, restoreDialogPlace, type MeasurementUnit } from "./ribbon-paragraph-dialog.js";

export const PICTURE_DIALOG_CSS = `
.dxr-picdialog { width: min(470px, calc(100vw - 32px)); }
.dxr-picdialog .dxr-dgrid input[type="text"] { width: 6em; }
.dxr-picdialog .dxr-pdpos { display: grid; grid-template-columns: auto auto auto auto; gap: 6px 8px; align-items: center; margin: 4px 0 10px; }
.dxr-picdialog .dxr-pdpos label { white-space: nowrap; }
.dxr-picdialog .dxr-pdpos input[type="text"] { width: 6em; }
.dxr-picdialog h3 { margin: 10px 0 4px; font-size: 12px; font-weight: 600; color: var(--dxr-ink); }
.dxr-picdialog textarea { width: 100%; box-sizing: border-box; min-height: 5.5em; resize: vertical; font: inherit; }
.dxr-picdialog input[data-pd2="title"] { width: 100%; box-sizing: border-box; }
.dxr-picdialog .dxr-pdorig { margin-top: 10px; display: flex; gap: 10px; align-items: center; color: var(--dxr-muted); }
.dxr-picdialog fieldset { border: 0; margin: 0; padding: 0; min-width: 0; }
.dxr-picdialog fieldset:disabled { opacity: .55; }
.dxr-picmenu { position: fixed; z-index: 2147483001; }
.dxr-picmenu .dxr-picmenu-label { padding: 6px 10px 2px; font-size: 11px; color: var(--dxr-muted); }
.dxr-picmenu button[aria-checked="true"] span::before { content: "\\2713  "; }
.dxr-picmenu hr { border: 0; border-top: 1px solid var(--dxr-rule); margin: 4px 0; }
`;

/** The wrap choices, in Word's order, with Word's names. */
export const PICTURE_WRAP_LABELS: ReadonlyArray<readonly [PictureWrap, string]> = [
  ["inline", "In Line with Text"],
  ["square", "Square"],
  ["tight", "Tight"],
  ["through", "Through"],
  ["topAndBottom", "Top and Bottom"],
  ["behind", "Behind Text"],
  ["inFront", "In Front of Text"],
];

const WRAP_OPTIONS = PICTURE_WRAP_LABELS.map(([v, label]) => `<option value="${v}">${label}</option>`).join("");

export const PICTURE_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-picdialog" data-dxr="picturedialog">
  <form method="dialog" novalidate>
    <h2 data-pd2-title>Picture</h2>
    <div role="tablist">
      <button type="button" role="tab" data-pd2-tab="size" aria-selected="true">Size</button>
      <button type="button" role="tab" data-pd2-tab="layout" aria-selected="false">Text Wrapping and Position</button>
      <button type="button" role="tab" data-pd2-tab="alt" aria-selected="false">Alt Text</button>
    </div>
    <div role="tabpanel" data-pd2-panel="size">
      <div class="dxr-dgrid">
        <label>Height</label><input type="text" autocomplete="off" data-pd2="height" />
        <label>Width</label><input type="text" autocomplete="off" data-pd2="width" />
      </div>
      <label class="dxr-dcheck"><input type="checkbox" data-pd2="lock" checked /> Lock aspect ratio</label>
      <div class="dxr-pdorig"><span>Original size:</span> <span data-pd2-original></span> <button type="button" data-pd2="reset">Reset</button></div>
    </div>
    <div role="tabpanel" data-pd2-panel="layout" hidden>
      <div class="dxr-dgrid">
        <label>Wrapping style</label><select data-pd2="wrap">${WRAP_OPTIONS}</select>
      </div>
      <fieldset data-pd2-float>
        <h3>Horizontal</h3>
        <div class="dxr-pdpos">
          <label><input type="radio" name="pd2h" value="align" data-pd2="hmode-align" /> Alignment</label>
          <select data-pd2="halign">
            <option value="left">Left</option><option value="center">Centered</option><option value="right">Right</option>
            <option value="inside">Inside</option><option value="outside">Outside</option>
          </select>
          <label>relative to</label>
          <select data-pd2="hfrom-align">
            <option value="column">Column</option><option value="margin">Margin</option><option value="page">Page</option><option value="character">Character</option>
          </select>
          <label><input type="radio" name="pd2h" value="offset" data-pd2="hmode-offset" /> Absolute position</label>
          <input type="text" autocomplete="off" data-pd2="hoffset" />
          <label>to the right of</label>
          <select data-pd2="hfrom-offset">
            <option value="column">Column</option><option value="margin">Margin</option><option value="page">Page</option><option value="character">Character</option>
          </select>
        </div>
        <h3>Vertical</h3>
        <div class="dxr-pdpos">
          <label><input type="radio" name="pd2v" value="align" data-pd2="vmode-align" /> Alignment</label>
          <select data-pd2="valign">
            <option value="top">Top</option><option value="center">Centered</option><option value="bottom">Bottom</option>
            <option value="inside">Inside</option><option value="outside">Outside</option>
          </select>
          <label>relative to</label>
          <select data-pd2="vfrom-align">
            <option value="margin">Margin</option><option value="page">Page</option><option value="line">Line</option>
          </select>
          <label><input type="radio" name="pd2v" value="offset" data-pd2="vmode-offset" /> Absolute position</label>
          <input type="text" autocomplete="off" data-pd2="voffset" />
          <label>below</label>
          <select data-pd2="vfrom-offset">
            <option value="paragraph">Paragraph</option><option value="margin">Margin</option><option value="page">Page</option><option value="line">Line</option>
          </select>
        </div>
        <h3>Distance from text</h3>
        <div class="dxr-pdpos">
          <label>Top</label><input type="text" autocomplete="off" data-pd2="dtop" />
          <label>Left</label><input type="text" autocomplete="off" data-pd2="dleft" />
          <label>Bottom</label><input type="text" autocomplete="off" data-pd2="dbottom" />
          <label>Right</label><input type="text" autocomplete="off" data-pd2="dright" />
        </div>
      </fieldset>
    </div>
    <div role="tabpanel" data-pd2-panel="alt" hidden>
      <h3>Description</h3>
      <textarea data-pd2="alt" placeholder="How would you describe this picture to someone who cannot see it?"></textarea>
      <h3>Title</h3>
      <input type="text" autocomplete="off" data-pd2="title" />
    </div>
    <div class="dxr-derror" role="alert"><span data-pd2-msg="length" hidden>Type a measurement, such as 2" or 5 cm.</span><span data-pd2-msg="size" hidden>Type a size greater than zero.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

/** The right-click menu on a picture. */
export const PICTURE_MENU_HTML = `
<div class="dxr-pop dxr-filemenu dxr-picmenu" data-dxr="picturemenu" role="menu" aria-label="Picture">
  <button type="button" role="menuitem" data-pm="size"><span>Size and Position&#8230;</span></button>
  <div class="dxr-picmenu-label">Wrap Text</div>
  ${PICTURE_WRAP_LABELS.map(([v, label]) => `<button type="button" role="menuitemradio" aria-checked="false" data-pm-wrap="${v}"><span>${label}</span></button>`).join("\n  ")}
  <hr />
  <button type="button" role="menuitem" data-pm="alt"><span>Alt Text&#8230;</span></button>
  <button type="button" role="menuitem" data-pm="replace"><span>Change Picture&#8230;</span></button>
  <button type="button" role="menuitem" data-pm="delete"><span>Delete Picture</span></button>
</div>
`;

/** What OK changes; each part only when the user changed it. */
export interface PictureDialogResult {
  wrap?: PictureWrap;
  layout?: FloatingImageLayout;
  size?: { widthPoints: number; heightPoints: number };
  altText?: { altText: string; title: string };
}

export interface PictureDialogHost {
  /** The selected picture, or null when none is. */
  read(): PictureInfo | null;
  apply(result: PictureDialogResult): void;
}

export type PictureDialogTab = "size" | "layout" | "alt";

const EMU_PER_TWIP = 635;

/** The picture dialog over its markup (`PICTURE_DIALOG_HTML`, already in the page). */
export class PictureDialog {
  private unit: MeasurementUnit;
  private info: PictureInfo | null = null;
  private initial = "";

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: PictureDialogHost, unit: MeasurementUnit = "in") {
    this.unit = unit;
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-pd2-title]")!);
    const form = dialog.querySelector("form")!;
    for (const tab of Array.from(dialog.querySelectorAll<HTMLElement>("[data-pd2-tab]")))
      tab.addEventListener("click", () => this.selectTab(tab.dataset.pd2Tab as PictureDialogTab));
    form.addEventListener("input", (event) => {
      const target = event.target as HTMLElement;
      target.removeAttribute("aria-invalid");
      this.showError(null);
      const name = target.dataset.pd2;
      if (name === "width" || name === "height") this.keepAspect(name);
    });
    form.addEventListener("change", (event) => {
      const name = (event.target as HTMLElement).dataset.pd2;
      if (name === "wrap") this.enableLayout();
      // Picking a reference or a value picks its way of placing too.
      if (name === "halign" || name === "hfrom-align") this.check("hmode-align");
      if (name === "hoffset" || name === "hfrom-offset") this.check("hmode-offset");
      if (name === "valign" || name === "vfrom-align") this.check("vmode-align");
      if (name === "voffset" || name === "vfrom-offset") this.check("vmode-offset");
    });
    this.field("reset").addEventListener("click", () => {
      const info = this.info;
      if (!info?.originalWidthPoints || !info.originalHeightPoints) return;
      this.input("width").value = formatLength(info.originalWidthPoints * 20, this.unit);
      this.input("height").value = formatLength(info.originalHeightPoints * 20, this.unit);
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

  /** Open on the selected picture at `tab`. False (nothing opened) with no picture selected. */
  show(tab: PictureDialogTab = "size"): boolean {
    const info = this.host.read();
    if (!info) return false;
    this.info = info;
    const len = (points: number | null): string => (points === null ? "" : formatLength(points * 20, this.unit));
    const emu = (n: number | null | undefined): string => formatLength((n ?? 0) / EMU_PER_TWIP, this.unit);
    this.input("width").value = len(info.widthPoints);
    this.input("height").value = len(info.heightPoints);
    this.input("lock").checked = true;
    const original = this.dialog.querySelector<HTMLElement>("[data-pd2-original]")!;
    original.textContent = info.originalWidthPoints && info.originalHeightPoints
      ? `${len(info.originalWidthPoints)} × ${len(info.originalHeightPoints)}` : "—";
    for (const name of ["width", "height", "lock"]) this.field<HTMLInputElement>(name).disabled = !info.can.size;
    this.field<HTMLButtonElement>("reset").disabled = !info.can.size || !info.originalWidthPoints;

    const wrap = this.field<HTMLSelectElement>("wrap");
    wrap.value = info.wrap ?? "";
    wrap.disabled = !info.can.wrap;
    const l = info.layout;
    const hAligned = !!l?.horizontalAlignment;
    this.check(hAligned ? "hmode-align" : "hmode-offset");
    this.select("halign").value = l?.horizontalAlignment && l.horizontalAlignment !== "unknown" ? l.horizontalAlignment : "center";
    const hFrom = l?.horizontalRelativeFrom && l.horizontalRelativeFrom !== "unknown" ? l.horizontalRelativeFrom : "column";
    this.select("hfrom-align").value = hFrom;
    this.select("hfrom-offset").value = hFrom;
    this.input("hoffset").value = emu(l?.horizontalOffsetEmu);
    const vAligned = !!l?.verticalAlignment;
    this.check(vAligned ? "vmode-align" : "vmode-offset");
    this.select("valign").value = l?.verticalAlignment && l.verticalAlignment !== "unknown" ? l.verticalAlignment : "top";
    const vFrom = l?.verticalRelativeFrom && l.verticalRelativeFrom !== "unknown" ? l.verticalRelativeFrom : "paragraph";
    this.select("vfrom-align").value = vFrom === "paragraph" ? "margin" : vFrom;
    this.select("vfrom-offset").value = vFrom;
    this.input("voffset").value = emu(l?.verticalOffsetEmu);
    this.input("dtop").value = emu(l?.distanceTopEmu);
    this.input("dbottom").value = emu(l?.distanceBottomEmu);
    this.input("dleft").value = emu(l?.distanceLeftEmu ?? 114300);
    this.input("dright").value = emu(l?.distanceRightEmu ?? 114300);
    this.enableLayout();

    (this.dialog.querySelector("textarea[data-pd2=alt]") as HTMLTextAreaElement).value = info.altText;
    this.input("title").value = info.title;
    for (const name of ["alt", "title"]) this.field<HTMLInputElement>(name).disabled = !info.can.altText;

    this.initial = this.snapshot();
    this.showError(null);
    this.selectTab(tab);
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    const first = tab === "alt" ? this.field("alt") : tab === "layout" ? wrap : this.input("height");
    first.focus();
    return true;
  }

  selectTab(tab: PictureDialogTab): void {
    for (const t of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-pd2-tab]")))
      t.setAttribute("aria-selected", String(t.dataset.pd2Tab === tab));
    for (const panel of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-pd2-panel]")))
      panel.hidden = panel.dataset.pd2Panel !== tab;
  }

  /** The changes the fields ask for, or null (the bad field marked) when one is not valid. */
  collect(): PictureDialogResult | null {
    const info = this.info;
    if (!info) return null;
    const now = JSON.parse(this.snapshot()) as Record<string, unknown>;
    const was = JSON.parse(this.initial) as Record<string, unknown>;
    const changed = (...keys: string[]): boolean => keys.some((k) => JSON.stringify(now[k]) !== JSON.stringify(was[k]));
    const result: PictureDialogResult = {};

    if (changed("width", "height")) {
      const w = this.length("width"), h = this.length("height");
      if (w === undefined || h === undefined) return null;
      if (w <= 0) return this.invalid(this.input("width"), "size");
      if (h <= 0) return this.invalid(this.input("height"), "size");
      result.size = { widthPoints: w / 20, heightPoints: h / 20 };
    }

    const wrap = this.select("wrap").value as PictureWrap;
    const layoutKeys = ["hmode", "halign", "hfromAlign", "hfromOffset", "hoffset", "vmode", "valign", "vfromAlign",
      "vfromOffset", "voffset", "dtop", "dbottom", "dleft", "dright"];
    if (wrap === "inline") {
      if (info.wrap !== "inline" && changed("wrap")) result.wrap = "inline";
    } else if (changed("wrap", ...layoutKeys)) {
      const lengths: Record<string, number> = {};
      for (const name of ["hoffset", "voffset", "dtop", "dbottom", "dleft", "dright"]) {
        const twips = this.length(name, name.startsWith("d"));
        if (twips === undefined) return null;
        lengths[name] = Math.round(twips * EMU_PER_TWIP);
      }
      const hAlign = this.input("hmode-align").checked;
      const vAlign = this.input("vmode-align").checked;
      const base: FloatingImageLayout = info.layout ? { ...info.layout } : {};
      const mode = ({ square: "square", tight: "tight", through: "through", topAndBottom: "top_and_bottom",
        behind: "none", inFront: "none" } as const)[wrap];
      result.layout = {
        ...base,
        wrapMode: mode,
        behindDocument: wrap === "behind",
        wrapSide: base.wrapSide && base.wrapSide !== "unknown" ? base.wrapSide : "both_sides",
        horizontalRelativeFrom: this.select(hAlign ? "hfrom-align" : "hfrom-offset").value as ImageHorizontalReference,
        horizontalAlignment: hAlign ? this.select("halign").value as ImageHorizontalAlignment : null,
        horizontalOffsetEmu: hAlign ? null : lengths.hoffset,
        verticalRelativeFrom: this.select(vAlign ? "vfrom-align" : "vfrom-offset").value as ImageVerticalReference,
        verticalAlignment: vAlign ? this.select("valign").value as ImageVerticalAlignment : null,
        verticalOffsetEmu: vAlign ? null : lengths.voffset,
        distanceTopEmu: lengths.dtop,
        distanceBottomEmu: lengths.dbottom,
        distanceLeftEmu: lengths.dleft,
        distanceRightEmu: lengths.dright,
      };
      if (!info.layout) {
        result.layout.relativeHeight ??= 251658240;
        result.layout.layoutInCell ??= true;
        result.layout.allowOverlap ??= true;
      }
    }

    if (changed("alt", "title")) {
      result.altText = {
        altText: (now.alt as string) ?? "",
        title: (now.title as string) ?? "",
      };
    }
    return result;
  }

  /** The fields as they stand, to compare with how they opened. */
  private snapshot(): string {
    const v = (name: string): string => this.field<HTMLInputElement>(name).value;
    return JSON.stringify({
      width: v("width"), height: v("height"), wrap: v("wrap"),
      hmode: this.input("hmode-align").checked, halign: v("halign"), hfromAlign: v("hfrom-align"),
      hfromOffset: v("hfrom-offset"), hoffset: v("hoffset"),
      vmode: this.input("vmode-align").checked, valign: v("valign"), vfromAlign: v("vfrom-align"),
      vfromOffset: v("vfrom-offset"), voffset: v("voffset"),
      dtop: v("dtop"), dbottom: v("dbottom"), dleft: v("dleft"), dright: v("dright"),
      alt: v("alt"), title: v("title"),
    });
  }

  /** With the aspect locked, typing one side changes the other to keep the shape. */
  private keepAspect(changed: "width" | "height"): void {
    const info = this.info;
    if (!this.input("lock").checked || !info?.widthPoints || !info.heightPoints) return;
    const typed = parseLength(this.input(changed).value, this.unit);
    if (typed === null || typed <= 0) return;
    const other = changed === "width" ? "height" : "width";
    const ratio = changed === "width" ? info.heightPoints / info.widthPoints : info.widthPoints / info.heightPoints;
    this.input(other).value = formatLength(typed * ratio, this.unit);
  }

  /** Position and distances apply to a floating picture only, and only where the document lets them change. */
  private enableLayout(): void {
    const fieldset = this.dialog.querySelector<HTMLFieldSetElement>("[data-pd2-float]")!;
    const wrap = this.select("wrap").value;
    const canPlace = !!this.info && (this.info.can.layout || (this.info.wrap === "inline" && this.info.can.wrap));
    fieldset.disabled = wrap === "inline" || !canPlace;
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

  private check(name: string): void {
    this.input(name).checked = true;
  }

  private invalid(field: HTMLElement, message: string): null {
    const panel = field.closest<HTMLElement>("[data-pd2-panel]");
    if (panel?.hidden) this.selectTab(panel.dataset.pd2Panel as PictureDialogTab);
    field.setAttribute("aria-invalid", "true");
    this.showError(message);
    field.focus();
    return null;
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-pd2-msg]")))
      span.hidden = span.dataset.pd2Msg !== message;
  }

  private field<T extends HTMLElement = HTMLElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-pd2="${name}"]`)!;
  }

  private input(name: string): HTMLInputElement {
    return this.field<HTMLInputElement>(name);
  }

  private select(name: string): HTMLSelectElement {
    return this.field<HTMLSelectElement>(name);
  }
}
