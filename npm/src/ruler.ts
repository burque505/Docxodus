/**
 * Word's horizontal ruler, above the document: the page's margins, the caret paragraph's indents
 * (first line, hanging, left, right) and tab stops, drawn to the zoom and dragged to change them.
 *
 * The ruler holds no document state. Each {@link Ruler.refresh} reads {@link RulerHost.state}
 * (the editor's `rulerState()`: where the paragraph sits on screen and its formatting) and draws
 * from it; a drag commits through the host as ONE edit when the mouse is released (Escape
 * cancels). Lengths are twips; positions on screen are client pixels.
 *
 * - First-line marker (top ▽): the first line's indent. Hanging marker (bottom △): the other lines'
 *   indent; the first line stays where it is. Left box: both together. Right marker: right indent.
 * - Click the ruler to set a tab stop of the kind the selector at the left shows (click the
 *   selector to change it); drag a stop to move it, drag it off the ruler to remove it,
 *   double-click one to open the Tabs dialog.
 * - With the paragraph in the section's own text column, drag the edge between grey and white to
 *   change the section's left or right margin.
 * - Drags snap to 1/16 inch (0.25 cm); hold Alt to place freely.
 */

import type { EditorParagraphFormat, EditorRulerState } from "./editor.js";
import type { TabStop, TabStopAlignment } from "./types.js";

/** Inches or centimetres, as the Paragraph dialog shows lengths. */
export type RulerUnit = "in" | "cm";

/** The ruler's tooltips and selector names, English; a host translates them through `t`. */
export const RULER_TEXTS = {
  left: "Left Tab",
  center: "Center Tab",
  right: "Right Tab",
  decimal: "Decimal Tab",
  bar: "Bar Tab",
  firstLine: "First Line Indent",
  hanging: "Hanging Indent",
  leftIndent: "Left Indent",
  rightIndent: "Right Indent",
  leftMargin: "Left Margin",
  rightMargin: "Right Margin",
  column: "Move Table Column",
  selector: "Click to choose the kind of tab stop the ruler sets",
} as const;

export interface RulerHost {
  /** The caret paragraph's geometry and formatting, or null when there is none. */
  state(): EditorRulerState | null;
  unit(): RulerUnit;
  /** Translate one of {@link RULER_TEXTS}. */
  t(text: string): string;
  setIndents(op: EditorParagraphFormat): void;
  setTabStops(stops: TabStop[]): void;
  setMargins(op: { marginLeftTwips?: number; marginRightTwips?: number }): void;
  /** The caret's table's grid column widths, twips (a column marker's drag). */
  setColumnWidths(widthsTwips: number[]): void;
  openTabsDialog(): void;
  /** The document's visible area, which the dashed guide line spans while dragging. */
  guideArea(): HTMLElement | null;
}

export const RULER_CSS = `
.dxr-ruler { position: relative; height: 26px; flex: none; overflow: hidden; user-select: none;
  background: var(--dxr-desk); border-bottom: 1px solid var(--dxr-rule); touch-action: none;
  --dxr-ruler-paper: var(--dxr-sheet); --dxr-ruler-margin: var(--dxr-rule-strong);
  --dxr-ruler-ink: var(--dxr-muted); --dxr-ruler-marker: var(--dxr-ink); --dxr-ruler-hot: var(--dxr-accent); }
.dxr-ruler[hidden] { display: none; }
.dxr-ruler canvas { position: absolute; inset: 0; width: 100%; height: 100%; pointer-events: none; }
/* Each marker is a 13px-wide grab area around its symbol, in a band of its own from top to bottom
   (first line, table column, hanging indent, left box), so neighbours are hard to mistake; the
   one under the mouse takes the accent colour before it is pressed. */
.dxr-ruler-mark { position: absolute; width: 13px; margin-left: -6.5px; cursor: ew-resize;
  color: var(--dxr-ruler-marker); display: flex; justify-content: center; }
.dxr-ruler-mark svg { display: block; width: 9px; overflow: visible; }
.dxr-ruler-mark svg [fill]:not([fill="currentColor"]) { fill: var(--dxr-ruler-paper); }
.dxr-ruler-mark:hover, .dxr-ruler-mark[data-dragging] { color: var(--dxr-ruler-hot); }
.dxr-ruler-mark:hover svg [fill]:not([fill="currentColor"]),
.dxr-ruler-mark[data-dragging] svg [fill]:not([fill="currentColor"]) { fill: var(--dxr-wash-on, #ccfbf1); }
.dxr-ruler-mark[data-mark="firstLine"] { top: 0; height: 9px; align-items: flex-start; }
.dxr-ruler-mark[data-mark="hanging"], .dxr-ruler-mark[data-mark="rightIndent"] { top: 14px; height: 6px; align-items: flex-end; }
.dxr-ruler-mark[data-mark="leftIndent"] { top: 20px; height: 6px; align-items: flex-start; }
.dxr-ruler-mark[data-mark="tab"] { top: 12px; height: 11px; cursor: default; align-items: center; }
.dxr-ruler-mark[data-mark="leftMargin"], .dxr-ruler-mark[data-mark="rightMargin"] { top: 9px; height: 5px; cursor: col-resize; }
.dxr-ruler-mark[data-mark="column"] { top: 6px; height: 14px; cursor: col-resize; }
.dxr-ruler-mark[data-mark="column"]::before { content: ""; width: 7px; height: 100%; box-sizing: border-box;
  background: repeating-linear-gradient(0deg, currentColor 0 1px, var(--dxr-ruler-margin) 1px 3px);
  border: 1px solid currentColor; border-radius: 1px; }
.dxr-ruler-mark[data-mark="column"] { color: var(--dxr-ruler-ink); }
.dxr-ruler-mark[data-mark="column"]:hover, .dxr-ruler-mark[data-mark="column"][data-dragging] { color: var(--dxr-ruler-hot); }
.dxr-ruler-mark[data-removing] { opacity: .3; }
.dxr-ruler-selector { position: absolute; left: 2px; top: 4px; width: 18px; height: 18px; padding: 0;
  border: 1px solid var(--dxr-rule-strong); border-radius: 3px; background: var(--dxr-chrome);
  color: var(--dxr-ruler-marker); cursor: pointer; z-index: 1; }
.dxr-ruler-selector svg { display: block; margin: auto; width: 9px; height: 9px; overflow: visible; }
.dxr-ruler-guide { position: fixed; width: 0; border-left: 1px dashed var(--dxr-ink, #1e293b);
  pointer-events: none; z-index: 2147483000; }
.dxr-ruler-tip { position: fixed; padding: 1px 5px; font: 11px var(--dxr-ui, system-ui); border-radius: 3px;
  background: var(--dxr-ink, #1e293b); color: var(--dxr-sheet, #fff); pointer-events: none; z-index: 2147483001; }
/* The focus outline of a paragraph whose first line hangs left of its box (drawn by the ribbon
   around the box and that line, where a CSS outline would leave the line outside). */
.dxr-focusring { position: fixed; pointer-events: none; z-index: 1; border: 2px solid var(--dxr-accent);
  border-radius: 2px; box-sizing: border-box; }
.dxr-focusring[hidden] { display: none; }
`;

const TAB_KINDS: TabStopAlignment[] = ["left", "center", "right", "decimal", "bar"];

/** The glyph of a tab stop kind, in a 9×9 box: Word's L, ⊥, ⅃, ⊥ with a dot, and a bar. */
function tabGlyph(kind: TabStopAlignment): string {
  const path = {
    left: "M1.5 0 V7.5 H8",
    center: "M4.5 0 V7.5 M0.5 7.5 H8.5",
    right: "M7.5 0 V7.5 H1",
    decimal: "M4.5 0 V7.5 M0.5 7.5 H8.5",
    bar: "M4.5 0 V9",
  }[kind];
  const dot = kind === "decimal" ? `<rect x="6.5" y="3" width="1.6" height="1.6" fill="currentColor"/>` : "";
  return `<svg viewBox="0 0 9 9" aria-hidden="true"><path d="${path}" fill="none" stroke="currentColor" stroke-width="1.6"/>${dot}</svg>`;
}

const MARKER_SVG: Record<string, string> = {
  firstLine: `<svg viewBox="0 0 9 7" aria-hidden="true"><path d="M0.5 0.5 H8.5 L4.5 6.5 Z" fill="var(--dxr-ruler-paper)" stroke="currentColor"/></svg>`,
  hanging: `<svg viewBox="0 0 9 5" aria-hidden="true"><path d="M4.5 0.5 L8.5 4.5 H0.5 Z" fill="var(--dxr-ruler-paper)" stroke="currentColor"/></svg>`,
  leftIndent: `<svg viewBox="0 0 9 4" aria-hidden="true"><rect x="0.5" y="0" width="8" height="3.5" fill="var(--dxr-ruler-paper)" stroke="currentColor"/></svg>`,
  rightIndent: `<svg viewBox="0 0 9 5" aria-hidden="true"><path d="M4.5 0.5 L8.5 4.5 H0.5 Z" fill="var(--dxr-ruler-paper)" stroke="currentColor"/></svg>`,
  leftMargin: "",
  rightMargin: "",
};

/** The narrowest text an indent drag leaves a paragraph: 1/8 inch. */
const MIN_TEXT_TWIPS = 180;

/** Twips per snap step: 1/16 inch, or 0.25 cm. */
const SNAP: Record<RulerUnit, number> = { in: 90, cm: 1440 / 2.54 / 4 };

type MarkKind = "firstLine" | "hanging" | "leftIndent" | "rightIndent" | "tab" | "leftMargin" | "rightMargin" | "column";

interface Drag {
  kind: MarkKind;
  element: HTMLElement | null;
  /** For a tab: its index in the state's stops, or -1 for a new one from a click. For a column
   *  marker: its index in the state's table boundaries. */
  tab: number;
  startX: number;
  startY: number;
  moved: boolean;
  removing: boolean;
  /** The value (twips) the drag would commit. */
  twips: number;
}

export class Ruler {
  readonly element: HTMLElement;
  private readonly host: RulerHost;
  private readonly canvas: HTMLCanvasElement;
  private readonly marks: HTMLElement;
  private readonly selector: HTMLButtonElement;
  private kind: TabStopAlignment = "left";
  private current: EditorRulerState | null = null;
  private drag: Drag | null = null;
  private guide: HTMLElement | null = null;
  private tip: HTMLElement | null = null;
  /** The last click on a tab stop (its position and when), to tell a double-click: the marks are
   *  drawn again after every click, so the browser's dblclick lands on the ruler, not the stop. */
  private lastTabClick: { twips: number; at: number } | null = null;
  /** What the ruler last drew from (see refresh). */
  private drawn = "";

  constructor(element: HTMLElement, host: RulerHost) {
    this.element = element;
    this.host = host;
    element.classList.add("dxr-ruler");
    element.replaceChildren();
    const doc = element.ownerDocument;
    this.canvas = doc.createElement("canvas");
    this.marks = doc.createElement("div");
    this.selector = doc.createElement("button");
    this.selector.type = "button";
    this.selector.className = "dxr-ruler-selector";
    this.selector.dataset.dxr = "rulertab";
    this.selector.addEventListener("click", () => {
      this.kind = TAB_KINDS[(TAB_KINDS.indexOf(this.kind) + 1) % TAB_KINDS.length];
      this.drawSelector();
    });
    element.append(this.canvas, this.marks, this.selector);
    this.drawSelector();
    element.addEventListener("pointerdown", this.onPointerDown);
  }

  /** The kind of tab stop a click sets. */
  get tabKind(): TabStopAlignment {
    return this.kind;
  }

  set tabKind(kind: TabStopAlignment) {
    this.kind = kind;
    this.drawSelector();
  }

  /** Write the tooltips again (after the host's language changed). */
  relabel(): void {
    this.drawSelector();
    this.drawn = "";
    this.refresh();
  }

  /** Read the caret paragraph's state again and redraw (cheap; call it per frame at most). */
  refresh(): void {
    if (this.drag || this.element.hidden) return;
    this.current = this.host.state();
    // Unchanged (the same paragraph, layout, zoom and ruler width): nothing to draw again.
    const key = `${JSON.stringify(this.current)}|${this.host.unit()}|${this.element.getBoundingClientRect().width}`;
    if (key === this.drawn) return;
    this.drawn = key;
    this.draw();
  }

  destroy(): void {
    this.endDrag(false);
    this.element.removeEventListener("pointerdown", this.onPointerDown);
    this.element.replaceChildren();
  }

  // ── drawing ───────────────────────────────────────────────────────────────

  private drawSelector(): void {
    this.selector.innerHTML = tabGlyph(this.kind);
    this.selector.title = `${this.host.t(RULER_TEXTS[this.kind])}. ${this.host.t(RULER_TEXTS.selector)}`;
    this.selector.setAttribute("aria-label", this.host.t(RULER_TEXTS[this.kind]));
    this.selector.dataset.kind = this.kind;
  }

  private x(clientX: number): number {
    return clientX - this.element.getBoundingClientRect().left;
  }

  private draw(): void {
    const s = this.current;
    const rect = this.element.getBoundingClientRect();
    const ratio = (typeof devicePixelRatio === "number" && devicePixelRatio) || 1;
    const width = Math.max(1, Math.round(rect.width));
    const height = Math.max(1, Math.round(rect.height));
    if (this.canvas.width !== Math.round(width * ratio)) this.canvas.width = Math.round(width * ratio);
    if (this.canvas.height !== Math.round(height * ratio)) this.canvas.height = Math.round(height * ratio);
    const ctx = typeof this.canvas.getContext === "function" ? this.canvas.getContext("2d") : null;
    this.marks.replaceChildren();
    if (!ctx) return;
    ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
    ctx.clearRect(0, 0, width, height);
    if (!s) return;

    const css = getComputedStyle(this.element);
    const color = (name: string, fallback: string) => css.getPropertyValue(name).trim() || fallback;
    const left = s.pageLeft - rect.left;
    const right = s.pageRight - rect.left;
    const originLeft = s.originLeft - rect.left;
    const originRight = s.originRight - rect.left;
    const top = 4, bandHeight = height - 8;
    ctx.fillStyle = color("--dxr-ruler-margin", "#cbd5e1");
    ctx.fillRect(left, top, right - left, bandHeight);
    ctx.fillStyle = color("--dxr-ruler-paper", "#fff");
    ctx.fillRect(originLeft, top, originRight - originLeft, bandHeight);

    // Ticks from the paragraph's origin outwards: eighths of an inch (quarter centimetres),
    // a taller one at each half, a number at each whole unit.
    const unit = this.host.unit();
    const unitTwips = unit === "cm" ? 1440 / 2.54 : 1440;
    const parts = unit === "cm" ? 4 : 8;
    const step = unitTwips / parts * s.pxPerTwip;
    ctx.strokeStyle = ctx.fillStyle = color("--dxr-ruler-ink", "#475569");
    ctx.lineWidth = 1;
    ctx.font = `9px ${css.getPropertyValue("--dxr-ui").trim() || "system-ui"}`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    const mid = top + bandHeight / 2;
    if (step > 1.5) {
      for (const direction of [1, -1]) {
        for (let i = direction > 0 ? 1 : 1; ; i++) {
          const x = originLeft + direction * i * step;
          if (x < left - 0.5 || x > right + 0.5) break;
          const k = i % parts;
          if (k === 0) {
            const label = String(i / parts);
            if (step * parts > 14) ctx.fillText(label, x, mid + 0.5);
          } else {
            const h = k === parts / 2 ? 3 : 1.5;
            const px = Math.round(x) + 0.5;
            ctx.beginPath();
            ctx.moveTo(px, mid - h);
            ctx.lineTo(px, mid + h);
            ctx.stroke();
          }
        }
      }
    }

    // Default tab stops after the last set one: faint marks along the bottom edge.
    const lastStop = s.tabs.reduce((max, t) => Math.max(max, t.positionTwips), 0);
    if (s.defaultTabStopTwips > 0) {
      ctx.strokeStyle = color("--dxr-ruler-ink", "#475569");
      ctx.globalAlpha = 0.5;
      for (let twips = s.defaultTabStopTwips; ; twips += s.defaultTabStopTwips) {
        const x = originLeft + twips * s.pxPerTwip;
        if (x > originRight) break;
        if (twips <= lastStop) continue;
        const px = Math.round(x) + 0.5;
        ctx.beginPath();
        ctx.moveTo(px, top + bandHeight);
        ctx.lineTo(px, top + bandHeight + 2);
        ctx.stroke();
      }
      ctx.globalAlpha = 1;
    }

    // The margin edges first, so the indent markers drawn after them are on top: as in Word, a
    // margin is dragged by the band between the top and bottom markers.
    if (s.inSectionColumn) {
      this.addMark("leftMargin", s.textLeft - rect.left);
      this.addMark("rightMargin", s.textRight - rect.left);
    }
    // A table's column boundaries (Word's Move Table Column markers), under the indent markers.
    s.table?.boundaries.forEach((b, index) => {
      this.addMark("column", b.x - rect.left).dataset.column = String(index);
    });
    const at = (twips: number) => s.originLeft + twips * s.pxPerTwip - rect.left;
    const indentX = at(s.leftIndent);
    this.addMark("firstLine", at(s.leftIndent + s.firstLine));
    this.addMark("hanging", indentX);
    this.addMark("leftIndent", indentX);
    this.addMark("rightIndent", originRight - s.rightIndent * s.pxPerTwip);
    s.tabs.forEach((tab, index) => {
      const mark = this.addMark("tab", at(tab.positionTwips), tabGlyph(tab.alignment), RULER_TEXTS[tab.alignment]);
      mark.dataset.tab = String(index);
      mark.dataset.kind = tab.alignment;
    });
  }

  private addMark(kind: MarkKind, x: number, svg?: string, title?: string): HTMLElement {
    const mark = this.element.ownerDocument.createElement("div");
    mark.className = "dxr-ruler-mark";
    mark.dataset.mark = kind;
    mark.style.left = `${x}px`;
    mark.innerHTML = svg ?? MARKER_SVG[kind] ?? "";
    mark.title = this.host.t(title ?? RULER_TEXTS[kind as keyof typeof RULER_TEXTS]);
    this.marks.appendChild(mark);
    return mark;
  }

  // ── dragging ──────────────────────────────────────────────────────────────

  /** Twips from the paragraph's origin at client x, snapped unless `free`. */
  private twipsAt(clientX: number, free: boolean, from: number): number {
    const s = this.current!;
    const raw = (clientX - from) / s.pxPerTwip;
    if (free) return Math.round(raw);
    const step = SNAP[this.host.unit()];
    return Math.round(Math.round(raw / step) * step);
  }

  private readonly onPointerDown = (ev: PointerEvent): void => {
    if (ev.button !== 0 || !this.current || ev.target === this.selector || this.selector.contains(ev.target as Node)) return;
    const mark = (ev.target as Element).closest<HTMLElement>(".dxr-ruler-mark");
    const s = this.current;
    let kind: MarkKind;
    let tab = -1;
    if (mark) {
      kind = mark.dataset.mark as MarkKind;
      if (kind === "column") tab = Number(mark.dataset.column);
      if (kind === "tab") {
        tab = Number(mark.dataset.tab);
        const twips = s.tabs[tab]?.positionTwips;
        const now = Date.now();
        if (this.lastTabClick && this.lastTabClick.twips === twips && now - this.lastTabClick.at < 500) {
          this.lastTabClick = null;
          ev.preventDefault();
          this.host.openTabsDialog();
          return;
        }
        this.lastTabClick = twips === undefined ? null : { twips, at: now };
      }
    } else {
      // A click on the white part of the ruler sets a tab stop there.
      if (ev.clientX < s.originLeft || ev.clientX > s.originRight) return;
      kind = "tab";
    }
    ev.preventDefault();
    this.element.setPointerCapture?.(ev.pointerId);
    mark?.setAttribute("data-dragging", "");
    this.drag = { kind, element: mark, tab, startX: ev.clientX, startY: ev.clientY, moved: false, removing: false, twips: 0 };
    this.element.addEventListener("pointermove", this.onPointerMove);
    this.element.addEventListener("pointerup", this.onPointerUp);
    this.element.addEventListener("pointercancel", this.onPointerCancel);
    this.element.ownerDocument.addEventListener("keydown", this.onKeyDown, true);
    if (!mark) {
      // A new stop: placed at once where clicked, moved with the mouse until released.
      this.drag.moved = true;
      this.drag.element = this.addMark("tab", ev.clientX - this.element.getBoundingClientRect().left,
        tabGlyph(this.kind), RULER_TEXTS[this.kind]);
    }
    this.move(ev);
  };

  private readonly onPointerMove = (ev: PointerEvent): void => {
    if (!this.drag) return;
    if (Math.abs(ev.clientX - this.drag.startX) > 2 || Math.abs(ev.clientY - this.drag.startY) > 2) this.drag.moved = true;
    this.move(ev);
  };

  private move(ev: PointerEvent): void {
    const drag = this.drag!;
    const s = this.current!;
    const free = ev.altKey;
    let clientX: number;
    switch (drag.kind) {
      case "rightIndent": {
        // The right indent stops short of the left and first-line indents (Word keeps a sliver of
        // text width: a paragraph with none can no longer be clicked into), and at its reach.
        const width = (s.originRight - s.originLeft) / s.pxPerTwip;
        const maxTwips = Math.floor(width - Math.max(s.leftIndent, s.leftIndent + s.firstLine) - MIN_TEXT_TWIPS);
        const minTwips = Math.round((s.originRight - s.boundsRight) / s.pxPerTwip);
        drag.twips = Math.max(minTwips, Math.min(maxTwips, -this.twipsAt(ev.clientX, free, s.originRight)));
        clientX = s.originRight - drag.twips * s.pxPerTwip;
        break;
      }
      case "leftMargin": {
        const twips = this.twipsAt(ev.clientX, free, s.pageLeft);
        const max = Math.round((s.textRight - s.pageLeft) / s.pxPerTwip) - 720;
        drag.twips = Math.max(0, Math.min(max, twips));
        clientX = s.pageLeft + drag.twips * s.pxPerTwip;
        break;
      }
      case "rightMargin": {
        const twips = -this.twipsAt(ev.clientX, free, s.pageRight);
        const max = Math.round((s.pageRight - s.textLeft) / s.pxPerTwip) - 720;
        drag.twips = Math.max(0, Math.min(max, twips));
        clientX = s.pageRight - drag.twips * s.pxPerTwip;
        break;
      }
      case "column": {
        // Twips the boundary moves; the column before it keeps at least a quarter inch, and so
        // does the one after it (which gives up what the one before takes).
        const b = s.table!.boundaries[drag.tab];
        const before = s.table!.widthsTwips[b.grid - 1] ?? 0;
        const after = b.grid < s.table!.widthsTwips.length ? s.table!.widthsTwips[b.grid] : Infinity;
        const raw = free ? Math.round((ev.clientX - b.x) / s.pxPerTwip) : this.twipsAt(ev.clientX, false, b.x);
        drag.twips = Math.max(360 - before, Math.min(after - 360, raw,
          Math.round((s.pageRight - b.x) / s.pxPerTwip)));
        clientX = b.x + drag.twips * s.pxPerTwip;
        break;
      }
      default: {
        // An indent stays within its reach (a table cell's edges, else the page's) and, as the
        // right indent, short of the other side: the first line and the other lines both keep
        // some text width. (A tab stop only keeps within the reach.)
        let minTwips = Math.round((s.boundsLeft - s.originLeft) / s.pxPerTwip);
        let maxTwips = Math.round((s.boundsRight - s.originLeft) / s.pxPerTwip);
        // The box carries the first line: with a hanging indent that sits left of the box, and it
        // must stay within the reach too (else it lands in the cell to the left).
        if (drag.kind === "leftIndent" && s.firstLine < 0) minTwips -= s.firstLine;
        const rightEdge = Math.floor((s.originRight - s.originLeft) / s.pxPerTwip - s.rightIndent - MIN_TEXT_TWIPS);
        if (drag.kind === "firstLine" || drag.kind === "hanging") maxTwips = Math.min(maxTwips, rightEdge);
        // (the box carries the first line with it)
        if (drag.kind === "leftIndent") maxTwips = Math.min(maxTwips, rightEdge, rightEdge - s.firstLine);
        drag.twips = Math.max(minTwips, Math.min(maxTwips, this.twipsAt(ev.clientX, free, s.originLeft)));
        clientX = s.originLeft + drag.twips * s.pxPerTwip;
      }
    }
    const rect = this.element.getBoundingClientRect();
    if (drag.kind === "tab") {
      drag.removing = ev.clientY > rect.bottom + 16 || ev.clientY < rect.top - 16;
      if (drag.element) drag.element.toggleAttribute("data-removing", drag.removing);
    }
    if (drag.element) drag.element.style.left = `${clientX - rect.left}px`;
    // The hanging marker and the left box move together, as do all three for the left box.
    if (drag.kind === "leftIndent" || drag.kind === "hanging") {
      const box = this.marks.querySelector<HTMLElement>(`[data-mark="${drag.kind === "leftIndent" ? "hanging" : "leftIndent"}"]`);
      if (box) box.style.left = `${clientX - rect.left}px`;
      if (drag.kind === "leftIndent") {
        const first = this.marks.querySelector<HTMLElement>('[data-mark="firstLine"]');
        if (first) first.style.left = `${clientX + s.firstLine * s.pxPerTwip - rect.left}px`;
      }
    }
    if (drag.moved) this.showGuide(clientX, drag.removing ? "" : this.describe(drag));
  }

  private describe(drag: Drag): string {
    const unit = this.host.unit();
    const value = unit === "cm" ? `${(drag.twips / (1440 / 2.54)).toFixed(2)} cm` : `${(drag.twips / 1440).toFixed(2)}"`;
    if (drag.kind === "column") {
      const s = this.current!;
      const b = s.table!.boundaries[drag.tab];
      const width = (s.table!.widthsTwips[b.grid - 1] ?? 0) + drag.twips;
      const shown = unit === "cm" ? `${(width / (1440 / 2.54)).toFixed(2)} cm` : `${(width / 1440).toFixed(2)}"`;
      return `${this.host.t(RULER_TEXTS.column)}: ${shown}`;
    }
    return `${this.host.t(drag.kind === "tab" ? RULER_TEXTS[this.tabKindOf(drag)] : RULER_TEXTS[drag.kind])}: ${value}`;
  }

  private tabKindOf(drag: Drag): TabStopAlignment {
    return drag.tab >= 0 ? this.current!.tabs[drag.tab]?.alignment ?? this.kind : this.kind;
  }

  private showGuide(clientX: number, text: string): void {
    const doc = this.element.ownerDocument;
    const area = this.host.guideArea()?.getBoundingClientRect();
    const ruler = this.element.getBoundingClientRect();
    if (!this.guide) {
      this.guide = doc.createElement("div");
      this.guide.className = "dxr-ruler-guide";
      this.guide.dataset.dxr = "rulerguide";
      doc.body.appendChild(this.guide);
    }
    const top = ruler.bottom;
    const bottom = area ? area.bottom : ruler.bottom + 400;
    this.guide.style.left = `${Math.round(clientX)}px`;
    this.guide.style.top = `${top}px`;
    this.guide.style.height = `${Math.max(0, bottom - top)}px`;
    if (!this.tip) {
      this.tip = doc.createElement("div");
      this.tip.className = "dxr-ruler-tip";
      doc.body.appendChild(this.tip);
    }
    this.tip.hidden = !text;
    this.tip.textContent = text;
    this.tip.style.left = `${Math.round(clientX) + 6}px`;
    this.tip.style.top = `${top + 4}px`;
  }

  private readonly onPointerUp = (ev: PointerEvent): void => {
    this.move(ev);
    this.endDrag(true);
  };

  private readonly onPointerCancel = (): void => this.endDrag(false);

  private readonly onKeyDown = (ev: KeyboardEvent): void => {
    if (ev.key !== "Escape" || !this.drag) return;
    ev.preventDefault();
    ev.stopPropagation();
    this.endDrag(false);
  };

  private endDrag(commit: boolean): void {
    const drag = this.drag;
    this.drag = null;
    this.element.removeEventListener("pointermove", this.onPointerMove);
    this.element.removeEventListener("pointerup", this.onPointerUp);
    this.element.removeEventListener("pointercancel", this.onPointerCancel);
    this.element.ownerDocument.removeEventListener("keydown", this.onKeyDown, true);
    this.guide?.remove();
    this.guide = null;
    this.tip?.remove();
    this.tip = null;
    const s = this.current;
    if (commit && drag && s && drag.moved) this.commit(drag, s);
    this.drawn = "";
    this.refresh();
  }

  private commit(drag: Drag, s: EditorRulerState): void {
    const firstLine = (f: number): EditorParagraphFormat => f >= 0 ? { firstLineIndent: f } : { hangingIndent: -f };
    switch (drag.kind) {
      case "firstLine":
        if (drag.twips !== s.leftIndent + s.firstLine) this.host.setIndents(firstLine(drag.twips - s.leftIndent));
        return;
      case "hanging":
        // The other lines move; the first line stays where it is.
        if (drag.twips !== s.leftIndent)
          this.host.setIndents({ leftIndent: drag.twips, ...firstLine(s.leftIndent + s.firstLine - drag.twips) });
        return;
      case "leftIndent":
        if (drag.twips !== s.leftIndent) this.host.setIndents({ leftIndent: drag.twips });
        return;
      case "rightIndent":
        if (drag.twips !== s.rightIndent) this.host.setIndents({ rightIndent: drag.twips });
        return;
      case "leftMargin":
        if (drag.twips !== s.marginLeftTwips) this.host.setMargins({ marginLeftTwips: drag.twips });
        return;
      case "rightMargin":
        if (drag.twips !== s.marginRightTwips) this.host.setMargins({ marginRightTwips: drag.twips });
        return;
      case "column": {
        if (drag.twips === 0 || !s.table) return;
        const widths = s.table.widthsTwips.slice();
        const grid = s.table.boundaries[drag.tab].grid;
        widths[grid - 1] += drag.twips;
        if (grid < widths.length) widths[grid] -= drag.twips;
        this.host.setColumnWidths(widths);
        return;
      }
      case "tab": {
        const stops: TabStop[] = s.tabs.map((t) => ({ ...t }));
        const kind = this.tabKindOf(drag);
        const moving = drag.tab >= 0 ? stops.splice(drag.tab, 1)[0] : { positionTwips: 0, alignment: kind, leader: "none" as const };
        if (!drag.removing) {
          // A stop dropped on another replaces it.
          const kept = stops.filter((t) => t.positionTwips !== drag.twips);
          kept.push({ ...moving, positionTwips: drag.twips });
          kept.sort((a, b) => a.positionTwips - b.positionTwips);
          if (drag.tab >= 0 && s.tabs[drag.tab].positionTwips === drag.twips) return;
          this.host.setTabStops(kept);
        } else if (drag.tab >= 0) {
          this.host.setTabStops(stops);
        }
        return;
      }
    }
  }
}
