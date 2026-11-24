/**
 * Word's Outline View, drawn over the continuous view: every body block is marked with its outline
 * level (a heading's own, body text its heading's), indented by it and given Word's symbol in the
 * margin (+ a heading with text under it, - one without, a dot body text); a collapsed heading's
 * part and the levels below the shown one are hidden. Only the view changes: the document does not.
 */
import type { OutlineEntry } from "./types.js";

/** What the layer reads from the editor. */
export interface OutlineLayerHost {
  /** The body's headings in document order (the engine's ListOutline). */
  outline(): OutlineEntry[];
}

/** One body block of the outline, in document order. */
export interface OutlineBlock {
  el: HTMLElement;
  unid: string;
  /** 1-9 for a heading, 0 for body text. */
  level: number;
  /** The heading this block is under (its unid), or null before the first heading. */
  parent: string | null;
  hidden: boolean;
}

export const OUTLINE_CSS = `
.docx-outline .docx-line-numbers { display: none !important; }
.docx-outline [data-outline-hidden] { display: none !important; }
.docx-outline [data-outline-sym] { position: relative; }
.docx-outline [data-outline-sym]::before {
  position: absolute; left: -20px; width: 16px; text-align: center; color: #555;
  font: 14px/1 "Segoe UI Symbol", "Segoe UI", sans-serif; font-weight: normal; font-style: normal;
  text-decoration: none; cursor: pointer; top: 0.25em;
}
.docx-outline [data-outline-sym="plus"]::before { content: "\\2295"; }
.docx-outline [data-outline-sym="minus"]::before { content: "\\2296"; }
.docx-outline [data-outline-sym="body"]::before { content: "\\25CB"; font-size: 9px; top: 0.6em; }
.docx-outline [data-outline-collapsed] { text-decoration: underline dotted #888; }
.docx-outline [data-outline-indent="0"] { margin-left: 24px !important; }
.docx-outline [data-outline-indent="1"] { margin-left: 24px !important; }
.docx-outline [data-outline-indent="2"] { margin-left: calc(24px + 0.5in) !important; }
.docx-outline [data-outline-indent="3"] { margin-left: calc(24px + 1in) !important; }
.docx-outline [data-outline-indent="4"] { margin-left: calc(24px + 1.5in) !important; }
.docx-outline [data-outline-indent="5"] { margin-left: calc(24px + 2in) !important; }
.docx-outline [data-outline-indent="6"] { margin-left: calc(24px + 2.5in) !important; }
.docx-outline [data-outline-indent="7"] { margin-left: calc(24px + 3in) !important; }
.docx-outline [data-outline-indent="8"] { margin-left: calc(24px + 3.5in) !important; }
.docx-outline [data-outline-indent="9"] { margin-left: calc(24px + 4in) !important; }
.docx-outline [data-outline-indent] { text-indent: 0 !important; }
.docx-outline.docx-outline-firstline [data-outline-sym="body"] {
  display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 1; overflow: hidden;
}
.docx-outline.docx-outline-plain [data-outline-sym],
.docx-outline.docx-outline-plain [data-outline-sym] * {
  font-family: Calibri, "Segoe UI", sans-serif !important; font-size: 11pt !important;
  font-weight: normal !important; font-style: normal !important; color: inherit !important;
  text-decoration: none !important; background: none !important;
}
.docx-outline.docx-outline-plain [data-outline-level] { font-weight: bold !important; }
`;

let cssInstalled = new WeakSet<Document>();

function ensureCss(doc: Document): void {
  if (cssInstalled.has(doc)) return;
  cssInstalled.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusOutline = "true";
  style.textContent = OUTLINE_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}

/** The body's top-level blocks under `root`, in order: paragraphs and tables, not their cells'
 *  paragraphs, not notes, header/footer bands or editor chrome. */
export function topLevelBlocks(root: HTMLElement): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>("[data-anchor]")).filter((el) =>
    !el.parentElement?.closest("[data-anchor]")
    && !el.closest("[data-hf-band], .footnotes, .endnotes, .docx-line-numbers"));
}

/** True when `el` sits in a part the outline hides. */
export function isOutlineHidden(el: Element | null | undefined): boolean {
  return !!el?.closest("[data-outline-hidden]");
}

export class OutlineLayer {
  /** The collapsed headings (unids). */
  readonly collapsed = new Set<string>();
  /** Show headings down to this level only (body text hidden), or null: everything. */
  showLevel: number | null = null;
  private root: HTMLElement | null = null;
  private observer: MutationObserver | null = null;
  private frame: number | null = null;
  private blocks: OutlineBlock[] = [];

  constructor(private readonly host: OutlineLayerHost) {}

  get attached(): boolean {
    return this.root !== null;
  }

  /** Watch `root` (the continuous view's flow) and mark its blocks; again after each change. */
  attach(root: HTMLElement): void {
    this.detach();
    ensureCss(root.ownerDocument);
    this.root = root;
    if (typeof MutationObserver !== "undefined") {
      this.observer = new MutationObserver((records) => {
        if (records.every((r) => (r.target.nodeType === 1 ? r.target as Element : r.target.parentElement)?.closest(".docx-line-numbers"))) return;
        this.schedule();
      });
      this.observer.observe(root, { childList: true, subtree: true });
    }
    this.apply();
  }

  detach(): void {
    this.observer?.disconnect();
    this.observer = null;
    if (this.frame !== null && typeof cancelAnimationFrame !== "undefined") cancelAnimationFrame(this.frame);
    this.frame = null;
    if (this.root) {
      for (const el of Array.from(this.root.querySelectorAll<HTMLElement>("[data-outline-sym]"))) clearMarks(el);
    }
    this.root = null;
    this.blocks = [];
  }

  private schedule(): void {
    if (this.frame !== null || typeof requestAnimationFrame === "undefined") return;
    this.frame = requestAnimationFrame(() => {
      this.frame = null;
      this.apply();
    });
  }

  /** The blocks as last marked (marked again first). */
  list(): OutlineBlock[] {
    this.apply();
    return this.blocks;
  }

  /** The outline block `el` is or is in (a cell's paragraph: its table), or null. */
  blockOf(el: Element | null | undefined): OutlineBlock | null {
    if (!el || !this.root) return null;
    const list = this.list();
    return list.find((b) => b.el === el || b.el.contains(el)) ?? null;
  }

  /** The heading `unid` and every block of its part (up to the next heading at its level or above). */
  partOf(unid: string): OutlineBlock[] {
    const list = this.list();
    const i = list.findIndex((b) => b.unid === unid);
    if (i < 0) return [];
    const head = list[i];
    if (head.level === 0) return [head];
    let j = i + 1;
    while (j < list.length && (list[j].level === 0 || list[j].level > head.level)) j++;
    return list.slice(i, j);
  }

  /** Mark every block: level, indent, symbol, hidden. */
  apply(): void {
    const root = this.root;
    if (!root) return;
    const byUnid = new Map<string, number>();
    for (const e of this.host.outline()) {
      const i = e.anchorId.lastIndexOf(":");
      byUnid.set(i >= 0 ? e.anchorId.slice(i + 1) : e.anchorId, e.level);
    }
    const blocks: OutlineBlock[] = [];
    // (the open headings: each with its level and whether it, or one above it, hides its part)
    const stack: Array<{ unid: string; level: number; hides: boolean; index: number }> = [];
    const hasContent = new Set<number>();
    for (const el of topLevelBlocks(root)) {
      const unid = el.getAttribute("data-anchor") ?? "";
      const level = byUnid.get(unid) ?? 0;
      if (level > 0) while (stack.length && stack[stack.length - 1].level >= level) stack.pop();
      const top = stack[stack.length - 1];
      if (top) hasContent.add(top.index);
      const underCollapsed = !!top?.hides;
      const hidden = underCollapsed || (this.showLevel !== null && (level === 0 || level > this.showLevel));
      blocks.push({ el, unid, level, parent: top?.unid ?? null, hidden });
      if (level > 0) stack.push({ unid, level, hides: underCollapsed || this.collapsed.has(unid), index: blocks.length - 1 });
    }
    // (a heading whose part shows nothing under it because of the shown level counts as collapsed)
    blocks.forEach((b, i) => {
      const parentLevel = b.level || (b.parent ? blocks.find((p) => p.unid === b.parent)?.level ?? 0 : 0);
      const shownCollapsed = b.level > 0 && hasContent.has(i)
        && (this.collapsed.has(b.unid) || (this.showLevel !== null && blocks[i + 1]?.hidden === true));
      setAttr(b.el, "data-outline-sym", b.level > 0 ? (hasContent.has(i) ? "plus" : "minus") : "body");
      setAttr(b.el, "data-outline-indent", String(b.level > 0 ? b.level : parentLevel));
      setAttr(b.el, "data-outline-level", b.level > 0 ? String(b.level) : null);
      setAttr(b.el, "data-outline-hidden", b.hidden ? "" : null);
      setAttr(b.el, "data-outline-collapsed", shownCollapsed ? "" : null);
    });
    this.blocks = blocks;
  }

  /** Show every block hiding `el` (expand the headings over it; the shown level goes to all). */
  reveal(el: Element): boolean {
    const list = this.list();
    const b = list.find((x) => x.el === el || x.el.contains(el));
    if (!b || !b.hidden) return false;
    if (this.showLevel !== null && (b.level === 0 || b.level > this.showLevel)) this.showLevel = null;
    let parent = b.parent;
    while (parent) {
      this.collapsed.delete(parent);
      parent = list.find((x) => x.unid === parent)?.parent ?? null;
    }
    this.apply();
    return true;
  }
}

function setAttr(el: HTMLElement, name: string, value: string | null): void {
  if (value === null) {
    if (el.hasAttribute(name)) el.removeAttribute(name);
  } else if (el.getAttribute(name) !== value) {
    el.setAttribute(name, value);
  }
}

function clearMarks(el: HTMLElement): void {
  for (const name of ["data-outline-sym", "data-outline-indent", "data-outline-level", "data-outline-hidden", "data-outline-collapsed"]) {
    el.removeAttribute(name);
  }
}

/** For tests: forget which documents have the stylesheet. */
export function resetOutlineCss(): void {
  cssInstalled = new WeakSet<Document>();
}
