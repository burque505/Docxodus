/**
 * Word's line numbers in the paginated view. Each body section's `w:lnNumType` reaches the page
 * container as `data-line-number-sections`, a JSON map from section index to { countBy, start,
 * distance (pt, absent = Word's automatic gap), restart (newPage, newSection, continuous) }. The
 * numbers are drawn from the laid-out lines, in a layer of their own in the margin left of each
 * text column, as Word's Print Layout does: body paragraphs only (not tables, headers and footers,
 * notes or text boxes), each paragraph numbered by its own section's settings (a page can hold
 * several sections), an empty paragraph counting as one line, a paragraph with
 * `data-suppress-line-numbers` neither numbered nor counted, and only every count-by-th number shown.
 * In a multi-column section the lines are numbered column by column.
 */

/** Word's automatic distance between the numbers and the text: a quarter inch. */
const AUTO_DISTANCE_PT = 18;
const LAYER_CLASS = "docx-line-numbers";
const SECTIONS_ATTRIBUTE = "data-line-number-sections";
const PARAGRAPHS = "p, h1, h2, h3, h4, h5, h6, li";
const SKIP_SELECTOR = [
  "table", "header", "footer", "[data-hf-band]", "[class*='header']", "[class*='footer']",
  "[class*='footnotes']", "[class*='endnotes']", "[class*='textbox']", "[class*='comment']",
  "." + LAYER_CLASS,
].join(", ");

/** One section's line numbering, as the page container carries it. */
export interface SectionLineNumbers {
  countBy: number;
  start: number;
  distance?: number;
  restart: "newPage" | "newSection" | "continuous";
}

/** Record the body sections' line numbering on the page container (the paginator does this). */
export function setSectionLineNumbers(container: HTMLElement, sections: Map<number, SectionLineNumbers>): void {
  if (sections.size) container.setAttribute(SECTIONS_ATTRIBUTE, JSON.stringify(Object.fromEntries(sections)));
  else container.removeAttribute(SECTIONS_ATTRIBUTE);
}

/** A section's line numbering from its section element's `data-line-number*` attributes, or null. */
export function readSectionLineNumbers(section: HTMLElement): SectionLineNumbers | null {
  const countBy = parseInt(section.dataset.lineNumbers ?? "", 10);
  if (!(countBy > 0)) return null;
  const restart = section.dataset.lineNumberRestart;
  const distance = parseFloat(section.dataset.lineNumberDistance ?? "");
  return {
    countBy,
    start: parseInt(section.dataset.lineNumberStart ?? "1", 10) || 1,
    ...(Number.isFinite(distance) ? { distance } : {}),
    restart: restart === "continuous" || restart === "newSection" ? restart : "newPage",
  };
}

function sectionsOf(root: ParentNode): Map<string, SectionLineNumbers> | null {
  const host = root instanceof Element && root.hasAttribute(SECTIONS_ATTRIBUTE)
    ? root
    : root.querySelector(`[${SECTIONS_ATTRIBUTE}]`);
  const raw = host?.getAttribute(SECTIONS_ATTRIBUTE);
  if (!raw) return null;
  try {
    return new Map(Object.entries(JSON.parse(raw) as Record<string, SectionLineNumbers>));
  } catch {
    return null;
  }
}

/** Whether any section under `root` numbers its lines. */
export function hasLineNumbers(root: ParentNode): boolean {
  return !!sectionsOf(root)?.size;
}

/** Remove the drawn numbers under `root`. */
export function clearLineNumbers(root: ParentNode): void {
  root.querySelectorAll("." + LAYER_CLASS).forEach((layer) => layer.remove());
}

/** Draw (again) the line numbers of every page under `root`. */
export function drawLineNumbers(root: ParentNode): void {
  clearLineNumbers(root);
  const sections = sectionsOf(root);
  if (!sections?.size) return;
  const pages = Array.from(root.querySelectorAll<HTMLElement>("[data-page-number][data-section-index]"));
  // The count runs over the whole document: a section that restarts each page starts again on each
  // page, one that restarts each section where it begins, and a continuous one goes on from the
  // section before.
  let counter = 0;
  let lastSection: string | null = null;
  for (const page of pages) {
    let pageStarted = false;
    const layer = newLayer(page);
    for (const paragraph of bodyParagraphs(page)) {
      const section = paragraph.closest<HTMLElement>("[data-pagination-section]")?.dataset.paginationSection
        ?? page.dataset.sectionIndex ?? "0";
      const settings = sections.get(section);
      if (!settings) { lastSection = section; continue; }
      if (section !== lastSection) {
        if (settings.restart !== "continuous" || lastSection === null || !sections.has(lastSection)) counter = settings.start - 1;
        lastSection = section;
        pageStarted = true;
      } else if (!pageStarted) {
        if (settings.restart === "newPage") counter = settings.start - 1;
        pageStarted = true;
      }
      // A suppressed paragraph is not counted; nor is the empty one that only carries a section
      // break (Word gives its mark no line of its own).
      if (paragraph.dataset.suppressLineNumbers === "true" || isSectionMarkOnly(paragraph)) continue;
      counter = drawParagraph(layer, page, paragraph, settings, counter);
    }
    if (layer.childNodes.length) page.appendChild(layer);
  }
}

function isSectionMarkOnly(paragraph: HTMLElement): boolean {
  return paragraph.dataset.sectionMark === "true" && !(paragraph.textContent ?? "").replace(/[\s\u00a0\u200b]/g, "");
}

function newLayer(page: HTMLElement): HTMLElement {
  const layer = page.ownerDocument.createElement("div");
  layer.className = LAYER_CLASS;
  layer.setAttribute("aria-hidden", "true");
  layer.setAttribute("contenteditable", "false");
  layer.style.cssText = "position:absolute;inset:0;pointer-events:none;user-select:none;z-index:1;";
  return layer;
}

function bodyParagraphs(page: HTMLElement): HTMLElement[] {
  return Array.from(page.querySelectorAll<HTMLElement>(PARAGRAPHS))
    .filter((p) => !p.closest(SKIP_SELECTOR) && !p.parentElement?.closest(PARAGRAPHS));
}

/** Number one paragraph's lines from `counter` on; returns the count after its last line. */
function drawParagraph(layer: HTMLElement, page: HTMLElement, paragraph: HTMLElement,
  settings: SectionLineNumbers, counter: number): number {
  const view = page.ownerDocument.defaultView;
  if (!view) return counter;
  const pageRect = page.getBoundingClientRect();
  // Page boxes may be zoomed: client rectangles are in zoomed pixels, the layer's in the box's own.
  const zoom = page.offsetWidth > 0 ? pageRect.width / page.offsetWidth : 1;
  if (!(zoom > 0)) return counter;
  const distancePx = ptToPx(settings.distance ?? AUTO_DISTANCE_PT);
  const columns = columnsOf(paragraph, page);
  const style = view.getComputedStyle(paragraph);
  for (const line of lineBoxes(paragraph, columns)) {
    counter++;
    if (counter % settings.countBy !== 0) continue;
    const number = page.ownerDocument.createElement("span");
    number.textContent = String(counter);
    const top = (line.top - pageRect.top) / zoom;
    const height = line.height / zoom;
    const right = (pageRect.right - line.columnLeft) / zoom + distancePx;
    number.style.cssText = `position:absolute;top:${top}px;height:${height}px;line-height:${height}px;`
      + `right:${right}px;white-space:nowrap;font-family:${style.fontFamily};font-size:${numberSize(style.fontSize)};`
      + "color:inherit;";
    layer.appendChild(number);
  }
  return counter;
}

/** The text columns a paragraph can sit in: their left edges in client pixels, left to right (one
 *  for a single-column section: the page's content area, the page box's child that holds it). */
function columnsOf(paragraph: HTMLElement, page: HTMLElement): number[] {
  const view = paragraph.ownerDocument.defaultView;
  if (!view) return [];
  let content: HTMLElement = paragraph;
  while (content.parentElement && content.parentElement !== page) content = content.parentElement;
  let multi: HTMLElement | null = null;
  for (let el = paragraph.parentElement; el && el !== page; el = el.parentElement) {
    if (parseInt(view.getComputedStyle(el).columnCount, 10) > 1) { multi = el; break; }
  }
  const area = multi ?? content;
  const box = area.getBoundingClientRect();
  const style = view.getComputedStyle(area);
  const zoom = area.offsetWidth > 0 ? box.width / area.offsetWidth : 1;
  const padLeft = (parseFloat(style.paddingLeft) || 0) * zoom;
  const left = box.left + padLeft + (parseFloat(style.borderLeftWidth) || 0) * zoom;
  if (!multi) return [left];
  const count = parseInt(style.columnCount, 10);
  const gap = (parseFloat(style.columnGap) || 0) * zoom;
  const inner = box.width - padLeft - (parseFloat(style.paddingRight) || 0) * zoom;
  const width = (inner - gap * (count - 1)) / count;
  return Array.from({ length: count }, (_, i) => left + i * (width + gap));
}

/** The column (index) whose left edge is the last one at or before `x`. */
function columnAt(columns: number[], x: number): number {
  let index = 0;
  for (let i = 0; i < columns.length; i++) if (x >= columns[i] - 2) index = i;
  return index;
}

/**
 * The paragraph's laid-out lines, column by column and top to bottom within a column, from its
 * text's client rectangles. Rectangles share a line only within one column (a paragraph that flows
 * on into the next column has lines at the same height there). An empty paragraph is one line,
 * its own box.
 */
function lineBoxes(paragraph: HTMLElement, columns: number[]):
  Array<{ top: number; height: number; columnLeft: number }> {
  const range = paragraph.ownerDocument.createRange();
  range.selectNodeContents(paragraph);
  const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);
  if (!rects.length) {
    const box = paragraph.getBoundingClientRect();
    if (!(box.height > 0)) return [];
    const column = columns[columnAt(columns, box.left)] ?? box.left;
    return [{ top: box.top, height: Math.min(box.height, lineHeightPx(paragraph) ?? box.height), columnLeft: column }];
  }
  const lines: Array<{ column: number; top: number; bottom: number; left: number }> = [];
  for (const r of rects) {
    const column = columnAt(columns, r.left);
    const middle = (r.top + r.bottom) / 2;
    const line = lines.find((l) => l.column === column && middle >= l.top && middle <= l.bottom);
    if (line) {
      line.top = Math.min(line.top, r.top);
      line.bottom = Math.max(line.bottom, r.bottom);
    } else lines.push({ column, top: r.top, bottom: r.bottom, left: r.left });
  }
  lines.sort((a, b) => a.column - b.column || a.top - b.top);
  return lines.map((l) => ({ top: l.top, height: l.bottom - l.top, columnLeft: columns[l.column] ?? l.left }));
}

function lineHeightPx(el: HTMLElement): number | null {
  const value = parseFloat(el.ownerDocument.defaultView?.getComputedStyle(el).lineHeight ?? "");
  return Number.isFinite(value) ? value : null;
}

/** Word numbers lines in its Line Number style, the body text's size: a heading's own size is
 *  capped at 11pt so its numbers match the rest. */
function numberSize(size: string): string {
  const px = parseFloat(size);
  const cap = ptToPx(11);
  return Number.isFinite(px) && px > cap ? `${cap}px` : size;
}

function ptToPx(pt: number): number {
  return (pt * 96) / 72;
}
