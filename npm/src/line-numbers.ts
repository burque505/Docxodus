/**
 * Word's line numbers in the paginated view. A section's `w:lnNumType` reaches the page boxes as
 * `data-line-numbers` (count by), `data-line-number-start`, `data-line-number-distance` (pt, absent
 * = Word's automatic gap) and `data-line-number-restart` (newPage, newSection, continuous). The
 * numbers are drawn from the laid-out lines, in a layer of their own in the margin left of the text
 * column, as Word's Print Layout does: body paragraphs only (not tables, headers and footers,
 * notes or text boxes), an empty paragraph counting as one line, a paragraph with
 * `data-suppress-line-numbers` neither numbered nor counted, and only every count-by-th number shown.
 */

/** Word's automatic distance between the numbers and the text: a quarter inch. */
const AUTO_DISTANCE_PT = 18;
const LAYER_CLASS = "docx-line-numbers";
const SKIP_SELECTOR = [
  "table", "header", "footer", "[data-hf-band]", "[class*='header']", "[class*='footer']",
  "[class*='footnotes']", "[class*='endnotes']", "[class*='textbox']", "[class*='comment']",
  "." + LAYER_CLASS,
].join(", ");

/** Whether any page under `root` numbers its lines. */
export function hasLineNumbers(root: ParentNode): boolean {
  return !!root.querySelector("[data-line-numbers]");
}

/** Remove the drawn numbers under `root`. */
export function clearLineNumbers(root: ParentNode): void {
  root.querySelectorAll("." + LAYER_CLASS).forEach((layer) => layer.remove());
}

/** Draw (again) the line numbers of every page box under `root` that carries `data-line-numbers`. */
export function drawLineNumbers(root: ParentNode): void {
  clearLineNumbers(root);
  const pages = Array.from(root.querySelectorAll<HTMLElement>("[data-page-number][data-section-index]"));
  let counter = 0;
  let lastSection: string | null = null;
  let numberedBefore = false;
  for (const page of pages) {
    const countBy = parseInt(page.dataset.lineNumbers ?? "", 10);
    const section = page.dataset.sectionIndex ?? "";
    if (!(countBy > 0)) {
      numberedBefore = false;
      lastSection = section;
      continue;
    }
    const start = parseInt(page.dataset.lineNumberStart ?? "1", 10) || 1;
    const restart = page.dataset.lineNumberRestart ?? "newPage";
    const newSection = section !== lastSection;
    if (restart === "newPage" || !numberedBefore || (restart === "newSection" && newSection)) counter = start - 1;
    lastSection = section;
    numberedBefore = true;
    counter = drawPage(page, countBy, counter);
  }
}

/** Number one page's lines from `counter` on; returns the count after its last line. */
function drawPage(page: HTMLElement, countBy: number, counter: number): number {
  const doc = page.ownerDocument;
  const view = doc.defaultView;
  if (!view) return counter;
  const pageRect = page.getBoundingClientRect();
  // Page boxes may be zoomed: client rectangles are in zoomed pixels, the layer's in the box's own.
  const zoom = page.offsetWidth > 0 ? pageRect.width / page.offsetWidth : 1;
  if (!(zoom > 0)) return counter;
  const distancePx = ptToPx(parseFloat(page.dataset.lineNumberDistance ?? "") || AUTO_DISTANCE_PT);

  const layer = doc.createElement("div");
  layer.className = LAYER_CLASS;
  layer.setAttribute("aria-hidden", "true");
  layer.setAttribute("contenteditable", "false");
  layer.style.cssText = "position:absolute;inset:0;pointer-events:none;user-select:none;z-index:1;";

  const paragraphs = Array.from(page.querySelectorAll<HTMLElement>("p, h1, h2, h3, h4, h5, h6, li"))
    .filter((p) => !p.closest(SKIP_SELECTOR) && !p.parentElement?.closest("p, h1, h2, h3, h4, h5, h6, li"));
  for (const paragraph of paragraphs) {
    if (paragraph.dataset.suppressLineNumbers === "true") continue;
    const lines = lineBoxes(paragraph);
    if (!lines.length) continue;
    const column = columnLeft(paragraph, page);
    const style = view.getComputedStyle(paragraph);
    for (const line of lines) {
      counter++;
      if (counter % countBy !== 0) continue;
      const number = doc.createElement("span");
      number.textContent = String(counter);
      const top = (line.top - pageRect.top) / zoom;
      const height = line.height / zoom;
      const right = (pageRect.right - (column ?? line.left)) / zoom + distancePx;
      number.style.cssText = `position:absolute;top:${top}px;height:${height}px;line-height:${height}px;`
        + `right:${right}px;white-space:nowrap;font-family:${style.fontFamily};font-size:${numberSize(style.fontSize)};`
        + "color:inherit;";
      layer.appendChild(number);
    }
  }
  if (layer.childNodes.length) page.appendChild(layer);
  return counter;
}

/** The paragraph's laid-out lines (top, height, left), from its text's client rectangles; an empty
 *  paragraph is one line, its own box. */
function lineBoxes(paragraph: HTMLElement): Array<{ top: number; height: number; left: number }> {
  const range = paragraph.ownerDocument.createRange();
  range.selectNodeContents(paragraph);
  const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);
  if (!rects.length) {
    const box = paragraph.getBoundingClientRect();
    return box.height > 0 ? [{ top: box.top, height: Math.min(box.height, lineHeightPx(paragraph) ?? box.height), left: box.left }] : [];
  }
  // One entry per line: rectangles whose vertical middles fall within a line already seen join it.
  const lines: Array<{ top: number; bottom: number; left: number }> = [];
  for (const r of rects.sort((a, b) => a.top - b.top || a.left - b.left)) {
    const middle = (r.top + r.bottom) / 2;
    const line = lines.find((l) => middle >= l.top && middle <= l.bottom && Math.abs(r.left - l.left) < 10000);
    if (line) {
      line.top = Math.min(line.top, r.top);
      line.bottom = Math.max(line.bottom, r.bottom);
      line.left = Math.min(line.left, r.left);
    } else lines.push({ top: r.top, bottom: r.bottom, left: r.left });
  }
  return lines.map((l) => ({ top: l.top, height: l.bottom - l.top, left: l.left }));
}

/** The left edge of the text column the paragraph sits in (client pixels): the page's content
 *  area (the page box's child that holds it), or in a multi-column section the edge of the column
 *  under the paragraph. */
function columnLeft(paragraph: HTMLElement, page: HTMLElement): number | null {
  const view = paragraph.ownerDocument.defaultView;
  if (!view) return null;
  let content: HTMLElement = paragraph;
  while (content.parentElement && content.parentElement !== page) content = content.parentElement;
  if (content.parentElement !== page) return null;
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
  if (!multi) return left;
  const count = parseInt(style.columnCount, 10);
  const gap = (parseFloat(style.columnGap) || 0) * zoom;
  const inner = box.width - padLeft - (parseFloat(style.paddingRight) || 0) * zoom;
  const width = (inner - gap * (count - 1)) / count;
  const at = paragraph.getBoundingClientRect().left - left;
  const index = Math.max(0, Math.min(count - 1, Math.floor((at + gap / 2) / (width + gap))));
  return left + index * (width + gap);
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
