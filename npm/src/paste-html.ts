/**
 * The clipboard's HTML, made ready for the engine's `InsertHtml`: from Word, a browser, or this
 * editor's own copy.
 *
 * The engine converts XHTML with its HTML-to-Word converter, which knows HTML's own structure
 * (paragraphs, headings, `ol`/`ul`, tables, links, pictures) and CSS. What the clipboard holds is
 * rarely that clean:
 *
 * - **Word** writes its lists as paragraphs (`style="mso-list:l0 level1 lfo1"`) whose number or
 *   bullet is text in a `mso-list:Ignore` span; leftover markup (`<o:p>`, `<v:shape>`, conditional
 *   comments); a style sheet with `@font-face`, `@page` and `@list` rules; pictures as
 *   `file:///…/clip_image001.png`.
 * - **This editor** renders list items as paragraphs too, the marker in a `data-list-marker` span,
 *   with its other generated chrome (note references, comment markers, rendered deletions).
 * - **Browsers** copy a fragment: inline content straight in `<body>`, remote pictures.
 *
 * `cleanPastedHtml` rebuilds the lists as `ol`/`ul` (levels from Word's `levelN`, or from the
 * paragraphs' indents), drops what is not content, wraps loose inline content in paragraphs, keeps
 * the style sheet's plain rules, and lists the pictures that still need their bytes;
 * `serializePaste` writes the XHTML once each has a `data:` URL or none.
 */

/** What `cleanPastedHtml` makes of the clipboard's HTML. */
export interface CleanedPaste {
  /** The cleaned document (its `<body>` holds the content). */
  readonly doc: Document;
  /** The source's style sheet: its plain rules, the at-rules left out. */
  readonly css: string;
  /** The `src` of each picture whose bytes the page must still find (not a `data:` URL). */
  readonly pendingImages: string[];
  /** True when the content has a picture. */
  readonly hasImages: boolean;
  /** True when nothing is left to paste. */
  readonly empty: boolean;
}

const BLOCK_TAGS = new Set([
  "address", "article", "aside", "blockquote", "center", "dd", "div", "dl", "dt", "figcaption", "figure",
  "footer", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "li", "main", "nav", "ol", "p", "pre",
  "section", "table", "ul",
]);

const DROPPED_TAGS = [
  "head", "script", "style", "meta", "link", "title", "noscript", "template", "iframe", "object", "embed",
  "button", "input", "select", "textarea", "canvas", "audio", "video", "svg", "math",
];

/** Characters XML 1.0 forbids (the engine parses the XHTML as XML). */
const XML_INVALID = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

/** An XML attribute name with no namespace prefix. */
const XML_NAME = /^[A-Za-z_][A-Za-z0-9_.-]*$/;

/**
 * Clean clipboard HTML. `chromeSelector` names the editor's own generated chrome (list markers
 * excepted, which become real lists) so a copy from the editor pastes only its content.
 */
export function cleanPastedHtml(html: string, chromeSelector = ""): CleanedPaste {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const body = doc.body;

  const css = Array.from(doc.querySelectorAll("style")).map((s) => s.textContent ?? "").map(cleanCss).join("\n");

  // Comments: Word's conditional comments (<![if !supportLists]>) parse as comments too.
  const comments: Comment[] = [];
  const walker = doc.createTreeWalker(body, NodeFilter.SHOW_COMMENT);
  while (walker.nextNode()) comments.push(walker.currentNode as Comment);
  comments.forEach((c) => c.remove());

  body.querySelectorAll(DROPPED_TAGS.join(",")).forEach((e) => e.remove());
  if (chromeSelector) {
    body.querySelectorAll(chromeSelector).forEach((e) => {
      if (!e.matches("[data-list-marker]")) e.remove();
    });
  }
  // The editor's hidden separators and placeholder breaks are chrome as well.
  body.querySelectorAll('[aria-hidden="true"], [data-placeholder-break]').forEach((e) => {
    if (!e.matches("[data-list-marker]") && !e.closest("[data-list-marker]")) e.remove();
  });

  // Word's comments: the reference marks ([R1]) and the comment list after the content go; the commented
  // text stays. Its tracked changes paste as they read with the changes accepted: deletions go,
  // insertions stay.
  body.querySelectorAll('a.msocomanchor, .MsoCommentReference, .msocomoff, .msocomtxt, del, .msoDel')
    .forEach((e) => e.remove());
  body.querySelectorAll("div").forEach((d) => {
    if (/mso-element:\s*(comment|footnote|endnote)-list/i.test(d.getAttribute("style") ?? "")) d.remove();
  });
  body.querySelectorAll("ins, .msoIns").forEach((e) => e.replaceWith(...Array.from(e.childNodes)));
  // A link without a target is only a name (Word's bookmarks and comment ranges): its text stays.
  body.querySelectorAll("a:not([href])").forEach((a) => a.replaceWith(...Array.from(a.childNodes)));
  // Word's own links back to its comments and notes inside the copy lead nowhere here.
  body.querySelectorAll('a[href^="#_msocom"], a[href^="#_ftn"], a[href^="#_edn"]').forEach((a) => a.remove());

  rebuildLists(body);

  // Office's own elements: VML shapes go (an <img> carries the same picture), others are unwrapped
  // (<o:p>&nbsp;</o:p> holds a blank paragraph's space).
  Array.from(body.querySelectorAll("*")).filter((e) => e.tagName.includes(":")).forEach((e) => {
    if (/^v:/i.test(e.tagName)) e.remove();
    else e.replaceWith(...Array.from(e.childNodes));
  });

  for (const img of Array.from(body.querySelectorAll("img"))) {
    if (!img.getAttribute("src")) img.remove();
  }

  // A table's border attribute is a border on every cell, as a browser draws it.
  body.querySelectorAll("table[border]").forEach((table) => {
    if ((parseInt(table.getAttribute("border") ?? "0", 10) || 0) <= 0) return;
    table.querySelectorAll("td, th").forEach((cell) => {
      const el = cell as HTMLElement;
      if (!/border/i.test(el.getAttribute("style") ?? "")) el.style.border = "1px solid #000";
    });
  });

  // A paragraph holding only spaces (Word's blank line) is an empty paragraph.
  body.querySelectorAll("p, h1, h2, h3, h4, h5, h6").forEach((p) => {
    if (!p.querySelector("img, br, table") && !(p.textContent ?? "").replace(/[\s\u00a0]/g, "")) p.replaceChildren();
  });

  // HTML5's blocks the converter does not know (a link inside one lost its text): a caption is a
  // paragraph, the rest plain divisions.
  body.querySelectorAll("figcaption, summary").forEach((e) => renameElement(e, "p"));
  body.querySelectorAll("figure, section, article, header, footer, nav, aside, main, details, address, hgroup")
    .forEach((e) => renameElement(e, "div"));

  dropLayoutWhitespace(body);
  wrapLooseInline(doc, body);
  cleanAttributes(body);

  const textWalker = doc.createTreeWalker(body, NodeFilter.SHOW_TEXT);
  while (textWalker.nextNode()) {
    const t = textWalker.currentNode as Text;
    if (XML_INVALID.test(t.data)) t.data = t.data.replace(XML_INVALID, "");
    XML_INVALID.lastIndex = 0;
  }

  const images = Array.from(body.querySelectorAll("img"));
  const pendingImages = [...new Set(images.map((i) => i.getAttribute("src")!).filter((src) => !/^data:/i.test(src)))];
  const empty = !body.querySelector("img, table, hr") && !(body.textContent ?? "").trim()
    && body.querySelectorAll("p, h1, h2, h3, h4, h5, h6, li").length === 0;
  return { doc, css, pendingImages, hasImages: images.length > 0, empty };
}

/**
 * The XHTML for `InsertHtml`. `images` gives each pending picture's `data:` URL; a picture with
 * none (or `null`) is left out.
 */
export function serializePaste(cleaned: CleanedPaste, images: ReadonlyMap<string, string | null> = new Map()): string {
  const doc = cleaned.doc;
  for (const img of Array.from(doc.body.querySelectorAll("img"))) {
    const src = img.getAttribute("src") ?? "";
    if (/^data:image\//i.test(src)) continue;
    const found = images.get(src);
    if (found) img.setAttribute("src", found);
    else img.remove();
  }
  // Only <body> goes: text between <html> and <body> would become a paragraph of its own.
  for (const node of Array.from(doc.documentElement.childNodes)) if (node !== doc.body) node.remove();
  // A fresh <html> around the body: Word's carries namespace declarations, one of them a second xmlns, which
  // XML does not allow.
  const html = new DOMParser().parseFromString("<!doctype html><html><body></body></html>", "text/html");
  const body = html.importNode(doc.body, true) as HTMLElement;
  for (const attr of Array.from(body.attributes)) if (attr.name !== "style" && attr.name !== "lang") body.removeAttribute(attr.name);
  html.body.replaceWith(body);
  return new XMLSerializer().serializeToString(html.documentElement);
}

/** A style sheet's plain rules: comments, HTML comment marks and at-rules (@font-face, @page,
 *  Word's @list, @media) left out, which the converter has no use for. */
function cleanCss(css: string): string {
  let s = css.replace(/\/\*[\s\S]*?\*\//g, "").replace(/<!--|-->/g, "");
  // Remove at-rules with a block (one level of nesting covers @media's).
  for (let i = 0; i < 4; i++) s = s.replace(/@[a-z-]+[^{};]*\{(?:[^{}]|\{[^{}]*\})*\}/gi, "");
  return s.replace(/@[a-z-]+[^{};]*;/gi, "").trim();
}

/** Replace an element with one of another name, keeping its attributes and content. */
function renameElement(el: Element, tag: string): void {
  const next = el.ownerDocument.createElement(tag);
  for (const attr of Array.from(el.attributes)) next.setAttribute(attr.name, attr.value);
  next.append(...Array.from(el.childNodes));
  el.replaceWith(next);
}

/**
 * Put loose inline content (text, links, pictures) beside blocks into paragraphs: at the top level of
 * the body, and in a division, quotation or cell that also holds blocks. The converter drops inline
 * content that stands beside a block (a figure's picture above its caption).
 */
function wrapLooseInline(doc: Document, body: HTMLElement): void {
  wrapLooseInlineIn(doc, body);
  for (const el of Array.from(body.querySelectorAll("div, blockquote, td, th, center"))) {
    if (Array.from(el.children).some((c) => BLOCK_TAGS.has(c.tagName.toLowerCase()))) wrapLooseInlineIn(doc, el);
  }
}

function wrapLooseInlineIn(doc: Document, container: Element): void {
  let run: Node[] = [];
  const flush = (before: Node | null) => {
    if (run.length && run.some((n) => n.nodeType === 1 || (n.textContent ?? "").trim())) {
      const p = doc.createElement("p");
      container.insertBefore(p, before);
      run.forEach((n) => p.append(n));
    } else run.forEach((n) => n.parentNode?.removeChild(n));
    run = [];
  };
  for (const node of Array.from(container.childNodes)) {
    const block = node.nodeType === 1 && BLOCK_TAGS.has((node as Element).tagName.toLowerCase());
    if (block) flush(node);
    else if (node.nodeType === 1 || node.nodeType === 3) run.push(node);
    else node.parentNode?.removeChild(node);
  }
  flush(null);
  // A <br> between two blocks is a blank line; at the edge of a block it is only the copy's end.
  container.querySelectorAll(":scope > p").forEach((p) => {
    if (p.childNodes.length === 1 && (p.firstChild as Element).tagName === "BR") p.replaceChildren();
  });
}

const CONTAINER_TAGS = new Set([
  "body", "div", "td", "th", "li", "table", "tbody", "thead", "tfoot", "tr", "ul", "ol", "blockquote", "section",
  "article", "header", "footer", "main", "nav", "aside", "figure", "dl", "center",
]);

/**
 * The line breaks and indents between blocks (Word's HTML is laid out with them) are not text: a
 * whitespace-only text node beside a block, or anywhere in a table's structure, would become a
 * paragraph of its own. A <div> that only holds blocks (Word's section wrapper) is unwrapped.
 */
function dropLayoutWhitespace(body: HTMLElement): void {
  const isBlock = (n: Node | null) => !!n && n.nodeType === 1 && BLOCK_TAGS.has((n as Element).tagName.toLowerCase());
  const walker = body.ownerDocument.createTreeWalker(body, NodeFilter.SHOW_TEXT);
  const drop: Text[] = [];
  while (walker.nextNode()) {
    const t = walker.currentNode as Text;
    if (/[^\s]/.test(t.data) || / /.test(t.data)) continue;  // (a no-break space is content; \s takes it too)
    const parent = t.parentElement;
    if (!parent || !CONTAINER_TAGS.has(parent.tagName.toLowerCase())) continue;
    const tableish = /^(table|tbody|thead|tfoot|tr|ul|ol|dl)$/i.test(parent.tagName);
    if (tableish || isBlock(t.previousSibling) || isBlock(t.nextSibling) || !t.previousSibling || !t.nextSibling) drop.push(t);
  }
  drop.forEach((t) => t.remove());
  for (const div of Array.from(body.querySelectorAll("div")).reverse()) {
    if (Array.from(div.childNodes).every((n) => isBlock(n))) div.replaceWith(...Array.from(div.childNodes));
  }
}

/** Keep only attributes XML can carry and the converter reads; drop the rest. */
function cleanAttributes(root: HTMLElement): void {
  for (const el of [root, ...Array.from(root.querySelectorAll("*"))]) {
    for (const attr of Array.from(el.attributes)) {
      const name = attr.name;
      if (!XML_NAME.test(name) || /^on/i.test(name) || /^data-/i.test(name) || name === "contenteditable"
        || name === "aria-hidden" || name === "tabindex" || name === "role") {
        el.removeAttribute(name);
      }
    }
    cleanStyle(el);
  }
}

/** The CSS properties the converter knows. Anything else (Word's mso-*, custom properties, grid,
 *  transforms, ...) is left out: one value its parser cannot read fails the whole paste. */
const KEPT_CSS = new Set([
  "color", "background-color", "font-family", "font-size", "font-weight", "font-style", "font-variant",
  "text-decoration", "text-decoration-line", "text-align", "text-indent", "text-transform", "vertical-align",
  "white-space", "direction", "list-style-type",
  "margin-top", "margin-right", "margin-bottom", "margin-left",
  "padding-top", "padding-right", "padding-bottom", "padding-left",
  "border-top-width", "border-right-width", "border-bottom-width", "border-left-width",
  "border-top-style", "border-right-style", "border-bottom-style", "border-left-style",
  "border-top-color", "border-right-color", "border-bottom-color", "border-left-color",
  "border-collapse", "width",
]);

/** A value the converter's CSS parser reads: no functions but rgb()/rgba(), no slashes, no !important. */
function plainCssValue(value: string): boolean {
  if (!value || /!important|\/|\bvar\(|\bcalc\(|\bclamp\(|\bmin\(|\bmax\(|\benv\(|\bcolor-mix\(|\blight-dark\(|\boklch\(|\blab\(|\bhsla?\(|\burl\(/i.test(value)) return false;
  const functions = value.match(/[a-z-]+\(/gi) ?? [];
  return functions.every((f) => /^rgba?\($/i.test(f));
}

/** Keep the style declarations the converter can read (the browser has parsed them already). */
function cleanStyle(el: Element): void {
  const style = (el as HTMLElement).style;
  if (!style || !el.hasAttribute("style")) return;
  const kept: string[] = [];
  for (let i = 0; i < style.length; i++) {
    const name = style.item(i);
    const value = style.getPropertyValue(name).trim();
    if (!KEPT_CSS.has(name) || !plainCssValue(value)) continue;
    // (a width only on tables and cells, in units the converter takes)
    if (name === "width" && (!/^(TABLE|TD|TH|COL|IMG)$/.test(el.tagName) || !/^[\d.]+(px|pt|in|cm|mm|%)$/.test(value))) continue;
    // (vertical-align: only what is a run's raise or lower, or a cell's alignment)
    if (name === "vertical-align" && !/^(sub|super|baseline|top|middle|bottom)$/.test(value)) continue;
    // (transparent backgrounds are no background)
    if (name === "background-color" && /^(transparent|rgba\([^)]*,\s*0\))$/.test(value)) continue;
    kept.push(`${name}: ${value}`);
  }
  if (kept.length) el.setAttribute("style", kept.join("; "));
  else el.removeAttribute("style");
}

// ─── Lists ─────────────────────────────────────────────────────────────

interface ListParagraph {
  el: Element;
  /** 1-based nesting level. */
  level: number;
  /** Which list it belongs to (Word's lfo), or null. */
  list: string | null;
  marker: string;
  ordered: boolean;
  styleType: string;
}

/** A paragraph's list marker, removed from it, and what it says about the list. */
function takeListParagraph(el: Element): ListParagraph | null {
  const style = el.getAttribute("style") ?? "";
  const wordList = /mso-list:\s*(l\d+)\s+level(\d+)\s+(lfo\d+)/i.exec(style);
  let markerEl: Element | null = null;
  if (wordList) {
    markerEl = Array.from(el.querySelectorAll("span")).find((s) => /mso-list:\s*Ignore/i.test(s.getAttribute("style") ?? "")) ?? null;
    // The marker span sits in a span setting its font (Symbol, Wingdings): drop that too when it
    // holds nothing else.
    while (markerEl?.parentElement && markerEl.parentElement !== el
      && markerEl.parentElement.textContent === markerEl.textContent) markerEl = markerEl.parentElement;
  } else {
    markerEl = el.querySelector(":scope > [data-list-marker]");
    if (!markerEl) return null;
  }
  const marker = (markerEl?.textContent ?? "").replace(/[\u00a0\s]+/g, " ").trim();
  markerEl?.remove();
  if (!wordList && !marker) return null;
  const kind = markerKind(marker, el.getAttribute("style") ?? "");
  return {
    el,
    level: wordList ? Number(wordList[2]) : 0,
    list: wordList ? `${wordList[1]}/${wordList[3]}` : null,
    marker,
    ...kind,
  };
}

/** Ordered or bulleted, and the CSS list-style-type that matches the marker. */
function markerKind(marker: string, style: string): { ordered: boolean; styleType: string } {
  const m = /^\(?([0-9]+|[a-zA-Z]+)[.)]?\)?$/.exec(marker);
  if (m) {
    const n = m[1];
    if (/^[0-9]+$/.test(n)) return { ordered: true, styleType: "decimal" };
    if (/^[ivxlcdm]+$/.test(n) && (n.length > 1 || "ivx".includes(n))) return { ordered: true, styleType: "lower-roman" };
    if (/^[IVXLCDM]+$/.test(n) && (n.length > 1 || "IVX".includes(n))) return { ordered: true, styleType: "upper-roman" };
    if (n.length === 1 && n !== "o") return { ordered: true, styleType: n === n.toLowerCase() ? "lower-alpha" : "upper-alpha" };
  }
  // Word's bullets are Symbol/Wingdings characters; "o" in Courier New is its hollow one.
  if (marker === "o" || marker === "\u25e6") return { ordered: false, styleType: "circle" };
  if (/[\u00a7\u25aa\u25a0\uf0a7]/.test(marker)) return { ordered: false, styleType: "square" };
  void style;
  return { ordered: false, styleType: "disc" };
}

/** An element's left indent in px as its inline style gives it (0 when it doesn't). */
function indentOf(el: Element): number {
  const style = (el as HTMLElement).style;
  const px = (v: string) => {
    const n = parseFloat(v);
    if (!Number.isFinite(n)) return 0;
    if (/in$/.test(v)) return n * 96;
    if (/cm$/.test(v)) return n * 96 / 2.54;
    if (/pt$/.test(v)) return n * 96 / 72;
    return n;
  };
  return px(style.marginLeft) + px(style.paddingLeft);
}

/**
 * Turn runs of list paragraphs into nested `ol`/`ul` with `li` items. Word's say their level;
 * this editor's are ranked by indent within the run.
 */
function rebuildLists(body: HTMLElement): void {
  const candidates = Array.from(body.querySelectorAll("p, h1, h2, h3, h4, h5, h6"))
    .filter((el) => /mso-list:\s*l\d+\s+level/i.test(el.getAttribute("style") ?? "")
      || el.querySelector(":scope > [data-list-marker]"));
  const items = candidates.map(takeListParagraph).filter((x): x is ListParagraph => !!x);
  if (!items.length) return;

  // Group into runs of items that follow each other (only whitespace between).
  const runs: ListParagraph[][] = [];
  for (const item of items) {
    const last = runs[runs.length - 1]?.at(-1);
    if (last && nextElementSkippingSpace(last.el) === item.el) runs[runs.length - 1].push(item);
    else runs.push([item]);
  }

  for (const run of runs) {
    // This editor's items: rank distinct indents to levels.
    if (run.some((i) => i.level === 0)) {
      const indents = [...new Set(run.map((i) => Math.round(indentOf(i.el))))].sort((a, b) => a - b);
      for (const i of run) if (i.level === 0) i.level = indents.indexOf(Math.round(indentOf(i.el))) + 1;
    }
    const doc = body.ownerDocument;
    const place = doc.createTextNode("");
    run[0].el.before(place);
    const lists = doc.createDocumentFragment();
    const stack: { level: number; list: Element; key: string }[] = [];
    for (const item of run) {
      const key = `${item.list ?? ""}|${item.ordered ? "o" : "u"}|${item.styleType}`;
      while (stack.length && stack[stack.length - 1].level > item.level) stack.pop();
      let current = stack[stack.length - 1];
      if (current && current.level === item.level && current.key !== key) {
        // Another list at the same level: it starts beside this one.
        stack.pop();
        const next = makeList(doc, item);
        current.list.after(next);
        current = { level: item.level, list: next, key };
        stack.push(current);
      } else if (!current || current.level < item.level) {
        const list = makeList(doc, item);
        if (current) (current.list.lastElementChild ?? current.list.appendChild(doc.createElement("li"))).append(list);
        else lists.append(list);
        current = { level: item.level, list, key };
        stack.push(current);
      }
      const li = doc.createElement("li");
      li.append(...Array.from(item.el.childNodes));
      current.list.append(li);
      item.el.remove();
    }
    place.replaceWith(lists);
  }
}

function makeList(doc: Document, item: ListParagraph): Element {
  const list = doc.createElement(item.ordered ? "ol" : "ul");
  list.setAttribute("style", `list-style-type: ${item.styleType}`);
  if (item.ordered && item.styleType === "decimal") {
    const start = parseInt(item.marker.replace(/\D/g, ""), 10);
    if (start > 1) list.setAttribute("start", String(start));
  }
  return list;
}

function nextElementSkippingSpace(el: Element): Element | null {
  let n = el.nextSibling;
  while (n && n.nodeType !== 1) {
    if (n.nodeType === 3 && (n.textContent ?? "").trim()) return null;
    n = n.nextSibling;
  }
  return n as Element | null;
}

/** Plain text from the clipboard, as the engine takes it (Windows line ends kept as lines). */
export function pastedPlainText(text: string): string {
  return text.replace(XML_INVALID, "");
}
