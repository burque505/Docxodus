/**
 * Merge fields in pasted HTML. A field copied in the editor is a span carrying its code
 * (`data-field-code`, stamped by the converter); one copied from Word is its result between
 * `<!--[if supportFields]>` comments holding the code (`mso-element:field-begin`, `-separator`,
 * `-end`). Either would paste as its shown text only («Name»): each is swapped here for a marker
 * (a private-use character pair round its index) for the editor to turn back into the field once
 * the HTML is in. Only the mail-merge kinds (MERGEFIELD, IF, =, FILLIN, ASK, NEXT, NEXTIF, SKIPIF,
 * MERGEREC, MERGESEQ); other fields paste as before.
 */

/** A marker's ends; the field's index is between them, in digits. */
export const FIELD_MARK_OPEN = "";
export const FIELD_MARK_CLOSE = "";
export const FIELD_MARK_PATTERN = /(\d+)/g;

const MERGE_KIND = /^\s*(?:=|(?:MERGEFIELD|IF|FILLIN|ASK|NEXT|NEXTIF|SKIPIF|MERGEREC|MERGESEQ)\b)/i;

/** Whether a field code is one of the mail-merge kinds. */
export function isMergeFieldCode(code: string): boolean {
  return MERGE_KIND.test(code);
}

const mark = (index: number) => `${FIELD_MARK_OPEN}${index}${FIELD_MARK_CLOSE}`;

function decodeEntities(text: string): string {
  return text
    .replace(/&nbsp;/gi, " ")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/gi, "&");
}

/** The text of an HTML fragment (tags dropped, entities decoded). */
function textOf(html: string): string {
  return decodeEntities(html.replace(/<[^>]*>/g, ""));
}

/**
 * Word's fields: the HTML between a field's begin and end (its code, the separator, its result)
 * becomes a marker, the code taken from the text before the separator (nested fields in braces).
 */
function wordFields(html: string, codes: string[]): string {
  const COND = /<!--\[if supportFields\]>([\s\S]*?)<!\[endif\]-->/gi;
  const EVENT = /<span[^>]*mso-element:\s*field-(begin|separator|end)[^>]*>(?:\s*<\/span>)?/gi;
  // the document as a list of pieces: HTML outside the comments, and the events and text inside them
  type Piece = { kind: "html"; text: string } | { kind: "code"; text: string } | { kind: "begin" | "separator" | "end" };
  const pieces: Piece[] = [];
  let last = 0;
  for (const m of html.matchAll(COND)) {
    if (m.index! > last) pieces.push({ kind: "html", text: html.slice(last, m.index) });
    const inner = m[1];
    let at = 0;
    for (const e of inner.matchAll(EVENT)) {
      if (e.index! > at) pieces.push({ kind: "code", text: inner.slice(at, e.index) });
      pieces.push({ kind: e[1].toLowerCase() as "begin" | "separator" | "end" });
      at = e.index! + e[0].length;
    }
    if (at < inner.length) pieces.push({ kind: "code", text: inner.slice(at) });
    last = m.index! + m[0].length;
  }
  if (last === 0) return html;
  if (last < html.length) pieces.push({ kind: "html", text: html.slice(last) });

  // Walk the fields: an outermost field's code is the code text (and nested fields' codes, in
  // braces) before its separator; its result (the HTML after the separator) is what it replaces.
  let out = "";
  type Open = { code: string; separated: boolean; start: number };
  const stack: Open[] = [];
  for (const piece of pieces) {
    if (piece.kind === "begin") {
      if (stack.length && !stack[stack.length - 1].separated) stack[stack.length - 1].code += "{ ";
      stack.push({ code: "", separated: false, start: out.length });
    } else if (piece.kind === "separator") {
      if (stack.length) stack[stack.length - 1].separated = true;
    } else if (piece.kind === "end") {
      const field = stack.pop();
      if (!field) continue;
      const code = field.code.replace(/\s+/g, " ").trim();
      if (stack.length) {
        // (a nested field: its code is part of the outer one's; what it showed is dropped with it)
        const outer = stack[stack.length - 1];
        if (!outer.separated) outer.code += `${code} } `;
        continue;
      }
      if (isMergeFieldCode(code)) {
        out = out.slice(0, field.start) + mark(codes.length);
        codes.push(code);
      }
    } else if (piece.kind === "code") {
      const top = stack[stack.length - 1];
      if (top && !top.separated) top.code += textOf(piece.text);
    } else if (piece.kind !== "html") {
      continue;
    } else if (stack.length && !stack[stack.length - 1].separated) {
      // (HTML inside a code, outside the comments: Word puts the code's own text there at times)
      stack[stack.length - 1].code += textOf(piece.text);
    } else out += piece.text;
  }
  return out;
}

/** The editor's own fields: a `data-field-code` span (its shown text) becomes a marker. */
function editorFields(html: string, codes: string[]): string {
  if (!/data-field-code=/i.test(html)) return html;
  const doc = new DOMParser().parseFromString(html, "text/html");
  const fields = Array.from(doc.querySelectorAll<HTMLElement>("[data-field-code]"))
    .filter((f) => !f.parentElement?.closest("[data-field-code]") && isMergeFieldCode(f.getAttribute("data-field-code") ?? ""));
  if (fields.length === 0) return html;
  for (const field of fields) {
    field.replaceWith(doc.createTextNode(mark(codes.length)));
    codes.push((field.getAttribute("data-field-code") ?? "").trim());
  }
  // (keep Word's StartFragment/EndFragment comments and the rest of the page as it was parsed)
  return doc.documentElement.outerHTML;
}

/** `html` with each merge field a marker, and the fields' codes by index. */
export function extractPastedFields(html: string): { html: string; codes: string[] } {
  const codes: string[] = [];
  const out = editorFields(wordFields(html, codes), codes);
  return { html: out, codes };
}
