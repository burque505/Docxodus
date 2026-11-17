// Format Painter's pure parts: which run formatting a copy takes, the op that paints it, and the
// word a click paints. The editor (DocxEditor.copyFormatting / pasteFormatting / formatPainter)
// does the reading and writing.

import type { InlineSpan, RunFormattingInfo } from "./types.js";

/** A character-formatting op (the engine's FormatOp) that makes a run's direct formatting exactly
 *  `direct`: every property `direct` leaves out is cleared, as Word's Format Painter does. */
export function runOpFromDirect(direct: RunFormattingInfo): Record<string, unknown> {
  return {
    bold: direct.bold ?? false,
    italic: direct.italic ?? false,
    underline: direct.underline ?? false,
    strike: direct.strike ?? false,
    // (Code formatting is a character style, so runStyle carries it.)
    color: direct.color ?? "",
    highlight: direct.highlight ?? "",
    vertAlign: direct.vertAlign ?? "",
    fontSizePts: direct.fontSizePts ?? 0,
    fontFamily: direct.fontFamily ?? "",
    runStyle: direct.styleId ?? "",
    caps: direct.caps ?? false,
    smallCaps: direct.smallCaps ?? false,
  };
}

/**
 * The direct run formatting a copy takes. With a selection: the formatting most of the selected
 * text has. With a caret: the formatting of the character before it (the one after it at a
 * paragraph's start), as typing there would take. `{}` for an empty paragraph.
 */
export function pickRunFormat(
  runs: InlineSpan[],
  selection: { start: number; length: number } | null,
  caret: number | null,
): RunFormattingInfo {
  const texty = runs.filter((r) => r.span.length > 0);
  if (texty.length === 0) return {};
  if (selection && selection.length > 0) {
    const from = selection.start, to = selection.start + selection.length;
    const weight = new Map<string, { direct: RunFormattingInfo; chars: number }>();
    for (const r of texty) {
      const overlap = Math.min(to, r.span.start + r.span.length) - Math.max(from, r.span.start);
      if (overlap <= 0) continue;
      const key = JSON.stringify(runOpFromDirect(r.direct));
      const entry = weight.get(key) ?? { direct: r.direct, chars: 0 };
      entry.chars += overlap;
      weight.set(key, entry);
    }
    let best: { direct: RunFormattingInfo; chars: number } | null = null;
    for (const entry of weight.values()) if (!best || entry.chars > best.chars) best = entry;
    if (best) return best.direct;
  }
  const at = caret ?? 0;
  const before = texty.find((r) => r.span.start < at && at <= r.span.start + r.span.length);
  return (before ?? texty.find((r) => r.span.start >= at) ?? texty[texty.length - 1]).direct;
}

const WORD_CHAR = /[\p{L}\p{N}\p{M}_'’-]/u;

/** The word around content offset `at` of `text` (a click paints it), or null between words. */
export function wordSpanAt(text: string, at: number): { start: number; length: number } | null {
  let start = Math.max(0, Math.min(at, text.length));
  let end = start;
  while (start > 0 && WORD_CHAR.test(text[start - 1])) start--;
  while (end < text.length && WORD_CHAR.test(text[end])) end++;
  return end > start ? { start, length: end - start } : null;
}

/** The pointer over the document while Format Painter is on: a paint brush. */
const BRUSH_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="24" viewBox="0 0 20 24">'
  + '<path d="M6 1h11v6H6z" fill="#fff" stroke="#000"/><path d="M6 4H3v6h8v3" fill="none" stroke="#000" stroke-width="1.5"/>'
  + '<path d="M9.5 13h3v9h-3z" fill="#000"/></svg>';

export const FORMAT_PAINTER_CSS = `
.docx-format-painting [contenteditable="true"] {
  cursor: url("data:image/svg+xml,${encodeURIComponent(BRUSH_SVG)}") 11 23, text;
}
`;
