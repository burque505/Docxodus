/**
 * Word's Cross-reference dialog in words: what a reference can point to, what it shows ("Insert
 * reference to"), and the REF / PAGEREF switches each choice stands for. Shared by the ribbon's
 * dialog and the `crossReference` command.
 */
import type { CrossReferenceOptions, ReferenceTargetPart } from "./types.js";

/** Reference type: a heading, a bookmark, a numbered paragraph, a caption (of one label: the
 *  dialog lists one type per label, "Figure", "Table"...). */
export type CrossReferenceType = "heading" | "bookmark" | "numbered" | "caption";

/** "Insert reference to": the target's text (a caption's entire caption), its page, its number (in
 *  relative, no or full context), only "above" / "below"; a caption's label and number, or its text. */
export type CrossReferenceInsert = "text" | "page" | "number" | "numberNoContext" | "numberFullContext" | "position"
  | "labelNumber" | "captionText";

export const CROSS_REFERENCE_TYPES: readonly CrossReferenceType[] = ["heading", "bookmark", "numbered", "caption"];

/** The choices of "Insert reference to" for each reference type, in Word's order. */
export const CROSS_REFERENCE_INSERTS: Readonly<Record<CrossReferenceType, readonly CrossReferenceInsert[]>> = {
  heading: ["text", "page", "number", "numberNoContext", "numberFullContext", "position"],
  bookmark: ["text", "page", "number", "numberNoContext", "numberFullContext", "position"],
  numbered: ["page", "number", "numberNoContext", "numberFullContext", "text", "position"],
  caption: ["text", "labelNumber", "captionText", "page", "position"],
};

/** Word's own label of each choice, per reference type (the ribbon translates them). */
export function crossReferenceInsertLabel(type: CrossReferenceType, insert: CrossReferenceInsert): string {
  const noun = type === "heading" ? "Heading" : type === "bookmark" ? "Bookmark" : "Paragraph";
  const number = type === "heading" ? "Heading number" : "Paragraph number";
  switch (insert) {
    case "text": return type === "caption" ? "Entire caption" : `${noun} text`;
    case "labelNumber": return "Only label and number";
    case "captionText": return "Only caption text";
    case "page": return "Page number";
    case "number": return number;
    case "numberNoContext": return `${number} (no context)`;
    case "numberFullContext": return `${number} (full context)`;
    case "position": return "Above/below";
  }
}

/** "Include above/below" applies to numbers and pages only, as in Word. */
export function crossReferenceTakesPosition(insert: CrossReferenceInsert): boolean {
  return insert !== "text" && insert !== "position" && insert !== "captionText";
}

/** Which part of the target paragraph the reference's hidden bookmark spans. */
export function crossReferencePart(insert: CrossReferenceInsert): ReferenceTargetPart {
  return insert === "labelNumber" ? "labelAndNumber" : insert === "captionText" ? "captionText" : "paragraph";
}

/** The engine options for a choice, "Insert as hyperlink" and "Include above/below". */
export function crossReferenceOptionsFor(insert: CrossReferenceInsert, hyperlink: boolean, position: boolean): CrossReferenceOptions {
  const options: CrossReferenceOptions = {};
  if (hyperlink) options.hyperlink = true;
  if (insert === "page") options.pageNumber = true;
  else if (insert === "number") options.referenceNumber = true;
  else if (insert === "numberNoContext") options.numberNoContext = true;
  else if (insert === "numberFullContext") options.numberFullContext = true;
  if (insert === "position" || (position && crossReferenceTakesPosition(insert))) options.includePosition = true;
  return options;
}
