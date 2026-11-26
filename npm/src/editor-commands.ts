/**
 * DocxEditor's command dispatcher: every editing command by name, with plain-object arguments
 * (`editor.exec("fontSize", { pts: 14 })`), and the list of them (`DocxEditor.commands()`).
 *
 * A host that drives the editor from outside the page (a scripting bridge, a command palette, a
 * macro recorder, tests) then needs neither the editor's internals nor a binding per method:
 * arguments and results are JSON-friendly, an unknown command or a wrong argument fails with a
 * typed `EditorCommandError`, and the list says what each command takes and returns.
 *
 * Commands act where the user would: on the selection, or at the caret in the active block (the
 * last block the user or a command was in). Each editing command is one undo step, as the
 * matching ribbon control is.
 */
import type { DocxEditor, EditorAlignment, EditorPageSetup, EditorParagraphFormat, FormatKey, PasteFormatting } from "./editor.js";
import type { FloatingImageLayout, ListFormat, ListLevelSpec, NumberFormat, TableBorderSpec, TableOfContentsOptions, TableOfFiguresOptions, CaptionOptions, StyleDefinitionOp, TablePropertiesOp, TableRowOptions, CellPropertiesOp } from "./types.js";
import { TrackedChangeMode } from "./types.js";
import { FIND_OPTION_KEYS, type FindOptions } from "./find-query.js";
import { CROSS_REFERENCE_INSERTS, CROSS_REFERENCE_TYPES, crossReferenceOptionsFor, crossReferencePart } from "./cross-reference.js";

/** The find options as command arguments (all optional booleans). */
const FIND_OPTION_ARGS = Object.fromEntries(FIND_OPTION_KEYS.map((k) => [`${k}?`, "boolean"])) as Record<string, string>;

/** The find options a command was given; an invalid pattern is a bad argument. */
function findOptionsOf(a: Args, command: string, e: DocxEditor, text: string): FindOptions {
  const options: FindOptions = {};
  for (const key of FIND_OPTION_KEYS) if (a.optBool(key)) options[key] = true;
  const error = e.findError(text, options);
  if (error) throw new EditorCommandError("bad_arguments", command, `${command}: ${error}`);
  return options;
}

/** Why `exec` failed: no such command, an argument missing or of the wrong kind, no document
 *  open, or the command could not act (the engine refused, or there is nothing to act on). */
export type EditorCommandErrorCode = "unknown_command" | "bad_arguments" | "closed" | "refused";

export class EditorCommandError extends Error {
  readonly code: EditorCommandErrorCode;
  readonly command: string;

  constructor(code: EditorCommandErrorCode, command: string, message: string) {
    super(message);
    this.name = "EditorCommandError";
    this.code = code;
    this.command = command;
  }
}

/** One command as `DocxEditor.commands()` lists it. */
export interface EditorCommandInfo {
  /** The name `exec` takes, e.g. `"bold"`. */
  name: string;
  /** Where it belongs, as the ribbon groups its controls: `"font"`, `"paragraph"`, `"insert"`,
   *  `"layout"`, `"table"`, `"picture"`, `"review"`, `"edit"`, `"view"`, `"query"`. A query only reads the
   *  document. */
  group: string;
  /** What it does, in a sentence. */
  description: string;
  /** Its arguments, name to kind (`"string"`, `"number"`, `"boolean"`, one of `"a|b|c"`, or
   *  `"object"`); a name ending in `?` may be left out. */
  args: Readonly<Record<string, string>>;
  /** What `exec` returns, when it returns something. */
  returns?: string;
}

/** The arguments `exec` passes on: whatever the caller gave, read through the checks below. */
export type EditorCommandArgs = Readonly<Record<string, unknown>>;

interface CommandSpec extends Omit<EditorCommandInfo, "name"> {
  run(editor: DocxEditor, args: Args): unknown;
}

/** Argument readers: each throws `bad_arguments` naming the command and the argument. */
class Args {
  constructor(private readonly command: string, private readonly values: EditorCommandArgs) {}

  private fail(name: string, want: string): never {
    throw new EditorCommandError("bad_arguments", this.command,
      `${this.command}: argument "${name}" must be ${want}`);
  }

  private value(name: string): unknown {
    const v = this.values[name];
    return v === null ? undefined : v;
  }

  string(name: string): string {
    const v = this.value(name);
    if (typeof v !== "string") this.fail(name, "a string");
    return v;
  }

  optString(name: string): string | undefined {
    return this.value(name) === undefined ? undefined : this.string(name);
  }

  number(name: string): number {
    const v = this.value(name);
    const n = typeof v === "string" && v.trim() !== "" ? Number(v) : v;
    if (typeof n !== "number" || !Number.isFinite(n)) this.fail(name, "a number");
    return n;
  }

  optNumber(name: string): number | undefined {
    return this.value(name) === undefined ? undefined : this.number(name);
  }

  /** A boolean; AutoHotkey's true and false (1 and 0) count as one. */
  bool(name: string): boolean {
    const v = this.value(name);
    if (typeof v === "boolean") return v;
    if (v === 1 || v === 0) return v === 1;
    return this.fail(name, "true or false");
  }

  optBool(name: string): boolean | undefined {
    return this.value(name) === undefined ? undefined : this.bool(name);
  }

  oneOf<T extends string>(name: string, choices: readonly T[]): T {
    const v = this.value(name);
    if (typeof v !== "string" || !(choices as readonly string[]).includes(v)) this.fail(name, `one of ${choices.join(", ")}`);
    return v as T;
  }

  optOneOf<T extends string>(name: string, choices: readonly T[]): T | undefined {
    return this.value(name) === undefined ? undefined : this.oneOf(name, choices);
  }

  object(name: string): Record<string, unknown> {
    const v = this.value(name);
    if (typeof v !== "object" || v === undefined || Array.isArray(v)) this.fail(name, "an object");
    return v as Record<string, unknown>;
  }

  optObject(name: string): Record<string, unknown> | undefined {
    return this.value(name) === undefined ? undefined : this.object(name);
  }

  /** An array of objects. */
  objects(name: string): Record<string, unknown>[] {
    const v = this.value(name);
    if (!Array.isArray(v) || v.some((x) => typeof x !== "object" || x === null || Array.isArray(x))) this.fail(name, "an array of objects");
    return v as Record<string, unknown>[];
  }
}

/** A level change as setListLevels takes it, read from a plain object (bad fields refused). */
function listLevelSpec(o: Record<string, unknown>, command: string): ListLevelSpec {
  const a = new Args(command, o);
  return {
    level: a.number("level"),
    format: a.optString("format"),
    text: a.optString("text"),
    start: a.optNumber("start"),
    numberPositionTwips: a.optNumber("numberPositionTwips"),
    textPositionTwips: a.optNumber("textPositionTwips"),
    legal: a.optBool("legal"),
  };
}

/** A command that reports "could not act" through a false result. */
function must(ok: boolean, command: string, why: string): void {
  if (!ok) throw new EditorCommandError("refused", command, `${command}: ${why}`);
}

/** The anchor of the heading a command names: by its text (the first heading with it), or the
 *  heading the caret is under when no text is given. */
function headingArg(e: DocxEditor, a: Args, command: string, name = "heading"): string {
  const text = a.optString(name);
  const id = text === undefined
    ? e.headingAtCaret()
    : e.outline().find((h) => h.text.trim() === text.trim())?.anchorId ?? null;
  if (!id) throw new EditorCommandError("refused", command, `${command}: ${text === undefined ? "the caret is under no heading" : `no heading "${text}"`}`);
  return id;
}

const COLUMN_SCOPES = ["section", "document", "forward", "selection"] as const;
const PICTURE_WRAPS = ["inline", "square", "tight", "through", "topAndBottom", "behind", "inFront"] as const;
const ALIGNMENTS: readonly EditorAlignment[] = ["left", "center", "right", "justify"];
const PASTE_FORMATTINGS: readonly PasteFormatting[] = ["merge", "keep", "text"];
const LIST_FORMATS: readonly ListFormat[] = [
  "none", "bullet", "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman",
  "decimalParenthesis", "lowerLetterParenthesis", "upperLetterParenthesis", "lowerRomanParenthesis",
  "upperRomanParenthesis",
];
const SECTION_STARTS = ["nextPage", "continuous", "evenPage", "oddPage", "nextColumn"] as const;
const PAGE_SETUP_TWIPS = [
  "marginTopTwips", "marginBottomTwips", "marginLeftTwips", "marginRightTwips", "pageWidthTwips", "pageHeightTwips",
] as const;
const PARAGRAPH_TWIPS = [
  "leftIndent", "rightIndent", "firstLineIndent", "hangingIndent", "spacingBefore", "spacingAfter", "lineSpacing",
] as const;
const PARAGRAPH_TOGGLES = ["contextualSpacing", "widowControl", "keepNext", "keepLines", "pageBreakBefore"] as const;
const PAGE_NUMBER_FORMATS: readonly NumberFormat[] = ["decimal", "upperLetter", "lowerLetter", "upperRoman", "lowerRoman"];
const FORMATS: ReadonlyArray<[FormatKey, string]> = [
  ["bold", "bold"], ["italic", "italic"], ["underline", "underlined"], ["strike", "struck through"],
  ["superscript", "superscript"], ["subscript", "subscript"], ["code", "code (monospace)"],
];

const formatCommands = Object.fromEntries(FORMATS.map(([key, word]): [string, CommandSpec] => [key, {
  group: "font",
  description: `Make the selection ${word}, or toggle it when "on" is left out (the whole paragraph with no selection).`,
  args: { "on?": "boolean" },
  run: (e, a) => e.format(key, a.optBool("on")),
}]));

const COMMANDS: Readonly<Record<string, CommandSpec>> = {
  // ── font ──
  ...formatCommands,
  fontSize: {
    group: "font", description: "Set the font size, in points.", args: { pts: "number" },
    run: (e, a) => {
      const pts = a.number("pts");
      if (!(pts > 0)) throw new EditorCommandError("bad_arguments", "fontSize", "fontSize: \"pts\" must be more than 0");
      e.setFontSize(pts);
    },
  },
  growFont: { group: "font", description: "Make the font one step larger.", args: {}, run: (e) => e.adjustFontSize(1) },
  shrinkFont: { group: "font", description: "Make the font one step smaller.", args: {}, run: (e) => e.adjustFontSize(-1) },
  fontFamily: { group: "font", description: "Set the font, by name.", args: { name: "string" }, run: (e, a) => e.setFontFamily(a.string("name")) },
  fontColor: {
    group: "font", description: "Set the text color (#RRGGBB).", args: { color: "string" },
    run: (e, a) => e.setFontColor(a.string("color")),
  },
  highlight: {
    group: "font", description: "Highlight the selection in one of Word's highlight colors (\"yellow\", \"green\", …); \"none\" removes it.",
    args: { color: "string" },
    run: (e, a) => { const c = a.string("color"); e.setHighlight(c === "none" ? "" : c); },
  },
  allCaps: { group: "font", description: "Show the selection in capitals (or not).", args: { on: "boolean" }, run: (e, a) => e.setAllCaps(a.bool("on")) },
  smallCaps: { group: "font", description: "Show the selection in small capitals (or not).", args: { on: "boolean" }, run: (e, a) => e.setSmallCaps(a.bool("on")) },
  clearFormatting: { group: "font", description: "Remove the selection's direct formatting.", args: {}, run: (e) => e.clearFormatting() },
  copyFormatting: {
    group: "font", description: "Copy the formatting at the caret or of the selection (Word's Ctrl+Shift+C); with a caret, or the paragraph's end selected, the paragraph's too.",
    args: {}, returns: "boolean", run: (e) => e.copyFormatting(),
  },
  pasteFormatting: {
    group: "font", description: "Give the selection (with a caret, the word there) the copied formatting (Word's Ctrl+Shift+V).",
    args: {}, returns: "boolean", run: (e) => e.pasteFormatting(),
  },
  formatPainter: {
    group: "font", description: "Format Painter: copy the formatting here and paint it on the next click or selection; \"sticky\" keeps painting until Esc. \"off\" turns it off.",
    args: { "mode?": "once|sticky|off" }, returns: "boolean",
    run: (e, a) => {
      const mode = a.optOneOf("mode", ["once", "sticky", "off"] as const) ?? "once";
      if (mode === "off") { e.stopFormatPainter(); return true; }
      return e.formatPainter(mode === "sticky");
    },
  },

  changeCase: {
    group: "font", description: "Change Case on the selection (with a caret, the word there), keeping its formatting: sentence, lower, upper, title (Capitalize Each Word), toggle, headline (Headline Case), or cycle (Word's Shift+F3).",
    args: { mode: "sentence|lower|upper|title|toggle|headline|cycle" }, returns: "boolean",
    run: (e, a) => e.changeCase(a.oneOf("mode", ["sentence", "lower", "upper", "title", "toggle", "headline", "cycle"] as const)),
  },

  // ── paragraph ──
  showMarks: {
    group: "paragraph", description: "Show/Hide ¶: show paragraph marks and tabs (on), hide them (off), or switch (no argument), as Word's Ctrl+Shift+8.",
    args: { "on?": "boolean" }, returns: "boolean",
    run: (e, a) => { const on = a.optBool("on") ?? !e.showFormattingMarks; e.setShowFormattingMarks(on); return on; },
  },
  joinParagraphs: {
    group: "paragraph", description: "Join the paragraphs the selection touches into one, keeping the last paragraph mark (mends text copied from a PDF: a space where two lines meet, none after a hyphen). One undo step.",
    args: {}, returns: "boolean", run: (e) => e.joinParagraphs(),
  },
  align: {
    group: "paragraph", description: "Align the paragraph.", args: { alignment: ALIGNMENTS.join("|") },
    run: (e, a) => e.setAlignment(a.oneOf("alignment", ALIGNMENTS)),
  },
  style: {
    group: "paragraph", description: "Give the paragraph a style, by its id (\"Heading1\", \"Normal\", …; see the \"styles\" query).",
    args: { styleId: "string" }, run: (e, a) => e.setParagraphStyle(a.string("styleId")),
  },
  characterStyle: {
    group: "font", description: "Give the selection a character style, by its id (\"Strong\", \"Emphasis\", …; see the \"styles\" query).",
    args: { styleId: "string" }, run: (e, a) => e.applyRunFormat({ runStyle: a.string("styleId") }),
  },
  fontFormat: {
    group: "font",
    description: "Word's Font dialog: set any of these on the selection (the whole paragraph with none); what is left out stays. size in points, color #RRGGBB, vertAlign superscript|subscript|baseline.",
    args: {
      "font?": "string", "size?": "number", "bold?": "boolean", "italic?": "boolean", "underline?": "boolean",
      "strike?": "boolean", "color?": "string", "smallCaps?": "boolean", "caps?": "boolean", "vertAlign?": "superscript|subscript|baseline",
    },
    run: (e, a) => {
      const op: Record<string, unknown> = {};
      const font = a.optString("font"); if (font !== undefined) op.fontFamily = font;
      const size = a.optNumber("size");
      if (size !== undefined) {
        if (!(size >= 1 && size <= 1638)) throw new EditorCommandError("bad_arguments", "fontFormat", "fontFormat: \"size\" must be 1 to 1638");
        op.fontSizePts = size;
      }
      for (const key of ["bold", "italic", "underline", "strike", "smallCaps", "caps"] as const) {
        const v = a.optBool(key); if (v !== undefined) op[key] = v;
      }
      const color = a.optString("color"); if (color !== undefined) op.color = color.replace(/^#/, "").toUpperCase();
      const vert = a.optOneOf("vertAlign", ["superscript", "subscript", "baseline"] as const); if (vert !== undefined) op.vertAlign = vert;
      if (Object.keys(op).length === 0) throw new EditorCommandError("bad_arguments", "fontFormat", "fontFormat: give at least one setting");
      e.applyRunFormat(op);
    },
  },
  setStyle: {
    group: "paragraph",
    description: "Word's Modify Style / Create New Style: create a paragraph or character style (no styleId; name required) or change one. paragraph takes paragraphFormat's settings, run fontFormat's engine names (fontFamily, fontSizePts, bold, …). Every paragraph in the style changes.",
    args: {
      "styleId?": "string", "name?": "string", "type?": "paragraph|character", "basedOn?": "string", "next?": "string",
      "quickFormat?": "boolean", "paragraph?": "object", "run?": "object",
    },
    run: (e, a) => {
      const op: Record<string, unknown> = {};
      for (const key of ["styleId", "name", "basedOn", "next"] as const) {
        const v = a.optString(key); if (v !== undefined) op[key] = v;
      }
      const type = a.optOneOf("type", ["paragraph", "character"] as const); if (type !== undefined) op.type = type;
      const quick = a.optBool("quickFormat"); if (quick !== undefined) op.quickFormat = quick;
      const paragraph = a.optObject("paragraph"); if (paragraph) op.paragraph = paragraph;
      const run = a.optObject("run"); if (run) op.run = run;
      if (op.styleId === undefined && op.name === undefined) throw new EditorCommandError("bad_arguments", "setStyle", "setStyle: give styleId to change a style, or name to create one");
      const id = e.setStyle(op as StyleDefinitionOp);
      must(id !== null, "setStyle", e.lastStyleError?.message ?? "the engine refused it");
      return id;
    },
  },
  deleteStyle: {
    group: "paragraph",
    description: "Word's Styles pane Delete: delete a custom paragraph or character style. What used it takes the style it was based on; a paragraph style's linked character style goes too. Built-in and default styles are refused.",
    args: { styleId: "string" },
    run: (e, a) => {
      must(e.deleteStyle(a.string("styleId")), "deleteStyle", e.lastStyleError?.message ?? "the engine refused it");
    },
  },
  indent: {
    group: "paragraph", description: "Change the left indent by this many twips (1440 an inch; negative to decrease).",
    args: { "twips?": "number" }, run: (e, a) => e.indent(a.optNumber("twips") ?? 720),
  },
  firstLineIndent: { group: "paragraph", description: "Indent the first line, in twips.", args: { twips: "number" }, run: (e, a) => e.setFirstLineIndent(a.number("twips")) },
  hangingIndent: { group: "paragraph", description: "Give the paragraph a hanging indent, in twips.", args: { twips: "number" }, run: (e, a) => e.setHangingIndent(a.number("twips")) },
  lineSpacing: { group: "paragraph", description: "Set line spacing as a multiple (1, 1.15, 1.5, 2, …).", args: { multiple: "number" }, run: (e, a) => e.setLineSpacing(a.number("multiple")) },
  paragraphSpacing: {
    group: "paragraph", description: "Set the space before and after the paragraph, in points.",
    args: { "before?": "number", "after?": "number" },
    run: (e, a) => e.setParagraphSpacing({ beforePt: a.optNumber("before"), afterPt: a.optNumber("after") }),
  },
  paragraphFormat: {
    group: "paragraph",
    description: "Word's Paragraph dialog: set any of these on the paragraph or the selected paragraphs, in twips (1440 an inch, 20 a point); what is left out stays. lineSpacing is 240ths of a line under \"auto\", twips under \"exact\" and \"atLeast\".",
    args: {
      "alignment?": "left|center|right|justify",
      "leftIndent?": "number", "rightIndent?": "number", "firstLineIndent?": "number", "hangingIndent?": "number",
      "spacingBefore?": "number", "spacingAfter?": "number", "lineSpacing?": "number", "lineSpacingRule?": "auto|exact|atLeast",
      "contextualSpacing?": "boolean", "widowControl?": "boolean", "keepNext?": "boolean", "keepLines?": "boolean",
      "pageBreakBefore?": "boolean",
    },
    run: (e, a) => {
      const op: Record<string, unknown> = {};
      const alignment = a.optOneOf("alignment", ["left", "center", "right", "justify"] as const);
      if (alignment !== undefined) op.alignment = alignment;
      for (const key of PARAGRAPH_TWIPS) {
        const v = a.optNumber(key);
        if (v !== undefined) op[key] = Math.round(v);
      }
      const rule = a.optOneOf("lineSpacingRule", ["auto", "exact", "atLeast"] as const);
      if (rule !== undefined) {
        if (op.lineSpacing === undefined) throw new EditorCommandError("bad_arguments", "paragraphFormat", "paragraphFormat: lineSpacingRule needs lineSpacing");
        op.lineSpacingRule = rule;
      }
      if (op.firstLineIndent !== undefined && op.hangingIndent !== undefined) {
        throw new EditorCommandError("bad_arguments", "paragraphFormat", "paragraphFormat: give firstLineIndent or hangingIndent, not both");
      }
      for (const key of PARAGRAPH_TOGGLES) {
        const v = a.optBool(key);
        if (v !== undefined) op[key] = v;
      }
      if (Object.keys(op).length === 0) throw new EditorCommandError("bad_arguments", "paragraphFormat", "paragraphFormat: give at least one setting");
      must(e.setParagraphFormat(op as EditorParagraphFormat), "paragraphFormat", "there is no paragraph to format");
    },
  },
  pageBreakBefore: {
    group: "paragraph", description: "Start the paragraph on a new page (or not).", args: { "on?": "boolean" },
    run: (e, a) => e.pageBreakBefore(a.optBool("on") ?? true),
  },
  list: {
    group: "paragraph", description: "Toggle a bulleted or numbered list.", args: { kind: "bullet|decimal" },
    run: (e, a) => e.toggleList(a.oneOf("kind", ["bullet", "decimal"] as const)),
  },
  listFormat: {
    group: "paragraph", description: "Set the list's number format (\"none\" ends the list).", args: { format: LIST_FORMATS.join("|") },
    run: (e, a) => e.setListFormat(a.oneOf("format", LIST_FORMATS)),
  },
  legalNumbering: {
    group: "paragraph",
    description: "Legal numbering, 1. / 1.1 / 1.1.1 by level, on the list (the paragraphs made a numbered list first): a period after the top level's number (\"top\"), every level's (\"all\") or none.",
    args: { "periods?": "top|all|none" },
    run: (e, a) => must(e.setLegalNumbering(a.optOneOf("periods", ["top", "all", "none"] as const) ?? "top"), "legalNumbering", "there is no paragraph to number"),
  },
  tabStops: {
    group: "paragraph",
    description: "Word's Tabs: make these the paragraph's tab stops (every selected paragraph's). Each object has positionTwips (from the left margin, not the indent), alignment left|center|right|decimal|bar, leader none|dot|hyphen|underscore|heavy|middleDot. [] clears them all.",
    args: { stops: "object" },
    run: (e, a) => must(e.setTabStops(a.objects("stops").map((o) => {
      const s = new Args("tabStops", o);
      return {
        positionTwips: Math.round(s.number("positionTwips")),
        alignment: s.optOneOf("alignment", ["left", "center", "right", "decimal", "bar"] as const) ?? "left",
        leader: s.optOneOf("leader", ["none", "dot", "hyphen", "underscore", "heavy", "middleDot"] as const) ?? "none",
      };
    })), "tabStops", "the tab stops could not be set (two at one position, one past 22 inches, or no paragraph)"),
  },
  defaultTabStop: {
    group: "paragraph", description: "The document's default tab interval, in twips (720 = half an inch).",
    args: { twips: "number" },
    run: (e, a) => must(e.setDefaultTabStop(Math.round(a.number("twips"))), "defaultTabStop", "the interval must be greater than 0"),
  },
  setListLevels: {
    group: "paragraph",
    description: "Word's Define New Multilevel List: change levels of the list (the paragraphs made a numbered list first). Each object names a level (0-8) and any of format, text (\"%1.%2\"), start, numberPositionTwips, textPositionTwips, legal.",
    args: { levels: "object" },
    run: (e, a) => must(e.setListLevels(a.objects("levels").map((o) => listLevelSpec(o, "setListLevels"))), "setListLevels",
      "the levels could not be set (a level outside 0-8, an unknown format, or a number text naming a deeper level)"),
  },
  restartNumbering: {
    group: "paragraph", description: "Restart the list's numbering here, at start (1 when left out).", args: { "start?": "number" },
    run: (e, a) => must(e.restartNumbering(a.optNumber("start") ?? 1), "restartNumbering", "the caret is not in a list item"),
  },
  continueNumbering: {
    group: "paragraph", description: "Continue the numbering of the list before this one.", args: {},
    run: (e) => must(e.continueNumbering(), "continueNumbering", "the caret is not in a list item, or no list comes before it"),
  },
  clearBorders: { group: "paragraph", description: "Remove the paragraph's borders.", args: {}, run: (e) => e.clearParagraphBorders() },

  // ── insert ──
  insertText: {
    group: "insert",
    description: "Type text at the caret, replacing the selection; each line break starts a new paragraph.",
    args: { text: "string" },
    run: (e, a) => must(e.insertText(a.string("text")), "insertText", "there is no caret in the document, or the selection spans paragraphs"),
  },
  newParagraph: {
    group: "insert", description: "Start a new paragraph at the caret (Enter).", args: {},
    run: (e) => must(e.insertText("\n"), "newParagraph", "there is no caret in the document"),
  },
  horizontalRule: {
    group: "insert", description: "Draw a horizontal rule above or below the paragraph.",
    args: { "weight?": "number", "style?": "string", "position?": "above|below" },
    run: (e, a) => e.insertHorizontalRule(a.optNumber("weight") ?? 12, a.optString("style") ?? "single", a.optOneOf("position", ["above", "below"] as const) ?? "below"),
  },
  footnote: { group: "insert", description: "Insert a footnote at the caret.", args: { "text?": "string" }, run: (e, a) => e.insertFootnote(a.optString("text")) },
  endnote: { group: "insert", description: "Insert an endnote at the caret.", args: { "text?": "string" }, run: (e, a) => e.insertEndnote(a.optString("text")) },
  pageNumber: {
    group: "insert", description: "Insert a page-number field (in a header or footer, or at the footer).",
    args: { "field?": "currentPage|totalPages|pageOfTotal" },
    run: (e, a) => e.insertPageNumber(a.optOneOf("field", ["currentPage", "totalPages", "pageOfTotal"] as const) ?? "currentPage"),
  },
  updateFields: {
    group: "insert", description: "Word's Update Field: the fields in the selection's paragraphs (F9), every field (scope all), or the tables of contents (scope tables, Update Table).",
    args: { "scope?": "selection|all|tables" },
    run: (e, a) => must(e.updateFields(a.optOneOf("scope", ["selection", "all", "tables"] as const) ?? "selection"), "updateFields", "the engine refused it (changes are tracked?)"),
  },
  tableOfContents: {
    group: "insert", description: "Insert a table of contents (options as TableOfContentsOptions).", args: { "options?": "object" },
    run: (e, a) => must(e.insertTableOfContents((a.optObject("options") ?? {}) as TableOfContentsOptions), "tableOfContents", "the engine refused it"),
  },
  tableOfFigures: {
    group: "insert", description: "Insert a table of figures for one caption label, filled at once (options as TableOfFiguresOptions).", args: { "options?": "object" },
    run: (e, a) => must(e.insertTableOfFigures((a.optObject("options") ?? {}) as TableOfFiguresOptions), "tableOfFigures", "the engine refused it"),
  },
  insertCaption: {
    group: "insert", description: "Word's Insert Caption on the caret's table (above) or paragraph (below): a label (Figure, Table, Equation or your own), its number and `text` after it.",
    args: { "label?": "string", "position?": "above|below", "excludeLabel?": "boolean", "numberFormat?": "ARABIC|alphabetic|ALPHABETIC|roman|ROMAN", "text?": "string" },
    run: (e, a) => {
      const options: CaptionOptions = {};
      const label = a.optString("label");
      if (label !== undefined) options.label = label;
      const position = a.optOneOf("position", ["above", "below"] as const);
      if (position) options.position = position;
      if (a.optBool("excludeLabel")) options.excludeLabel = true;
      const format = a.optOneOf("numberFormat", ["ARABIC", "alphabetic", "ALPHABETIC", "roman", "ROMAN"] as const);
      if (format) options.numberFormat = format;
      const text = a.optString("text");
      if (text) options.text = text;
      must(e.insertCaption(options), "insertCaption", "there is nothing to caption at the caret, or the engine refused it (changes are tracked?)");
    },
  },
  hyperlink: {
    group: "insert", description: "Link the selection to a web address, or to a bookmark (kind \"internal\").",
    args: { target: "string", "kind?": "external|internal" },
    run: (e, a) => must(e.insertHyperlink(a.string("target"), a.optOneOf("kind", ["external", "internal"] as const) ?? "external"), "hyperlink", "select the text to link first"),
  },
  removeHyperlink: { group: "insert", description: "Remove the link at the caret, keeping its text.", args: {}, run: (e) => { e.removeHyperlink(); } },
  addBookmark: {
    group: "insert", description: "Put a bookmark on the selection (or at the caret); a name the document already has moves that bookmark here, as Word's Bookmark > Add does.",
    args: { name: "string" },
    run: (e, a) => {
      const r = e.addBookmark(a.string("name"));
      must(r.ok, "addBookmark", r.code ? `the engine refused it (${r.code})` : "there is no caret in the document");
    },
  },
  deleteBookmark: {
    group: "insert", description: "Delete a bookmark (its text stays). Refused while a link or cross-reference uses it.",
    args: { name: "string" },
    run: (e, a) => {
      const r = e.removeBookmark(a.string("name"));
      must(r.ok, "deleteBookmark", r.code ? `the engine refused it (${r.code})` : "the engine refused it");
    },
  },
  goToHeading: {
    group: "insert", description: "Put the caret at a heading (by its text), as a click in the Navigation Pane does.",
    args: { heading: "string" },
    run: (e, a) => must(e.goToHeading(headingArg(e, a, "goToHeading")), "goToHeading", "the heading is not shown"),
  },
  moveHeading: {
    group: "insert", description: "Move a heading (by its text; the caret's when absent) with everything under it before or after another heading's part, as a drag in the Navigation Pane does.",
    args: { "heading?": "string", to: "string", "position?": "before|after" },
    run: (e, a) => {
      const id = headingArg(e, a, "moveHeading");
      const to = e.outline().find((h) => h.anchorId === headingArg(e, a, "moveHeading", "to"))!;
      const after = a.optOneOf("position", ["before", "after"] as const) === "after";
      const r = e.moveHeading(id, after ? to.lastAnchorId : to.anchorId, after ? "after" : "before");
      must(r.ok, "moveHeading", r.code ? `the engine refused it (${r.code})` : "the engine refused it");
    },
  },
  headingLevel: {
    group: "insert", description: "Promote or demote a heading (by its text; the caret's when absent) to outline level 1-9, through the document's \"heading N\" style.",
    args: { "heading?": "string", level: "number" },
    run: (e, a) => {
      const r = e.setHeadingLevel(headingArg(e, a, "headingLevel"), a.number("level"));
      must(r.ok, "headingLevel", r.code ? `the engine refused it (${r.code})` : "the level is 1-9");
    },
  },
  newHeading: {
    group: "insert", description: "Word's New Heading Before / After / New Subheading for a heading (by its text; the caret's when absent).",
    args: { "heading?": "string", where: "before|after|sub" },
    run: (e, a) => must(e.insertHeading(headingArg(e, a, "newHeading"), a.oneOf("where", ["before", "after", "sub"] as const)), "newHeading", "the engine refused it"),
  },
  deleteHeading: {
    group: "insert", description: "Delete a heading (by its text; the caret's when absent) and everything under it.",
    args: { "heading?": "string" },
    run: (e, a) => must(e.deleteHeadingSection(headingArg(e, a, "deleteHeading")), "deleteHeading", "the engine refused it"),
  },
  selectHeading: {
    group: "insert", description: "Word's Select Heading and Content for a heading (by its text; the caret's when absent).",
    args: { "heading?": "string" },
    run: (e, a) => must(e.selectHeadingSection(headingArg(e, a, "selectHeading")), "selectHeading", "the heading is not shown"),
  },
  crossReference: {
    group: "insert", description: "Word's Insert > Cross-reference at the caret (a selection is replaced): to a heading (by its text), a bookmark (by its name), a numbered item (by its text, or its number and text) or a caption (by its text; `label` is its label); `insert` is what it shows.",
    args: { type: CROSS_REFERENCE_TYPES.join("|"), target: "string", insert: "text|page|number|numberNoContext|numberFullContext|position|labelNumber|captionText", "label?": "string", "hyperlink?": "boolean", "position?": "boolean" },
    run: (e, a) => {
      const type = a.oneOf("type", CROSS_REFERENCE_TYPES);
      const insert = a.oneOf("insert", CROSS_REFERENCE_INSERTS[type]);
      const wanted = a.string("target").trim();
      let target: { bookmark: string } | { anchorId: string; part?: ReturnType<typeof crossReferencePart> } | null = null;
      if (type === "caption") {
        const id = (a.optString("label") ?? "").trim().replace(/ /g, "_").toLowerCase();
        const found = e.captions().find((c) => (!id || c.label.toLowerCase() === id) && c.text.trim() === wanted);
        if (found) target = { anchorId: found.anchorId, part: crossReferencePart(insert) };
      } else if (type === "bookmark") {
        if (e.bookmarks(true).some((b) => b.name === wanted)) target = { bookmark: wanted };
      } else if (type === "heading") {
        const id = e.outline().find((h) => h.text.trim() === wanted)?.anchorId;
        if (id) target = { anchorId: id };
      } else {
        const id = e.numberedItems().find((n) => n.text.trim() === wanted || `${n.number} ${n.text}`.trim() === wanted)?.anchorId;
        if (id) target = { anchorId: id };
      }
      if (!target) throw new EditorCommandError("refused", "crossReference", `crossReference: no ${type} "${wanted}"`);
      const r = e.insertCrossReference(null, target, crossReferenceOptionsFor(insert, !!a.optBool("hyperlink"), !!a.optBool("position")));
      must(r.ok, "crossReference", r.code ? `the engine refused it (${r.code})` : "there is no caret in the document body");
    },
  },
  goToBookmark: {
    group: "insert", description: "Select a bookmark's text (the caret, for an empty one) and scroll to it.",
    args: { name: "string" },
    run: (e, a) => must(e.goToBookmark(a.string("name")), "goToBookmark", "no such bookmark in the body, a header or a footer"),
  },
  table: {
    group: "insert", description: "Insert a table at the caret: rows by columns, optionally borderless, its cells aligned.",
    args: { rows: "number", columns: "number", "align?": ALIGNMENTS.join("|"), "borderless?": "boolean" },
    run: (e, a) => {
      const rows = a.number("rows"), columns = a.number("columns");
      if (!(rows >= 1 && columns >= 1)) throw new EditorCommandError("bad_arguments", "table", "table: \"rows\" and \"columns\" must be at least 1");
      e.insertTable(Math.trunc(rows), Math.trunc(columns), {
        borderless: a.optBool("borderless") ?? false,
        cellAlignment: a.optOneOf("align", ALIGNMENTS),
      });
    },
  },
  comment: {
    group: "insert", description: "Comment on the selection (or the paragraph).", args: { text: "string" }, returns: "the comment's id",
    run: (e, a) => {
      const entry = e.addComment(a.string("text"));
      must(!!entry, "comment", "there is no caret in the document");
      return entry!.anchorId;
    },
  },

  // ── layout ──
  lineNumbers: {
    group: "layout",
    description: "Word's Line Numbers for the caret's section (applyTo \"document\": every section): countBy (0 turns them off), start, distanceTwips (absent = Auto), restart newPage|newSection|continuous. They show in page view.",
    args: { countBy: "number", "start?": "number", "distanceTwips?": "number", "restart?": "newPage|newSection|continuous", "applyTo?": "section|document" },
    run: (e, a) => e.setLineNumbering({
      countBy: Math.round(a.number("countBy")),
      ...(a.optNumber("start") !== undefined ? { start: Math.round(a.optNumber("start")!) } : {}),
      ...(a.optNumber("distanceTwips") !== undefined ? { distanceTwips: Math.round(a.optNumber("distanceTwips")!) } : {}),
      ...(a.optOneOf("restart", ["newPage", "newSection", "continuous"] as const) ? { restart: a.optOneOf("restart", ["newPage", "newSection", "continuous"] as const) } : {}),
    }, a.optOneOf("applyTo", ["section", "document"] as const) ?? "section"),
  },
  suppressLineNumbers: {
    group: "layout", description: "Word's Suppress for Current Paragraph (or on: false to number it again).",
    args: { "on?": "boolean" }, run: (e, a) => e.setSuppressLineNumbers(a.optBool("on") ?? true),
  },
  columns: {
    group: "layout",
    description: "Set the text columns: how many (equal), the space between in twips, a line between; applyTo is Word's Apply to (selection when text is selected, else section).",
    args: { "count?": "number", "spaceTwips?": "number", "separator?": "boolean", "applyTo?": COLUMN_SCOPES.join("|") },
    run: (e, a) => {
      const op: { count?: number; spaceTwips?: number; separator?: boolean } = {};
      const count = a.optNumber("count");
      if (count !== undefined) op.count = Math.round(count);
      const space = a.optNumber("spaceTwips");
      if (space !== undefined) op.spaceTwips = Math.round(space);
      const separator = a.optBool("separator");
      if (separator !== undefined) op.separator = separator;
      must(e.setColumns(op, a.optOneOf("applyTo", COLUMN_SCOPES)), "columns",
        "columns go in the body, each at least half an inch wide; a new section is refused while changes are tracked");
    },
  },
  pageSetup: {
    group: "layout",
    description: "Set the section's page: margins and size in twips (1440 an inch), landscape or not; what is left out stays.",
    args: {
      "marginTopTwips?": "number", "marginBottomTwips?": "number", "marginLeftTwips?": "number", "marginRightTwips?": "number",
      "pageWidthTwips?": "number", "pageHeightTwips?": "number", "landscape?": "boolean",
      "sectionStart?": SECTION_STARTS.join("|"),
    },
    run: (e, a) => {
      const op: Record<string, number | boolean | string> = {};
      for (const key of PAGE_SETUP_TWIPS) {
        const v = a.optNumber(key);
        if (v !== undefined) op[key] = Math.round(v);
      }
      const landscape = a.optBool("landscape");
      if (landscape !== undefined) op.landscape = landscape;
      const sectionStart = a.optOneOf("sectionStart", SECTION_STARTS);
      if (sectionStart !== undefined) op.sectionStart = sectionStart;
      if (Object.keys(op).length === 0) throw new EditorCommandError("bad_arguments", "pageSetup", "pageSetup: give at least one setting");
      must(e.setPageSetup(op as EditorPageSetup), "pageSetup", "this engine build cannot change the page setup");
    },
  },
  insertBreak: {
    group: "layout", description: "Insert a page break (Ctrl+Enter) or a column break (Ctrl+Shift+Enter) at the caret.",
    args: { "kind?": "page|column" },
    run: (e, a) => must(e.insertBreak(a.optOneOf("kind", ["page", "column"] as const) ?? "page"), "insertBreak",
      "a break goes in a body paragraph, not in a table, header or footer, nor while changes are tracked"),
  },
  sectionBreak: {
    group: "layout", description: "End the section at the caret; the new one starts on the next page, continuously, or on an even or odd page.",
    args: { "start?": SECTION_STARTS.join("|") },
    run: (e, a) => must(e.insertSectionBreak(a.optOneOf("start", SECTION_STARTS) ?? "nextPage"), "sectionBreak",
      "a section break goes in a body paragraph, not in a table, header or footer, nor while changes are tracked"),
  },
  removeSectionBreak: {
    group: "layout", description: "Remove the section break that ends the caret's section (in the last section, the one before it); the text before it takes the layout of the section after it.",
    args: {}, run: (e) => must(e.removeSectionBreak(), "removeSectionBreak", "the document is one section: there is no section break to remove"),
  },
  pageNumbering: {
    group: "layout", description: "Number the section's pages in a format, or restart them at a value.",
    args: { "format?": PAGE_NUMBER_FORMATS.join("|"), "start?": "number" },
    run: (e, a) => {
      const format = a.optOneOf("format", PAGE_NUMBER_FORMATS), start = a.optNumber("start");
      if (format === undefined && start === undefined) throw new EditorCommandError("bad_arguments", "pageNumbering", "pageNumbering: give a format or a start");
      e.setPageNumbering({ ...(format ? { format } : {}), ...(start !== undefined ? { start: Math.trunc(start) } : {}) });
    },
  },
  clearPageNumbering: { group: "layout", description: "Continue the previous section's page numbering.", args: {}, run: (e) => e.clearPageNumbering() },
  goToHeaderFooter: {
    group: "layout", description: "Put the caret in the section's header or footer.", args: { which: "header|footer" },
    run: (e, a) => must(e.goToHeaderFooter(a.oneOf("which", ["header", "footer"] as const)), "goToHeaderFooter", "this view has no header or footer to go to"),
  },
  closeHeaderFooter: { group: "layout", description: "Leave the header or footer, back to the body.", args: {}, run: (e) => e.closeHeaderFooter() },
  headerFooterOption: {
    group: "layout", description: "A different header and footer on the first page, or on odd and even pages (or not).",
    args: { kind: "first|even", on: "boolean" },
    run: (e, a) => must(e.setHeaderFooterKindEnabled(a.oneOf("kind", ["first", "even"] as const), a.bool("on")), "headerFooterOption", "this engine build cannot change that option"),
  },

  // ── table ──
  insertRow: { group: "table", description: "Insert a table row.", args: { where: "above|below" }, run: (e, a) => e.insertTableRow(a.oneOf("where", ["above", "below"] as const)) },
  insertColumn: { group: "table", description: "Insert a table column.", args: { where: "left|right" }, run: (e, a) => e.insertTableColumn(a.oneOf("where", ["left", "right"] as const)) },
  deleteRow: { group: "table", description: "Delete the table row at the caret.", args: {}, run: (e) => e.deleteTableRow() },
  deleteColumn: { group: "table", description: "Delete the table column at the caret.", args: {}, run: (e) => e.deleteTableColumn() },
  mergeCells: {
    group: "table", description: "Merge the cell at the caret with its neighbors (rows down, columns right).",
    args: { rows: "number", columns: "number" }, run: (e, a) => e.mergeCells(a.number("rows"), a.number("columns")),
  },
  unmergeCells: { group: "table", description: "Split a merged cell back up.", args: {}, run: (e) => e.unmergeCells() },
  tableBorders: {
    group: "table", description: "Set the table's borders (spec as TableBorderSpec: scope, style, size, color).", args: { "spec?": "object" },
    run: (e, a) => e.setTableBorders((a.optObject("spec") ?? { scope: "all", style: "single", size: 4 }) as TableBorderSpec),
  },
  cellShading: {
    group: "table", description: "Shade the cell, row, column or table (#RRGGBB; \"\" removes it).",
    args: { fill: "string", "scope?": "cell|row|column|table" },
    run: (e, a) => e.setCellShading(a.string("fill"), a.optOneOf("scope", ["cell", "row", "column", "table"] as const) ?? "cell"),
  },
  repeatHeaderRow: { group: "table", description: "Repeat the first row on every page (or not).", args: { on: "boolean" }, run: (e, a) => e.setRepeatHeaderRow(a.bool("on")) },
  deleteTable: { group: "table", description: "Delete the table at the caret.", args: {}, run: (e) => e.deleteTable() },
  selectCells: {
    group: "table", description: "Select the cell, row, column or whole table at the caret (or around the cells selected), as whole cells: Word's Table Layout > Select.",
    args: { part: "cell|row|column|table" },
    run: (e, a) => must(e.selectTablePart(a.oneOf("part", ["cell", "row", "column", "table"] as const)), "selectCells", "the caret is not in a table"),
  },
  tableProperties: {
    group: "table", description: "Word's Table Properties for the table at the caret (a TablePropertiesOp: widthType auto|twips|percent, width, alignment left|center|right, indentTwips, cellMargins, altTitle, altDescription).",
    args: { op: "object" }, run: (e, a) => e.setTableProperties(a.object("op") as TablePropertiesOp),
  },
  applyTableProperties: {
    group: "table", description: "The Table Properties dialog's OK: any of table (a TablePropertiesOp), row (height, rule, allowBreak, repeatHeader), columnWidthTwips, cell (verticalAlignment), as one undo step.",
    args: { "table?": "object", "row?": "object", "columnWidthTwips?": "number", "cell?": "object" },
    run: (e, a) => e.applyTableProperties({
      table: a.optObject("table") as TablePropertiesOp | undefined,
      row: a.optObject("row") as TableRowOptions | undefined,
      columnWidthTwips: a.optNumber("columnWidthTwips"),
      cell: a.optObject("cell") as CellPropertiesOp | undefined,
    }),
  },
  cellAlignment: {
    group: "table", description: "The vertical alignment of the cell at the caret (or its row, column or table).",
    args: { align: "top|center|bottom", "scope?": "cell|row|column|table" },
    run: (e, a) => e.setCellProperties({ verticalAlignment: a.oneOf("align", ["top", "center", "bottom"] as const) },
      a.optOneOf("scope", ["cell", "row", "column", "table"] as const) ?? "cell"),
  },
  rowHeight: {
    group: "table", description: "The height of the row at the caret, twips (0 removes it), at least or exactly.",
    args: { twips: "number", "rule?": "atLeast|exact" },
    run: (e, a) => e.setTableRowOptions({ heightTwips: a.number("twips"), heightRule: a.optOneOf("rule", ["atLeast", "exact"] as const) ?? "atLeast" }),
  },
  rowOptions: {
    group: "table", description: "The row at the caret: allow it to break across pages, repeat it as a header row.",
    args: { "allowBreak?": "boolean", "repeatHeader?": "boolean" },
    run: (e, a) => e.setTableRowOptions({ allowBreakAcrossPages: a.optBool("allowBreak"), repeatHeader: a.optBool("repeatHeader") }),
  },
  columnWidth: {
    group: "table", description: "The width of the column at the caret, twips.",
    args: { twips: "number" }, run: (e, a) => e.setColumnWidth(a.number("twips")),
  },

  // ── review ──
  trackChanges: {
    group: "review", description: "Record edits as tracked changes (or stop).", args: { on: "boolean" },
    run: (e, a) => e.setTrackedChanges(a.bool("on") ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept),
  },
  acceptAllChanges: { group: "review", description: "Accept every tracked change.", args: {}, run: (e) => must(e.acceptAllRevisions(), "acceptAllChanges", "the engine refused it") },
  rejectAllChanges: { group: "review", description: "Reject every tracked change.", args: {}, run: (e) => must(e.rejectAllRevisions(), "rejectAllChanges", "the engine refused it") },
  showComments: { group: "review", description: "Show or hide the comments.", args: { on: "boolean" }, run: (e, a) => e.showComments(a.bool("on")) },
  showMarkup: {
    group: "review", description: "Show the tracked changes (or the document as if they were accepted; edits keep being tracked).",
    args: { on: "boolean" }, run: (e, a) => e.showRevisions(a.bool("on")),
  },

  // ── edit ──
  undo: { group: "edit", description: "Undo the last edit.", args: {}, run: (e) => e.undo() },
  redo: { group: "edit", description: "Redo the last undone edit.", args: {}, run: (e) => e.redo() },
  paste: {
    group: "edit",
    description: "Paste the clipboard at the caret, replacing the selection: \"merge\" (the default) takes the document's fonts, \"keep\" the source's, \"text\" the text only.",
    args: { "formatting?": "merge|keep|text" }, returns: "Promise<boolean>",
    run: (e, a) => e.pasteFromClipboard(a.optOneOf("formatting", PASTE_FORMATTINGS)),
  },
  pasteHtml: {
    group: "edit", description: "Paste HTML at the caret, as pasting it from the clipboard would.",
    args: { html: "string", "formatting?": "merge|keep|text" }, returns: "Promise<boolean>",
    run: (e, a) => e.pasteHtml(a.string("html"), a.optOneOf("formatting", PASTE_FORMATTINGS)),
  },
  pasteText: {
    group: "edit", description: "Paste plain text at the caret, taking the formatting where it lands; each line a paragraph.",
    args: { text: "string" }, returns: "Promise<boolean>",
    run: (e, a) => e.pasteText(a.string("text")),
  },
  find: {
    group: "edit", description: "Find text and select one match (the first, or the one at \"index\"). The options are Word's (see find-query.ts); an invalid pattern is an error.",
    args: { text: "string", ...FIND_OPTION_ARGS, "index?": "number" }, returns: "how many matches there are",
    run: (e, a) => {
      const text = a.string("text");
      const options = findOptionsOf(a, "find", e, text);
      const matches = e.find(text, options);
      const index = a.optNumber("index") ?? 0;
      if (matches.length > 0) e.selectMatch(matches[Math.max(0, Math.min(matches.length - 1, Math.trunc(index)))]);
      return matches.length;
    },
  },
  replaceAll: {
    group: "edit", description: "Replace every match of the text (one undo step). The replacement reads ^p, ^&, \\1 or $1 as the options say.",
    args: { text: "string", with: "string", ...FIND_OPTION_ARGS },
    returns: "how many were replaced",
    run: (e, a) => {
      const text = a.string("text");
      return e.replaceAll(text, a.string("with"), findOptionsOf(a, "replaceAll", e, text));
    },
  },
  moveCaret: {
    group: "edit", description: "Put the caret at the start or end of the document or of the paragraph.",
    args: { to: "documentStart|documentEnd|paragraphStart|paragraphEnd" },
    run: (e, a) => must(e.moveCaret(a.oneOf("to", ["documentStart", "documentEnd", "paragraphStart", "paragraphEnd"] as const)), "moveCaret", "there is no paragraph to move to"),
  },
  deleteParagraph: { group: "edit", description: "Delete the paragraph (or table) at the caret.", args: {}, run: (e) => e.deleteBlock() },

  // ── picture ──
  insertPicture: {
    group: "insert", description: "Insert a picture at the caret, in line with the text: base64 bytes of a PNG, JPEG, GIF, BMP, TIFF or WebP; width in points (the height keeps the shape).",
    args: { base64: "string", "altText?": "string", "width?": "number" },
    run: (e, a) => {
      const width = a.optNumber("width");
      const altText = a.optString("altText");
      must(e.insertImage(a.string("base64"), { ...(altText !== undefined ? { altText } : {}), ...(width !== undefined ? { widthPoints: width } : {}) }),
        "insertPicture", "the picture goes at the caret in an editable paragraph; check the bytes");
    },
  },
  selectPicture: {
    group: "picture", description: "Select a picture of the body, as a click on it does (index 0 = the first).",
    args: { "index?": "number" },
    run: (e, a) => must(e.selectPicture(Math.trunc(a.optNumber("index") ?? 0)), "selectPicture", "there is no such picture"),
  },
  pictureSize: {
    group: "picture", description: "Size the selected picture in points: one side keeps the shape (lockAspect, the default); both set both.",
    args: { "width?": "number", "height?": "number", "lockAspect?": "boolean" },
    run: (e, a) => must(e.setPictureSize(a.optNumber("width"), a.optNumber("height"), a.optBool("lockAspect") ?? true),
      "pictureSize", "select a picture first; give a width or a height"),
  },
  resetPictureSize: {
    group: "picture", description: "Give the selected picture its own size again.", args: {},
    run: (e) => must(e.resetPictureSize(), "resetPictureSize", "select a picture whose own size is known"),
  },
  pictureWrap: {
    group: "picture", description: "Word's Wrap Text for the selected picture.",
    args: { wrap: PICTURE_WRAPS.join("|") },
    run: (e, a) => must(e.setPictureWrap(a.oneOf("wrap", PICTURE_WRAPS)), "pictureWrap",
      "select a picture; this one's wrapping cannot be changed"),
  },
  pictureLayout: {
    group: "picture", description: "Position and wrap the selected picture (floating): a FloatingImageLayout object, offsets in EMUs.",
    args: { layout: "object" },
    run: (e, a) => must(e.setPictureLayout(a.object("layout") as FloatingImageLayout), "pictureLayout",
      "select a picture whose position can be changed; check the layout"),
  },
  pictureAltText: {
    group: "picture", description: "The selected picture's alt text (description) and title.",
    args: { altText: "string", "title?": "string" },
    run: (e, a) => must(e.setPictureAltText(a.string("altText"), a.optString("title")), "pictureAltText", "select a picture first"),
  },
  deletePicture: {
    group: "picture", description: "Delete the selected picture.", args: {},
    run: (e) => must(e.removePicture(), "deletePicture", "select a picture first"),
  },

  // ── view ──
  zoom: { group: "view", description: "Set the zoom (1 = 100%).", args: { scale: "number" }, run: (e, a) => e.setZoom(a.number("scale")) },
  pageWidth: { group: "view", description: "Zoom so the page is as wide as the window (kept as it resizes).", args: {}, run: (e) => e.fitPageWidth() },
  onePage: { group: "view", description: "Zoom so a whole page shows (kept as the window resizes).", args: {}, run: (e) => e.fitOnePage() },
  pageView: { group: "view", description: "Show the pages (or one continuous sheet).", args: { on: "boolean" }, run: (e, a) => e.setPaginated(a.bool("on")) },
  outlineView: { group: "view", description: "Word's Outline View on or off (off: back to the view before).", args: { on: "boolean" }, run: (e, a) => e.setOutlineView(a.bool("on")) },
  outlineLevel: {
    group: "view", description: "Outline View: give the selected paragraphs an outline level (1-9; 0 = body text).",
    args: { level: "number" }, run: (e, a) => must(e.setOutlineLevel(a.number("level")), "outlineLevel", "turn Outline View on first"),
  },
  outlinePromote: { group: "view", description: "Outline View: promote the selected paragraphs one level.", args: {}, run: (e) => must(e.outlinePromote(-1), "outlinePromote", "turn Outline View on first") },
  outlineDemote: { group: "view", description: "Outline View: demote the selected paragraphs one level.", args: {}, run: (e) => must(e.outlinePromote(1), "outlineDemote", "turn Outline View on first") },
  outlineMove: {
    group: "view", description: "Outline View: move the selected paragraphs (a collapsed heading with its part) up or down.",
    args: { direction: "up|down" }, run: (e, a) => must(e.outlineMove(a.oneOf("direction", ["up", "down"] as const)), "outlineMove", "nothing to move past"),
  },
  outlineExpand: { group: "view", description: "Outline View: expand the selected headings.", args: {}, run: (e) => e.outlineExpand() },
  outlineCollapse: { group: "view", description: "Outline View: collapse the selected headings.", args: {}, run: (e) => e.outlineCollapse() },
  outlineShowLevel: {
    group: "view", description: "Outline View: show headings down to this level only (0 = all levels).",
    args: { level: "number" }, run: (e, a) => e.outlineShowLevel(a.number("level") || null),
  },
  outlineFirstLine: { group: "view", description: "Outline View: show only the first line of body text.", args: { on: "boolean" }, run: (e, a) => e.setOutlineFirstLineOnly(a.bool("on")) },
  outlineFormatting: { group: "view", description: "Outline View: show the text's formatting (off: plain).", args: { on: "boolean" }, run: (e, a) => e.setOutlineTextFormatting(a.bool("on")) },

  // ── query ──
  text: { group: "query", description: "The document's text, a line for each paragraph (and table cell).", args: {}, returns: "string", run: (e) => e.documentText() },
  selectedText: { group: "query", description: "The selected text (lines between paragraphs); \"\" with no selection.", args: {}, returns: "string", run: (e) => e.selectedText() },
  cellSelection: {
    group: "query", description: "The cells selected as whole cells: {rows: [first, last], columns: [first, last], cells, tableRows, tableColumns} (from 0, on the table's grid); null when no cells are.",
    args: {}, returns: "object, or null", run: (e) => e.cellSelection(),
  },
  selectedHtml: {
    group: "query", description: "The selection as copy puts it on the clipboard (HTML with its formatting inline; pasteHtml takes it back), without touching the clipboard; \"\" with no selection.",
    args: {}, returns: "string", run: (e) => e.selectionHtml()?.html ?? "",
  },
  paragraphText: { group: "query", description: "The text of the paragraph at the caret.", args: {}, returns: "string", run: (e) => e.paragraphText() },
  wordCount: { group: "query", description: "How many words the body has.", args: {}, returns: "number", run: (e) => e.wordCount() },
  paragraphFormats: {
    group: "query",
    description: "The effective paragraph formatting of each selected paragraph (the caret's when nothing is selected): the paragraphFormat fields, in twips.",
    args: {}, returns: "array of objects", run: (e) => e.paragraphFormats(),
  },
  listLevels: {
    group: "query", description: "The levels of the list at the caret: level, format, text, start, numberPositionTwips, textPositionTwips, legal.",
    args: {}, returns: "array of objects, or null", run: (e) => e.listLevels(),
  },
  styleAtCaret: { group: "query", description: "The paragraph style at the caret.", args: {}, returns: "style id, or null", run: (e) => e.styleAtCaret() },
  fontSizeAtCaret: { group: "query", description: "The font size at the caret, in points.", args: {}, returns: "number, or null", run: (e) => e.fontSizeAtCaret() },
  formatState: {
    group: "query", description: "Whether the selection is bold, italic, underlined, … (as the ribbon shows it).", args: {},
    returns: "{ bold, italic, underline, strike, code, superscript, subscript }", run: (e) => e.queryFormatState(),
  },
  styles: {
    group: "query", description: "The document's paragraph styles.", args: {}, returns: "[{ id, name, … }]",
    run: (e) => e.styles(),
  },
  picture: {
    group: "query", description: "The selected picture: id, wrap, size in points, alt text, title, layout, what can change.",
    args: {}, returns: "object, or null", run: (e) => e.selectedPicture(),
  },
  outline: {
    group: "query", description: "The document's headings, as the Navigation Pane lists them: level (1-9) and text.",
    args: {}, returns: "array of { level, text }",
    run: (e) => e.outline().map((h) => ({ level: h.level, text: h.text })),
  },
  bookmarks: {
    group: "query", description: "The document's bookmarks in document order: name and text (hidden: Word's own _Toc…, _Ref…, _GoBack too).",
    args: { "hidden?": "boolean" }, returns: "array of { name, text }",
    run: (e, a) => e.bookmarks(a.optBool("hidden") ?? false).map((b) => ({ name: b.name, text: b.text })),
  },
  tableInfo: {
    group: "query", description: "The Table Properties of the table, row and cell at the caret (width, alignment, indent, margins, alt text, row height, column widths, vertical alignment).",
    args: {}, returns: "object, or null", run: (e) => e.tableProperties(),
  },
  pageInfo: { group: "query", description: "In page view, the caret's page and the page count.", args: {}, returns: "{ page, total }, or null", run: (e) => e.pageInfo() },
  version: { group: "query", description: "The document's version: it goes up with every committed edit.", args: {}, returns: "number, or null", run: (e) => e.version },
};

/** The commands `exec` knows, in groups. */
export function editorCommands(): EditorCommandInfo[] {
  return Object.entries(COMMANDS).map(([name, spec]) => ({
    name, group: spec.group, description: spec.description, args: { ...spec.args },
    ...(spec.returns ? { returns: spec.returns } : {}),
  }));
}

/** Run one command on `editor` (DocxEditor.exec). */
export function runEditorCommand(editor: DocxEditor, name: string, args: EditorCommandArgs = {}): unknown {
  const spec = Object.prototype.hasOwnProperty.call(COMMANDS, name) ? COMMANDS[name] : undefined;
  if (!spec) throw new EditorCommandError("unknown_command", name, `There is no editor command "${name}".`);
  if (args === null || typeof args !== "object" || Array.isArray(args)) {
    throw new EditorCommandError("bad_arguments", name, `${name}: the arguments must be an object`);
  }
  const result = spec.run(editor, new Args(name, args));
  return result === undefined ? null : result;
}
