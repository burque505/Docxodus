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
import type { DocxEditor, EditorAlignment, EditorPageSetup, FormatKey } from "./editor.js";
import type { ListFormat, NumberFormat, TableBorderSpec, TableOfContentsOptions } from "./types.js";
import { TrackedChangeMode } from "./types.js";

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
   *  `"layout"`, `"table"`, `"review"`, `"edit"`, `"view"`, `"query"`. A query only reads the
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
}

/** A command that reports "could not act" through a false result. */
function must(ok: boolean, command: string, why: string): void {
  if (!ok) throw new EditorCommandError("refused", command, `${command}: ${why}`);
}

const ALIGNMENTS: readonly EditorAlignment[] = ["left", "center", "right", "justify"];
const LIST_FORMATS: readonly ListFormat[] = [
  "none", "bullet", "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman",
  "decimalParenthesis", "lowerLetterParenthesis", "upperLetterParenthesis", "lowerRomanParenthesis",
  "upperRomanParenthesis",
];
const PAGE_SETUP_TWIPS = [
  "marginTopTwips", "marginBottomTwips", "marginLeftTwips", "marginRightTwips", "pageWidthTwips", "pageHeightTwips",
] as const;
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

  // ── paragraph ──
  align: {
    group: "paragraph", description: "Align the paragraph.", args: { alignment: ALIGNMENTS.join("|") },
    run: (e, a) => e.setAlignment(a.oneOf("alignment", ALIGNMENTS)),
  },
  style: {
    group: "paragraph", description: "Give the paragraph a style, by its id (\"Heading1\", \"Normal\", …; see the \"styles\" query).",
    args: { styleId: "string" }, run: (e, a) => e.setParagraphStyle(a.string("styleId")),
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
  tableOfContents: {
    group: "insert", description: "Insert a table of contents (options as TableOfContentsOptions).", args: { "options?": "object" },
    run: (e, a) => must(e.insertTableOfContents((a.optObject("options") ?? {}) as TableOfContentsOptions), "tableOfContents", "the engine refused it"),
  },
  hyperlink: {
    group: "insert", description: "Link the selection to a web address, or to a bookmark (kind \"internal\").",
    args: { target: "string", "kind?": "external|internal" },
    run: (e, a) => must(e.insertHyperlink(a.string("target"), a.optOneOf("kind", ["external", "internal"] as const) ?? "external"), "hyperlink", "select the text to link first"),
  },
  removeHyperlink: { group: "insert", description: "Remove the link at the caret, keeping its text.", args: {}, run: (e) => { e.removeHyperlink(); } },
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
  pageSetup: {
    group: "layout",
    description: "Set the section's page: margins and size in twips (1440 an inch), landscape or not; what is left out stays.",
    args: {
      "marginTopTwips?": "number", "marginBottomTwips?": "number", "marginLeftTwips?": "number", "marginRightTwips?": "number",
      "pageWidthTwips?": "number", "pageHeightTwips?": "number", "landscape?": "boolean",
    },
    run: (e, a) => {
      const op: Record<string, number | boolean> = {};
      for (const key of PAGE_SETUP_TWIPS) {
        const v = a.optNumber(key);
        if (v !== undefined) op[key] = Math.round(v);
      }
      const landscape = a.optBool("landscape");
      if (landscape !== undefined) op.landscape = landscape;
      if (Object.keys(op).length === 0) throw new EditorCommandError("bad_arguments", "pageSetup", "pageSetup: give at least one setting");
      must(e.setPageSetup(op as EditorPageSetup), "pageSetup", "this engine build cannot change the page setup");
    },
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
  find: {
    group: "edit", description: "Find text and select one match (the first, or the one at \"index\").",
    args: { text: "string", "matchCase?": "boolean", "index?": "number" }, returns: "how many matches there are",
    run: (e, a) => {
      const matches = e.find(a.string("text"), { matchCase: a.optBool("matchCase") });
      const index = a.optNumber("index") ?? 0;
      if (matches.length > 0) e.selectMatch(matches[Math.max(0, Math.min(matches.length - 1, Math.trunc(index)))]);
      return matches.length;
    },
  },
  replaceAll: {
    group: "edit", description: "Replace every match of the text.", args: { text: "string", with: "string", "matchCase?": "boolean" },
    returns: "how many were replaced",
    run: (e, a) => e.replaceAll(a.string("text"), a.string("with"), { matchCase: a.optBool("matchCase") }),
  },
  moveCaret: {
    group: "edit", description: "Put the caret at the start or end of the document or of the paragraph.",
    args: { to: "documentStart|documentEnd|paragraphStart|paragraphEnd" },
    run: (e, a) => must(e.moveCaret(a.oneOf("to", ["documentStart", "documentEnd", "paragraphStart", "paragraphEnd"] as const)), "moveCaret", "there is no paragraph to move to"),
  },
  deleteParagraph: { group: "edit", description: "Delete the paragraph (or table) at the caret.", args: {}, run: (e) => e.deleteBlock() },

  // ── view ──
  zoom: { group: "view", description: "Set the zoom (1 = 100%).", args: { scale: "number" }, run: (e, a) => e.setZoom(a.number("scale")) },
  pageWidth: { group: "view", description: "Zoom so the page is as wide as the window (kept as it resizes).", args: {}, run: (e) => e.fitPageWidth() },
  onePage: { group: "view", description: "Zoom so a whole page shows (kept as the window resizes).", args: {}, run: (e) => e.fitOnePage() },
  pageView: { group: "view", description: "Show the pages (or one continuous sheet).", args: { on: "boolean" }, run: (e, a) => e.setPaginated(a.bool("on")) },

  // ── query ──
  text: { group: "query", description: "The document's text, a line for each paragraph (and table cell).", args: {}, returns: "string", run: (e) => e.documentText() },
  selectedText: { group: "query", description: "The selected text (lines between paragraphs); \"\" with no selection.", args: {}, returns: "string", run: (e) => e.selectedText() },
  paragraphText: { group: "query", description: "The text of the paragraph at the caret.", args: {}, returns: "string", run: (e) => e.paragraphText() },
  wordCount: { group: "query", description: "How many words the body has.", args: {}, returns: "number", run: (e) => e.wordCount() },
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
