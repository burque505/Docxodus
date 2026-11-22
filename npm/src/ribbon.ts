/**
 * The ribbon surface — the full editor UI as a mountable module.
 *
 * `DocxEditor` (editor.ts) is the document engine: it owns the live `DocxSession`,
 * the block wiring and every command. It has deliberately no chrome. This module is
 * the chrome — the tabbed ribbon, the status bar (with the anchor rail), the table
 * picker, the find bar and the loading overlay — wired onto exactly that command
 * surface and nothing else.
 *
 * It exists because the same UI was hand-written three times (the standalone demo,
 * the GitHub Pages landing page, the compact iframe player) and drifted. Now the
 * surface has one owner, and the three hosts differ only in how they obtain the
 * WASM exports and how much of the chrome they turn on:
 *
 * ```ts
 * // Host that already booted the .NET runtime itself:
 * const ribbon = mountRibbon(document.querySelector("#app")!, { exports });
 * ribbon.open(bytes, "contract.docx");
 *
 * // Host that wants one call and a CDN (see embed.ts):
 * await createRibbonEditor("#app", "./contract.docx", { chrome: "auto" });
 * ```
 *
 * Chrome density is measured from the ROOT ELEMENT, not the viewport: a narrow
 * embed on a wide desktop page is narrow. See ribbon-chrome.ts for the layout.
 */

import { DocxEditor } from "./editor.js";
import type {
  DocxEditorExports,
  DocxEditorOptions,
  EditorAlignment,
  EditorMatch,
  EditorPageSetup,
  FormatKey,
  PasteFormatting,
} from "./editor.js";
import { threadMembers } from "./editor-comments.js";
import { EditorCommandError } from "./editor-commands.js";
import type { EditorCommandArgs, EditorCommandInfo } from "./editor-commands.js";
import { RibbonHistory } from './ribbon-history.js';
import type { RibbonHistoryBinding } from './ribbon-history.js';
export type { RibbonHistoryBinding, RibbonHistoryOptions } from './ribbon-history.js';
import type { BandWhich } from "./editor-headerfooter.js";
import { TrackedChangeMode } from "./types.js";
import type { ColumnsApplyTo, ColumnsOp, DocumentRange, LineNumbering, ListFormat, ListLevelSpec, NumberFormat, SectionStart } from "./types.js";
import {
  COMMON_FONTS,
  HIGHLIGHT_COLORS,
  RIBBON_CSS,
  RIBBON_HINT_HTML,
  RIBBON_HTML,
  RIBBON_STATUS_TEXTS,
  RIBBON_STYLE_ATTR,
  RIBBON_STYLE_VERSION,
} from "./ribbon-chrome.js";
import { PARAGRAPH_DIALOG_CSS, PARAGRAPH_DIALOG_HTML, ParagraphDialog } from "./ribbon-paragraph-dialog.js";
import type { MeasurementUnit } from "./ribbon-paragraph-dialog.js";
import type { CaseMode, HeadlineCaseLists } from "./change-case.js";
import { FIND_OPTION_KEYS, findOptionApplies, normalizeFindOptions, specialFindCodes, specialReplaceCodes, type FindOptions } from "./find-query.js";
import { LIST_DIALOG_CSS, LIST_DIALOG_HTML, ListDialog } from "./ribbon-list-dialog.js";
import { COLUMNS_DIALOG_CSS, COLUMNS_DIALOG_HTML, ColumnsDialog } from "./ribbon-columns-dialog.js";
import { LINE_NUMBERS_DIALOG_CSS, LINE_NUMBERS_DIALOG_HTML, LineNumbersDialog } from "./ribbon-linenumbers-dialog.js";
import { TABS_DIALOG_CSS, TABS_DIALOG_HTML, TabsDialog } from "./ribbon-tabs-dialog.js";
import { BOOKMARK_DIALOG_CSS, BOOKMARK_DIALOG_HTML, BookmarkDialog } from "./ribbon-bookmark-dialog.js";
import { CROSSREF_DIALOG_CSS, CROSSREF_DIALOG_HTML, CrossReferenceDialog } from "./ribbon-crossref-dialog.js";
import { crossReferenceOptionsFor } from "./cross-reference.js";
import { NAV_PANE_CSS, NavPane, type NavPaneCommand } from "./ribbon-navpane.js";
import type { OutlineEntry } from "./types.js";
import { Ruler, RULER_CSS } from "./ruler.js";
import { TABLE_DIALOG_CSS, TABLE_DIALOG_HTML, TableDialog, type TableDialogResult, type TableDialogTab } from "./ribbon-table-dialog.js";
import { PICTURE_DIALOG_CSS, PICTURE_DIALOG_HTML, PICTURE_MENU_HTML, PictureDialog, type PictureDialogResult, type PictureDialogTab } from "./ribbon-picture-dialog.js";
import { formatLength, parseLength } from "./ribbon-paragraph-dialog.js";
import type { PictureInfo, PictureWrap } from "./editor.js";
export type { MeasurementUnit } from "./ribbon-paragraph-dialog.js";

// Re-exported so the IIFE bundle built from this module (window.DocxodusEditor)
// carries both the shell and the engine it wraps.
export { DocxEditor } from "./editor.js";
export type { DocxEditorExports, DocxEditorOptions, FormatPainterState, PasteContent, PasteFormatting } from "./editor.js";
export { DEFAULT_HEADLINE_CASE_TEXT, parseHeadlineCaseLists } from "./change-case.js";
export type { CaseMode, HeadlineCaseLists } from "./change-case.js";
export { compileFind, expandReplacement, specialFindCodes, specialReplaceCodes } from "./find-query.js";
export type { FindOptions } from "./find-query.js";

/** Which layout the chrome uses. "auto" measures the root and switches at `compactBreakpoint`. */
export type RibbonChromeMode = "full" | "compact" | "auto";

/** The lifecycle the surface reports on its root as `data-state`. */
export type RibbonState = "idle" | "loading" | "ready" | "error";

/** One step of the loading narrative: what is happening, and how far along it is. */
export interface RibbonLoaderStage {
  title: string;
  copy: string;
  /** 0–100. Drives the progress bar. */
  progress: number;
  /** The short machine-side label under the bar. */
  label: string;
}

/** A rotating capability card shown while the engine streams. */
export interface RibbonLoaderFeature {
  number: string;
  title: string;
  copy: string;
}

export interface RibbonLoaderOptions {
  /** Replace the four default stages. */
  stages?: RibbonLoaderStage[];
  /** Replace the rotating cards. Pass `[]` to hide the card entirely. */
  features?: RibbonLoaderFeature[];
  /** Micro-caps line above the title. */
  eyebrow?: string;
  /** Right-hand label under the progress bar. */
  meta?: string;
  /** Milliseconds between card rotations. Default 1750. */
  rotateMs?: number;
  /** What the Retry button does. Default: reload the page. */
  onRetry?: () => void;
}

export interface RibbonOptions extends DocxEditorOptions {
  /** Optional version-history services. createRibbonEditor accepts history: true to wire browser storage. */
  history?: RibbonHistoryBinding;
  /** WASM exports. Omit to mount chrome first (showing the loader) and call `setExports` later. */
  exports?: DocxEditorExports;
  /** Layout density. Default "auto". */
  chrome?: RibbonChromeMode;
  /**
   * Boundary treatment. "flush" (default) is edge-to-edge for full-bleed hosts
   * and hosts that frame the surface themselves; "card" makes the surface carry
   * its own boundary — rounded corners, hairline border, house shadow — for a
   * drop-in embed in the middle of a page. `createRibbonEditor` defaults to
   * "card", since that is the drop-in path.
   */
  frame?: "card" | "flush";
  /** Root width, in px, below which "auto" picks compact. Default 720. */
  compactBreakpoint?: number;
  /** Show the status bar with the anchor rail (full chrome only). Default true. */
  rail?: boolean;
  /** Show New / Open / Save. Default true. */
  fileActions?: boolean;
  /**
   * Editing hint above the document: false to hide, or a string to replace it.
   * Rendered as HTML (the default copy uses `<kbd>`), so pass author-controlled
   * markup only. Default true.
   */
  hint?: boolean | string;
  /** Loading overlay: false to suppress it, or an object to reword it. Default true. */
  loader?: boolean | RibbonLoaderOptions;
  /** Name shown in the title bar and used for the downloaded file. Default "untitled.docx". */
  documentName?: string;
  /**
   * Element-id prefix. Controls are addressable as `data-dxr="<name>"` always, and
   * additionally get `id = idPrefix + name`. Omit and the surface uses bare ids when
   * they are free on the page, falling back to a generated prefix when they are not —
   * so a second ribbon never collides with the first.
   */
  idPrefix?: string;
  /**
   * Replace the default Save behaviour (download the bytes). The document stays changed
   * (`isDirty`) until the host calls `markSaved()`, once it has stored the bytes: a host whose
   * save can be cancelled or fail must not lose the changed state.
   */
  onSave?: (bytes: Uint8Array, name: string) => void;
  /**
   * Called when the document turns changed, or back to unchanged (it was saved, or another
   * document opened). See `isDirty`.
   */
  onDirtyChange?: (dirty: boolean) => void;
  /**
   * Replace New's default (open a blank document at once). For a host that asks to save the
   * changed document first; it opens the blank one itself (`openBlank`).
   */
  onNewRequest?: () => void;
  /**
   * Replace Open's file picker (and Ctrl+O's). For a host with a picker of its own, such as a
   * desktop host that needs the file's path to save it back; it opens the bytes itself (`open`).
   */
  onOpenRequest?: () => void;
  /**
   * Open a document from the Recent list beside Open, which shows the entries given to
   * `setRecent`. Without this option there is no Recent list.
   */
  onOpenRecent?: (entry: RibbonRecentEntry) => void;
  /**
   * Save As in the File menu: the host asks where and saves there. Without this option the File
   * menu has no Save As (the default Save downloads, and a browser asks where itself).
   */
  onSaveAsRequest?: () => void;
  /**
   * The host's own commands at the foot of the File menu (Print, Settings, Close, ...), after the
   * ribbon's New, Open, Save, Save As and recent documents. `setFileMenu` replaces them later.
   */
  fileMenu?: RibbonFileMenuItem[];
  /** One of the host's File menu commands was chosen, by its `id`. */
  onFileMenu?: (id: string) => void;
  /**
   * The ribbon's text in another language: each label, menu entry, tooltip and placeholder, keyed
   * by the English it replaces (`{ "Home": "Inicio", "Bold (Ctrl+B)": "Negrita (Ctrl+N)" }`). Text
   * without an entry stays English; the document is never touched. `setLabels` changes it later.
   */
  labels?: Record<string, string>;
  /** How the Paragraph dialog shows indents: inches (`"in"`, the default) or centimetres
   *  (`"cm"`). Any unit can be typed either way. `setMeasurementUnit` changes it later. */
  measurementUnit?: MeasurementUnit;
  /** The levels every new list starts from (the editor's `listDefaults`); `setListDefaults`
   *  changes them later. */
  listDefaults?: ListLevelSpec[];
  /**
   * The List Levels dialog's "Also for new documents" was ticked: the host keeps these levels
   * (for later documents too) and gives them back as `listDefaults`. Without this option the
   * dialog has no such box.
   */
  onListDefaults?: (levels: ListLevelSpec[]) => void;
  /** Headline Case's word lists (the editor's `headlineCaseLists`; a host that keeps them in a
   *  file reads it with `parseHeadlineCaseLists`); `setHeadlineCaseLists` changes them later. */
  headlineCaseLists?: HeadlineCaseLists;
  /** Show/Hide ¶ at the start (Word's paragraph marks and tabs); `setShowFormattingMarks`
   *  changes it later. It stays as it is from one document to the next. */
  showFormattingMarks?: boolean;
  /** Show/Hide ¶ was switched (the button, Ctrl+Shift+8 or the command), for a host to keep. */
  onFormattingMarksChange?: (on: boolean) => void;
  /** Open with the Navigation Pane showing (View > Navigation Pane). */
  navigationPane?: boolean;
  /** The Navigation Pane was opened or closed, for a host to keep. */
  onNavigationPaneChange?: (on: boolean) => void;
  /** Find's options at the start (Match case, whole words, wildcards...: see find-query.ts);
   *  `setFindOptions` changes them later. */
  findOptions?: FindOptions;
  /** Find's options were changed in the find bar's Options menu, for a host to keep (only the
   *  options that are on). */
  onFindOptionsChange?: (options: FindOptions) => void;
  /** Called after any document opens, with the live editor. */
  onOpen?: (editor: DocxEditor) => void;
  /** Called whenever the status line changes. */
  onStatus?: (text: string) => void;
  /**
   * Called after every ribbon command, with its label and measured duration; for one that
   * succeeded, also the command it was as `exec` takes it (`{ name: "fontSize", args: { pts: 14 } }`),
   * whatever ran it: a button, a list, a shortcut (the editor's own Ctrl+B and Ctrl+Z among them),
   * or `exec` itself. Running those again repeats what the user did: what a macro recorder keeps.
   * A command with no `exec` equivalent (a comment's Resolve, the next tracked change) comes
   * without one.
   */
  onCommand?: (label: string, ms: number, command?: RibbonCommand) => void;
}

/** One document in the Recent list beside Open (`setRecent`, `onOpenRecent`). */
export interface RibbonRecentEntry {
  /** The host's key for it, such as a file path. Handed back to `onOpenRecent`. */
  id: string;
  /** What the list shows, such as the file name. */
  name: string;
  /** The entry's tooltip, such as the full path. */
  title?: string;
}

/** One of the host's commands in the File menu (`fileMenu`, `setFileMenu`, `onFileMenu`). */
export interface RibbonFileMenuItem {
  /** Handed back to `onFileMenu`. */
  id: string;
  /** What the menu shows. */
  label: string;
  /** The item's tooltip. */
  title?: string;
  /** The key that does the same, shown beside the label (the host handles the key itself). */
  shortcut?: string;
  /** A rule above this item. */
  separatorBefore?: boolean;
  /** Shown but not available. */
  disabled?: boolean;
}

/** Drives the loading overlay. Safe to call when the loader is disabled — every method no-ops. */
export interface RibbonLoader {
  /** Re-show the overlay (e.g. before loading a second document). */
  show(): void;
  /** Jump to a numbered stage, or set one inline. */
  stage(step: number | Partial<RibbonLoaderStage>): void;
  /** Move the bar without changing the copy. */
  progress(percent: number, label?: string): void;
  /** Fade the overlay out and stop the rotation. */
  done(): void;
  /** Show the failure state with a Retry button. */
  fail(error: unknown): void;
}

/** A command as `exec` takes it: what `onCommand` reports. */
export interface RibbonCommand {
  name: string;
  args: EditorCommandArgs;
}
/** A control's command: known when it is wired, or worked out once it has run (zoom in: the zoom it reached). */
type CommandOf = RibbonCommand | (() => RibbonCommand | undefined);
const cmd = (name: string, args: EditorCommandArgs = {}): RibbonCommand => ({ name, args });

/** The commands that only read the document (DocxEditor.commands(), group "query"). */
const QUERY_COMMANDS: ReadonlySet<string> = new Set(DocxEditor.commands().filter((c) => c.group === "query").map((c) => c.name));

export interface RibbonEditor {
  readonly history: RibbonHistory | null;
  /** The mounted root element (carries `data-state` and `data-chrome`). */
  readonly element: HTMLElement;
  /** The element the document renders into. */
  readonly surface: HTMLElement;
  /** The live editor, or null before a document is open. */
  readonly editor: DocxEditor | null;
  /** The density actually in effect right now. */
  readonly chrome: "full" | "compact";
  /** The loading overlay controller. */
  readonly loader: RibbonLoader;
  /** Supply (or replace) the WASM exports after mounting. */
  setExports(exports: DocxEditorExports): void;
  /** Open a document, replacing any open one. Throws if exports are not set yet. */
  open(bytes: Uint8Array, name?: string): DocxEditor;
  /** {@link open} window by window (`DocxEditor.openAsync`): a long document mounts while the page
   *  keeps painting, `onProgress` telling how far it has got. */
  openAsync(bytes: Uint8Array, name?: string, options?: { onProgress?: (mounted: number, total: number) => void; windowSize?: number }): Promise<DocxEditor>;
  /** Open a fresh blank document. */
  openBlank(name?: string): DocxEditor;
  /** Lossless DOCX bytes, or null when nothing is open. */
  save(): Uint8Array | null;
  /** Save and hand the bytes to `onSave` (default: browser download). */
  download(name?: string): void;
  /**
   * True when the document has changed since it was opened or last marked saved: the session
   * moved to another version, or the user typed in it (typing the editor has not committed yet
   * counts too). False when nothing is open.
   */
  readonly isDirty: boolean;
  /** The document as it is now counts as saved. The default Save does this itself; a host with `onSave` calls it. */
  markSaved(): void;
  /** The documents the Recent list beside Open shows, most recent first ([] hides the list). Needs `onOpenRecent`. */
  setRecent(entries: RibbonRecentEntry[]): void;
  /** Put the ribbon's text in another language (see the `labels` option); `{}` brings back the English. */
  setLabels(labels: Record<string, string>): void;
  /** Replace the host's commands at the foot of the File menu (see the `fileMenu` option). */
  setFileMenu(items: RibbonFileMenuItem[]): void;
  /** Open the File menu (as its tab's click, or Alt+F, does). */
  openFileMenu(): void;
  /** Open the Paragraph dialog on the selected paragraphs (as the Paragraph group's launcher
   *  does); false when no document is open or no paragraph is active. */
  openParagraphDialog(): boolean;
  /** Show the Paragraph dialog's indents in inches or centimetres (the `measurementUnit` option). */
  setMeasurementUnit(unit: MeasurementUnit): void;
  /** Open the List Levels dialog on the caret's list; false when no document is open. */
  openListDialog(): boolean;
  /** Open the Columns dialog (Word's More Columns) on the caret's section; false when no document is open. */
  openColumnsDialog(): boolean;
  /** Open the Line Numbers dialog (Word's Line Numbering Options) on the caret's section; false
   *  when no document is open. */
  openLineNumbersDialog(): boolean;
  /** Open the Tabs dialog on the caret paragraph's tab stops; false when there is no paragraph. */
  openTabsDialog(): boolean;
  /** Show or hide the ruler (the View tab's Ruler box). */
  setRulerVisible(on: boolean): void;
  /** Open the Bookmark dialog (Word's Insert > Bookmark) on the selection; false when no document is open. */
  openBookmarkDialog(): boolean;
  /** Open the Cross-reference dialog (Word's Insert > Cross-reference) at the selection; false when
   *  no document is open. */
  openCrossReferenceDialog(): boolean;
  /** Show or hide bookmarks in the text as gray brackets (the View tab's Bookmarks box, Word's
   *  Show bookmarks). */
  setShowBookmarks(on: boolean): void;
  /** Word's Update Field: the fields in the selection (F9), every field (`all`), or the tables of
   *  contents (`tables`, Update Table). False when nothing is open or the engine refused. */
  updateFields(scope?: "selection" | "all" | "tables"): boolean;
  /** Show or hide the Navigation Pane (the View tab's Navigation Pane box). */
  setNavigationPaneVisible(on: boolean): void;
  /** Open the picture dialog (Size, Text Wrapping and Position, Alt Text) on the selected picture;
   *  false when no picture is selected. */
  openPictureDialog(tab?: "size" | "layout" | "alt"): boolean;
  /** Open the Table Properties dialog (Table, Row, Column, Cell, Alt Text) on the caret's table;
   *  false when the caret is not in a table. */
  openTableDialog(tab?: "table" | "row" | "column" | "cell" | "alt"): boolean;
  /** The levels new lists start from (see the `listDefaults` option). */
  setListDefaults(levels: ListLevelSpec[]): void;
  /** Headline Case's word lists (see the `headlineCaseLists` option); null for the built-in ones. */
  setHeadlineCaseLists(lists: HeadlineCaseLists | null): void;
  /** Show/Hide ¶ (see the `showFormattingMarks` option). */
  setShowFormattingMarks(on: boolean): void;
  /** Find's options (see the `findOptions` option). */
  setFindOptions(options: FindOptions): void;
  /** Find's options as they are now. */
  getFindOptions(): FindOptions;
  /** Set the status line. */
  setStatus(text: string): void;
  /** Activate a ribbon tab by name ("home" | "insert" | "layout" | "references" | "review" | "view" | "table" | "headerfooter" | "picture"). */
  selectTab(name: string): void;
  /**
   * Run an editor command by name (`DocxEditor.exec`) as a ribbon control runs it: the rail and
   * `onCommand` report it, and the controls that show the state (view, zoom, Track Changes,
   * comments) follow. Returns the command's result; throws `EditorCommandError`, also when no
   * document is open.
   */
  exec(name: string, args?: EditorCommandArgs): unknown;
  /** The commands `exec` knows (`DocxEditor.commands()`). */
  commands(): EditorCommandInfo[];
  /** Force a density, or hand control back to measurement with "auto". */
  setChrome(mode: RibbonChromeMode): void;
  /** Look up a control by its `data-dxr` name. */
  control<T extends HTMLElement = HTMLElement>(name: string): T | null;
  /** Close the session, drop listeners and empty the container. */
  destroy(): void;
}

const DEFAULT_STAGES: RibbonLoaderStage[] = [
  {
    title: "Booting .NET inside your browser",
    copy: "Streaming the trimmed WebAssembly runtime. No document bytes are sent to a server.",
    progress: 16,
    label: "Loading engine",
  },
  {
    title: "Opening the document",
    copy: "Parsing OOXML parts, styles, numbering, tables, notes, and tracked revisions locally.",
    progress: 54,
    label: "Reading document",
  },
  {
    title: "Wiring lossless editing",
    copy: "Connecting every editable block to its native WordprocessingML anchor.",
    progress: 84,
    label: "Mounting editor",
  },
  {
    title: "Your local editor is ready",
    copy: "Select text, use the ribbon, switch page view, and save a real DOCX.",
    progress: 100,
    label: "Ready",
  },
];

const DEFAULT_FEATURES: RibbonLoaderFeature[] = [
  { number: "01", title: "Zero-upload architecture", copy: "Your source file and edits never leave this browser session." },
  { number: "02", title: "Word-grade OOXML fidelity", copy: "Tables, numbering, footnotes, redlines, comments, and styles stay native." },
  { number: "03", title: "Surgical editing", copy: "Formatting and text edits target real document anchors instead of flattening the file." },
  { number: "04", title: "Lossless DOCX out", copy: "Undo, redo, edit, and save a Word document that remains a Word document." },
];

/** Every `data-dxr` name the template defines — the id-collision probe set. */
const CONTROL_NAMES = [
  "docname", "new", "file", "recent", "save", "undo", "redo", "status", "ribbon",
  "fontsize", "fontsizes", "fontfamily", "fontgrow", "fontshrink", "clearformat", "smallcaps",
  "fontcolorbtn", "fontcolorbar", "fontcolor", "highlight", "listmenu", "linespacing",
  "style", "delblock", "findtoggle", "replacetoggle",
  "table", "picturefile", "link", "unlink", "insertcomment", "gotoheader", "gotofooter", "pagenummenu",
  "hr", "hrThick", "hrDouble", "rulepos", "hrClear",
  "margins", "orientation", "pagesize", "breaks", "sectionstart", "pagesetupnote", "spacebefore", "spaceafter",
  "specialindent", "specialindentval", "pgfmt", "pgstart", "pgclear", "paginated", "headerfooter",
  "toc", "footnote", "endnote", "pagenum", "totalpages",
  "comment", "commentdelete", "commentprev", "commentnext", "commentresolve", "showcomments",
  "commentcount", "trackchanges", "markup", "author", "revcount",
  "viewpage", "viewweb", "zoom", "zoomfit", "zoomonepage", "showrail", "showhint", "showruler", "ruler", "tabsdialog",
  "shadebtn", "shadebar", "shade", "repeatheader", "tblheight", "tblwidth", "tblprops", "tabledialog", "linenumbers", "linenumbersdialog",
  "hfFirst", "hfEven", "hfstory",
  "findbar", "findtext", "findprev", "findnext", "findcount", "findcase", "findoptions", "findspecial",
  "findoptmenu", "findspecialmenu", "replacegroup",
  "replacetext", "replaceone", "replaceall", "findclose",
  "railAnchor", "railBlocks", "railSession", "railOp", "pageinfo", "wordcount",
  "zoomout", "zoomlevel", "zoomin",
  "gridpicker", "gridcells", "gridlabel", "gridalign", "gridborderless",
  "filebutton", "filemenu", "filerecent", "filerecentitems", "filehostitems",
  "linkpop", "linkform", "linkurl", "linkbookmark", "linkcancel",
  "bookmark", "bookmarkdialog", "showbookmarks", "crossref", "crossrefdialog", "updatetoc", "fieldmenu", "navpane", "navmenu", "shownav",
  "editor",
  "loader", "loaderEyebrow", "loaderTitle", "loaderCopy", "loaderAd", "loaderNumber",
  "loaderAdTitle", "loaderAdCopy", "loaderBar", "loaderLabel", "loaderMeta", "loaderRetry",
];

const GRID_ROWS = 8;
const GRID_COLS = 10;
const ZOOM_STEPS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2];

let nextIdPrefixSeed = 0;

/** Inject the ribbon stylesheet once per document (replacing an older version). */
function ensureStyles(doc: Document): void {
  const existing = doc.querySelector<HTMLStyleElement>(`style[${RIBBON_STYLE_ATTR}]`);
  if (existing?.getAttribute(RIBBON_STYLE_ATTR) === RIBBON_STYLE_VERSION) return;
  existing?.remove();
  const style = doc.createElement("style");
  style.setAttribute(RIBBON_STYLE_ATTR, RIBBON_STYLE_VERSION);
  style.textContent = RIBBON_CSS + PARAGRAPH_DIALOG_CSS + LIST_DIALOG_CSS + COLUMNS_DIALOG_CSS + PICTURE_DIALOG_CSS + TABLE_DIALOG_CSS + LINE_NUMBERS_DIALOG_CSS + TABS_DIALOG_CSS + BOOKMARK_DIALOG_CSS + CROSSREF_DIALOG_CSS + NAV_PANE_CSS + RULER_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}

/**
 * Pick an id prefix. An explicit one always wins. Otherwise bare ids are used when
 * every name is free on the page — which keeps the historical ids of the standalone
 * demo intact — and a generated prefix is used the moment any of them is taken.
 */
function resolveIdPrefix(explicit: string | undefined, doc: Document): string {
  if (explicit !== undefined) return explicit;
  if (!CONTROL_NAMES.some((name) => doc.getElementById(name))) return "";
  for (;;) {
    const candidate = `dxr${++nextIdPrefixSeed}-`;
    if (!CONTROL_NAMES.some((name) => doc.getElementById(candidate + name))) return candidate;
  }
}

function resolveContainer(container: string | HTMLElement): HTMLElement {
  if (typeof container !== "string") return container;
  const el = document.querySelector<HTMLElement>(container);
  if (!el) throw new Error(`Docxodus ribbon: no element matches "${container}"`);
  return el;
}

/** CSS colour → "#rrggbb" (for the colour inputs), or null when it cannot be expressed. */
function toHex(color: string): string | null {
  const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(color);
  if (m) {
    return "#" + [m[1], m[2], m[3]].map((n) => parseInt(n, 10).toString(16).padStart(2, "0")).join("");
  }
  if (/^#[0-9a-f]{6}$/i.test(color)) return color.toLowerCase();
  return null;
}

/**
 * Mount the ribbon surface into `container`.
 *
 * The chrome paints synchronously — including the loading overlay — so a host can
 * mount before its WASM runtime exists, narrate the boot through `loader`, then call
 * `setExports` and `open`.
 */
export function mountRibbon(
  container: string | HTMLElement,
  options: RibbonOptions = {},
): RibbonEditor {
  return new RibbonSurface(resolveContainer(container), options);
}

class RibbonSurface implements RibbonEditor {
  readonly element: HTMLElement;
  readonly surface: HTMLElement;
  readonly loader: RibbonLoader;
  readonly history: RibbonHistory | null;

  private readonly options: RibbonOptions;
  private readonly idPrefix: string;
  private readonly loaderOptions: RibbonLoaderOptions | null;
  private readonly stages: RibbonLoaderStage[];
  private readonly features: RibbonLoaderFeature[];

  private exports: DocxEditorExports | null;
  private live: DocxEditor | null = null;
  private documentName: string;
  private chromeMode: RibbonChromeMode;
  private density: "full" | "compact" = "full";
  private headerFooter: boolean;
  private author: string;
  /** The author's initials on comments; "" takes them from the name (see DocxEditorOptions). */
  private initials: string;
  /** The Markup choice, carried to every document opened; unset, the display follows tracking. */
  private showRevisions: boolean | undefined;
  private destroyed = false;

  private featureTimer: ReturnType<typeof setInterval> | null = null;
  private featureIndex = 0;
  private resizeObserver: ResizeObserver | null = null;
  private strips: HTMLElement[] = [];
  private stripObserver: ResizeObserver | null = null;
  private lastAnchorText = "";
  private selectionFrame: number | null = null;
  private statsTimer: ReturnType<typeof setTimeout> | null = null;
  private findMatches: EditorMatch[] = [];
  private findIndex = -1;
  private findOptions: FindOptions = {};
  private findSpecialTarget: "find" | "replace" = "find";
  private findErrorText: string | null = null;
  // Changed state (isDirty): the session version and the count of typing events at the last
  // open or markSaved. Typing counts before the editor commits it (on a pause, a word or blur),
  // so the document reads as changed from the first key.
  private savedVersion: number | null = null;
  private typed = 0;
  private savedTyped = 0;
  private dirty = false;
  private dirtyCheckQueued = false;
  private recentEntries: RibbonRecentEntry[] = [];
  private fileMenuItems: RibbonFileMenuItem[] = [];
  private paragraphDialog!: ParagraphDialog;
  private columnsDialog!: ColumnsDialog;
  // (the selection the Columns dialog opened on: its fields take the focus, OK puts it back)
  private columnsRange: Range | null = null;
  private pictureDialog!: PictureDialog;
  private tableDialog!: TableDialog;
  private lineNumbersDialog!: LineNumbersDialog;
  private tabsDialog!: TabsDialog;
  private ruler!: Ruler;
  private rulerFrame: number | null = null;
  private rulerUnit: MeasurementUnit = "in";
  private rulerObserver: ResizeObserver | null = null;
  private focusRing: HTMLElement | null = null;
  private bookmarkDialog!: BookmarkDialog;
  private crossRefDialog!: CrossReferenceDialog;
  /** The layer Show bookmarks draws its brackets in (made the first time they show). */
  private bookmarkLayer: HTMLElement | null = null;
  private showBookmarks = false;
  private navPane!: NavPane;
  private pictureUnit: MeasurementUnit = "in";
  private listDialog!: ListDialog;
  private listDefaults: ListLevelSpec[] | undefined;
  private headlineCaseLists: HeadlineCaseLists | undefined;
  private showMarks = false;
  // (whether the last break insertBreakAt tried was refused: read by breakCommand)
  private breakRefused = false;
  // The chrome's text in the host's language (setLabels), and the English each text had.
  private labels: Record<string, string> = {};
  private readonly englishText = new WeakMap<Text, { english: string; written: string }>();
  private readonly englishAttrs = new WeakMap<Element, Record<string, { english: string; written: string }>>();
  private readonly onSelectionChange: () => void;
  private readonly onDocumentMouseDown: (event: MouseEvent) => void;
  private readonly onKeydown: (event: KeyboardEvent) => void;
  private readonly onInput: (event: Event) => void;

  constructor(container: HTMLElement, options: RibbonOptions) {
    this.options = options;
    this.fileMenuItems = options.fileMenu?.slice() ?? [];
    this.exports = options.exports ?? null;
    this.documentName = options.documentName ?? "untitled.docx";
    this.chromeMode = options.chrome ?? "auto";
    // Headers and footers are document content, so the surface edits them by default; the
    // engine's own default stays off for consumers that mount `DocxEditor` bare.
    this.headerFooter = options.headerFooter ?? true;
    this.showRevisions = options.showRevisions;
    this.author = options.commentAuthor ?? options.revisionAuthor ?? "Reviewer";
    this.initials = options.commentInitials ?? "";
    // false suppresses the overlay; true/omitted takes the defaults; an object rewords it.
    this.loaderOptions =
      options.loader === false
        ? null
        : typeof options.loader === "object"
          ? options.loader
          : {};
    this.stages = this.loaderOptions?.stages ?? DEFAULT_STAGES;
    this.features = this.loaderOptions?.features ?? DEFAULT_FEATURES;

    const doc = container.ownerDocument ?? document;
    ensureStyles(doc);
    this.idPrefix = resolveIdPrefix(options.idPrefix, doc);

    const root = doc.createElement("div");
    root.className = "dxr";
    root.dataset.state = "idle";
    root.dataset.frame = this.options.frame ?? "flush";
    root.innerHTML = RIBBON_HTML + PARAGRAPH_DIALOG_HTML + LIST_DIALOG_HTML + COLUMNS_DIALOG_HTML + PICTURE_DIALOG_HTML + PICTURE_MENU_HTML + TABLE_DIALOG_HTML + LINE_NUMBERS_DIALOG_HTML + TABS_DIALOG_HTML + BOOKMARK_DIALOG_HTML + CROSSREF_DIALOG_HTML;
    for (const el of Array.from(root.querySelectorAll<HTMLElement>("[data-dxr]"))) {
      el.id = this.idPrefix + el.dataset.dxr;
    }
    for (const el of Array.from(root.querySelectorAll<HTMLElement>("[data-dxr-list]"))) {
      el.setAttribute("list", this.idPrefix + el.dataset.dxrList);
    }
    container.replaceChildren(root);
    this.element = root;
    this.surface = this.require("editor");

    this.applyStaticOptions();
    this.buildGrid();
    this.populateStaticMenus();
    this.wire();
    this.wireScrollFades();
    this.applyChrome();
    if (options.labels) this.setLabels(options.labels);

    this.loader = this.createLoaderController();
    if (this.loaderOptions) this.loader.show();

    // Coalesced to one frame: selectionchange fires per keystroke and continuously
    // during a drag, and the sync reads computed styles.
    this.onSelectionChange = () => {
      if (this.selectionFrame != null) cancelAnimationFrame(this.selectionFrame);
      this.selectionFrame = requestAnimationFrame(() => {
        this.selectionFrame = null;
        this.syncSelection();
      });
    };
    doc.addEventListener("selectionchange", this.onSelectionChange);
    this.onDocumentMouseDown = (event) => this.maybeClosePopovers(event);
    doc.addEventListener("mousedown", this.onDocumentMouseDown);
    this.onKeydown = (event) => this.handleShortcut(event);
    root.addEventListener("keydown", this.onKeydown);
    // Typing in the document marks it changed at once, before the editor commits it (isDirty).
    // Text boxes of the chrome and the comment gutter are not contenteditable, so they don't count.
    this.onInput = (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement) || !target.isContentEditable || !this.surface.contains(target)) return;
      this.typed++;
      this.scheduleDirtyCheck();
    };
    root.addEventListener("input", this.onInput);

    if (this.chromeMode === "auto" && typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => this.applyChrome());
      this.resizeObserver.observe(root);
    }
    this.history = options.history ? new RibbonHistory(this, options.history) : null;
  }

  // ── element lookup ──────────────────────────────────────────────────────────

  control<T extends HTMLElement = HTMLElement>(name: string): T | null {
    return this.element.querySelector<T>(`[data-dxr="${name}"]`);
  }

  private require<T extends HTMLElement = HTMLElement>(name: string): T {
    const el = this.control<T>(name);
    if (!el) throw new Error(`Docxodus ribbon: template is missing "${name}"`);
    return el;
  }

  // ── mount-time configuration ────────────────────────────────────────────────

  private applyStaticOptions(): void {
    const hintEl = this.element.querySelector<HTMLElement>("[data-dxr-hint]");
    if (hintEl) {
      if (this.options.hint === false) hintEl.remove();
      else hintEl.innerHTML = typeof this.options.hint === "string" ? this.options.hint : RIBBON_HINT_HTML;
    }
    if (this.options.rail === false) {
      this.element.querySelector("[data-dxr-rail]")?.remove();
    }
    if (this.options.fileActions === false) {
      this.element.querySelector("[data-dxr-files]")?.remove();
    }
    if (!this.loaderOptions) this.control("loader")?.remove();
    if (!this.options.history) this.control('history')?.remove();
    else {
      const file = this.control<HTMLInputElement>('file');
      if (file) { file.accept = '.docx,.docxhistory'; file.setAttribute('aria-label', 'Open a document or history file'); }
    }

    this.require("docname").textContent = this.documentName;
    this.require<HTMLInputElement>("paginated").checked = this.options.paginated ?? false;
    this.require<HTMLInputElement>("headerfooter").checked = this.headerFooter;
    this.require<HTMLInputElement>("trackchanges").checked =
      this.options.trackedChanges === TrackedChangeMode.RenderInline;
    this.require<HTMLInputElement>("author").value = this.author;
    this.surface.dataset.view = this.options.paginated ? "paginated" : "continuous";
    this.surface.dataset.comments = this.options.comments === false ? "off" : "on";
  }

  private buildGrid(): void {
    const cells = this.require("gridcells");
    const fragment = document.createDocumentFragment();
    for (let r = 0; r < GRID_ROWS; r++) {
      for (let c = 0; c < GRID_COLS; c++) {
        const cell = document.createElement("div");
        cell.dataset.r = String(r);
        cell.dataset.c = String(c);
        fragment.appendChild(cell);
      }
    }
    cells.replaceChildren(fragment);
  }

  /** Menus whose entries do not depend on the document. */
  private populateStaticMenus(): void {
    const highlight = this.require<HTMLSelectElement>("highlight");
    for (const color of HIGHLIGHT_COLORS) {
      const opt = document.createElement("option");
      opt.value = color.value;
      opt.textContent = `■ ${color.label}`;
      opt.style.color = color.css;
      highlight.appendChild(opt);
    }
    this.populateFonts([]);
  }

  /** The font menu: the document's own fonts first, then the common list, no duplicates. */
  private populateFonts(documentFonts: string[]): void {
    const select = this.require<HTMLSelectElement>("fontfamily");
    const seen = new Set<string>();
    select.replaceChildren();
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Font…";
    select.appendChild(placeholder);
    const add = (name: string, group: HTMLElement | HTMLSelectElement) => {
      const key = name.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      opt.style.fontFamily = `"${name}", sans-serif`;
      group.appendChild(opt);
    };
    if (documentFonts.length > 0) {
      const group = document.createElement("optgroup");
      group.label = "In this document";
      for (const name of documentFonts) add(name, group);
      select.appendChild(group);
    }
    const common = document.createElement("optgroup");
    common.label = "Common fonts";
    for (const name of COMMON_FONTS) add(name, common);
    select.appendChild(common);
  }

  /** The style gallery: every paragraph style the document defines, quick styles first, plus
   *  the built-ins Word always offers (the engine creates a missing built-in on first use). */
  private populateStyles(): void {
    const select = this.require<HTMLSelectElement>("style");
    const styles = this.live?.styles() ?? [];
    const paragraphStyles = styles.filter((s) => s.type === "paragraph" && !s.semiHidden);
    if (paragraphStyles.length === 0) return; // older bundle: keep the static five
    const have = new Set(paragraphStyles.map((s) => s.id.toLowerCase()));
    const builtIns: Array<[string, string, number]> = [
      ["Normal", "Normal", 0], ["Heading1", "Heading 1", 9], ["Heading2", "Heading 2", 9],
      ["Heading3", "Heading 3", 9], ["Title", "Title", 10], ["Subtitle", "Subtitle", 11],
      ["Quote", "Quote", 29], ["ListParagraph", "List Paragraph", 34],
    ];
    for (const [id, name, priority] of builtIns) {
      if (have.has(id.toLowerCase())) continue;
      paragraphStyles.push({
        id, name, type: "paragraph", isDefault: false, isCustom: false, hasLatentException: false,
        uiPriority: priority, quickFormat: true,
      });
    }
    paragraphStyles.sort((a, b) => {
      const qa = a.quickFormat ? 0 : 1;
      const qb = b.quickFormat ? 0 : 1;
      if (qa !== qb) return qa - qb;
      const pa = a.uiPriority ?? 99;
      const pb = b.uiPriority ?? 99;
      if (pa !== pb) return pa - pb;
      return a.name.localeCompare(b.name);
    });
    select.replaceChildren();
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Style…";
    select.appendChild(placeholder);
    for (const style of paragraphStyles) {
      const opt = document.createElement("option");
      opt.value = style.id;
      opt.textContent = style.name;
      select.appendChild(opt);
    }
    this.styleNames = new Map(styles.filter((s) => s.type === "paragraph").map((s) => [s.id, s.name]));
    this.defaultStyleId = styles.find((s) => s.type === "paragraph" && s.isDefault)?.id ?? "Normal";
    this.syncStyleBox();
  }

  /** Paragraph style names by id (for a style the gallery leaves out) and the default one. */
  private styleNames = new Map<string, string>();
  private defaultStyleId = "Normal";

  /** The Styles box shows the caret paragraph's style, as Word's gallery marks it. A style the
   *  gallery leaves out (a hidden one) is added to the list for as long as it is shown. */
  private syncStyleBox(): void {
    const select = this.control<HTMLSelectElement>("style");
    if (!select || !this.live || select.ownerDocument.activeElement === select) return;
    let id: string | null = null;
    try { id = this.live.styleAtCaret(); } catch { id = null; }
    id ??= this.selectionElement() ? this.defaultStyleId : null;
    for (const extra of Array.from(select.querySelectorAll("option[data-extra]"))) if ((extra as HTMLOptionElement).value !== id) extra.remove();
    if (id && !Array.from(select.options).some((o) => o.value === id)) {
      const opt = select.ownerDocument.createElement("option");
      opt.value = id;
      opt.textContent = this.styleNames.get(id) ?? id;
      opt.dataset.extra = "";
      select.appendChild(opt);
    }
    select.value = id ?? "";
  }

  // ── chrome density ──────────────────────────────────────────────────────────

  get chrome(): "full" | "compact" {
    return this.density;
  }

  setChrome(mode: RibbonChromeMode): void {
    this.chromeMode = mode;
    if (mode === "auto" && !this.resizeObserver && typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => this.applyChrome());
      this.resizeObserver.observe(this.element);
    }
    this.applyChrome();
  }

  private applyChrome(): void {
    const breakpoint = this.options.compactBreakpoint ?? 720;
    const width = this.element.clientWidth || this.element.getBoundingClientRect().width;
    const next: "full" | "compact" =
      this.chromeMode === "auto"
        // Width 0 means the root is not laid out yet (display:none, or measured before
        // first paint). Assume the roomier layout and let the observer correct it.
        ? (width > 0 && width < breakpoint ? "compact" : "full")
        : this.chromeMode;
    if (next === this.density && this.element.dataset.chrome) return;
    this.density = next;
    this.element.dataset.chrome = next;
    // A popover's absolute position is meaningless after a density flip.
    this.closePopovers();
    // A density flip relays out every strip; re-measure without waiting for
    // the observer so the fades never lag a frame behind the new layout.
    this.syncAllStripFades();
  }

  // ── scroll-edge affordances ─────────────────────────────────────────────────

  private syncStripFade(strip: HTMLElement): void {
    const max = strip.scrollWidth - strip.clientWidth;
    const tokens =
      max <= 1 ? "" : `${strip.scrollLeft > 1 ? "l " : ""}${strip.scrollLeft < max - 1 ? "r" : ""}`.trim();
    if (tokens) strip.setAttribute("data-fade", tokens);
    else strip.removeAttribute("data-fade");
  }

  private syncAllStripFades(): void {
    for (const strip of this.strips) this.syncStripFade(strip);
  }

  private wireScrollFades(): void {
    this.strips = Array.from(
      this.element.querySelectorAll<HTMLElement>(".dxr-titlebar, .dxr-tabs, .dxr-panel, .dxr-rail, .dxr-findbar"),
    );
    if (typeof ResizeObserver !== "undefined") {
      // One observer for all strips. It also fires when a panel becomes the
      // active one (display flips from none), which is what keeps a freshly
      // selected tab's fades honest without per-tab wiring.
      this.stripObserver = new ResizeObserver((entries) => {
        for (const entry of entries) this.syncStripFade(entry.target as HTMLElement);
      });
    }
    for (const strip of this.strips) {
      strip.addEventListener("scroll", () => this.syncStripFade(strip), { passive: true });
      this.stripObserver?.observe(strip);
      this.syncStripFade(strip);
    }
  }

  // ── status ──────────────────────────────────────────────────────────────────

  /** Publish the lifecycle on the root, where CSS and host pages can key off it. */
  private setState(state: RibbonState): void {
    this.element.dataset.state = state;
  }

  setStatus(text: string): void {
    const el = this.control("status");
    if (el) el.textContent = text;
    this.options.onStatus?.(text);
  }

  // ── document lifecycle ──────────────────────────────────────────────────────

  get editor(): DocxEditor | null {
    return this.live;
  }

  setExports(exports: DocxEditorExports): void {
    this.exports = exports;
  }

  open(bytes: Uint8Array, name?: string): DocxEditor {
    const { candidateSurface, paginated, started } = this.beginOpen();
    try {
      const candidate = DocxEditor.open(candidateSurface, bytes, this.exports!, this.editorOptions(paginated));
      return this.finishOpen(candidate, paginated, started, name);
    } finally {
      candidateSurface.remove();
    }
  }

  /**
   * {@link open} for a long document: it mounts window by window (`DocxEditor.openAsync`), so the
   * page keeps painting and `onProgress` can show how far it has got. The same editor options as
   * open's, so a document opened either way behaves the same.
   */
  async openAsync(
    bytes: Uint8Array,
    name?: string,
    { onProgress, windowSize }: { onProgress?: (mounted: number, total: number) => void; windowSize?: number } = {},
  ): Promise<DocxEditor> {
    const { candidateSurface, paginated, started } = this.beginOpen();
    try {
      const candidate = await DocxEditor.openAsync(candidateSurface, bytes, this.exports!, {
        ...this.editorOptions(paginated), windowSize, onProgress,
      });
      return this.finishOpen(candidate, paginated, started, name);
    } finally {
      candidateSurface.remove();
    }
  }

  private beginOpen(): { candidateSurface: HTMLElement; paginated: boolean; started: number } {
    if (!this.exports) throw new Error("Docxodus ribbon: WASM exports are not set yet");
    this.history?.beforeOpen();
    const paginated = this.require<HTMLInputElement>("paginated").checked;
    // Render the candidate before retiring the current session. Invalid uploads keep its
    // DOM, selection and undo stack intact. The viewport measures again after attachment.
    const candidateSurface = this.surface.cloneNode(false) as HTMLElement;
    candidateSurface.dataset.view = paginated ? "paginated" : "continuous";
    // (not a second "editor" control while both are in the document)
    candidateSurface.removeAttribute("data-dxr");
    candidateSurface.removeAttribute("data-dxr-surface");
    // Page view measures as it lays pages out, which needs a surface in the document: the
    // candidate renders there, hidden, at the surface's width. Opening straight into page view
    // renders the document once, where opening continuous and switching rendered it twice.
    if (paginated && this.surface.parentElement) {
      Object.assign(candidateSurface.style, {
        position: "absolute", left: "0", top: "0", width: `${this.surface.clientWidth || 800}px`,
        visibility: "hidden", pointerEvents: "none",
      });
      this.surface.parentElement.appendChild(candidateSurface);
    }
    return { candidateSurface, paginated, started: performance.now() };
  }

  /** The options every document's editor gets (open, openAsync). */
  private editorOptions(paginated: boolean): DocxEditorOptions {
    const tracked = this.require<HTMLInputElement>("trackchanges").checked
      ? TrackedChangeMode.RenderInline
      : (this.options.trackedChanges ?? TrackedChangeMode.Accept);
    return {
      cssPrefix: this.options.cssPrefix,
      fabricateClasses: this.options.fabricateClasses,
      editable: this.options.editable,
      scale: this.options.scale,
      columnWidth: this.options.columnWidth,
      fitToWidth: this.options.fitToWidth,
      pasteFormatting: this.options.pasteFormatting,
      resolvePasteImage: this.options.resolvePasteImage,
      onEditFailed: this.options.onEditFailed,
      onEdit: (info) => {
        this.history?.edited();
        this.scheduleDirtyCheck();
        this.options.onEdit?.(info);
        this.scheduleStats();
        this.scheduleRuler();
      },
      onMove: info => { this.history?.edited(); this.scheduleDirtyCheck(); this.options.onMove?.(info); },
      onStoryChange: (which) => this.onStoryChange(which),
      onCommentsChange: (info) => { this.scheduleDirtyCheck(); this.onCommentsChange(info); },
      onFormatPainterChange: (state) => { this.control("formatpainter")?.setAttribute("aria-pressed", String(state !== "off")); },
      paginated,
      headerFooter: this.headerFooter,
      blockDrag: this.options.blockDrag ?? true,
      trackedChanges: tracked,
      showRevisions: this.showRevisions,
      revisionAuthor: this.author,
      revisionIdentities: this.options.revisionIdentities,
      comments: this.options.comments,
      commentAuthor: this.author,
      typingUndo: this.options.typingUndo,
      listDefaults: this.listDefaults,
      headlineCaseLists: this.headlineCaseLists,
      showFormattingMarks: this.showMarks,
      onFormattingMarksChange: (on) => this.marksChanged(on),
      onPictureChange: (picture) => this.pictureChanged(picture),
      commentInitials: this.initials,
    };
  }

  private finishOpen(candidate: DocxEditor, paginated: boolean, started: number, name?: string): DocxEditor {
    this.crossRefDialog?.close();
    this.bookmarkDialog?.close();
    this.live?.close();
    candidate.adoptContainer(this.surface);
    this.surface.dataset.view = paginated ? "paginated" : "continuous";
    this.live = candidate;
    if (name) this.documentName = name;
    this.require('docname').textContent = this.documentName;
    this.closeFindBar();
    this.history?.documentOpened(this.documentName);
    this.markSaved(); // a document just opened is unchanged
    const saveButton = this.control<HTMLButtonElement>("save");
    if (saveButton) saveButton.disabled = false;
    this.require("ribbon").setAttribute("aria-disabled", "false");
    this.setState("ready");
    this.setStatus(`Rendered in ${Math.round(performance.now() - started)} ms`);
    this.populateStyles();
    this.populateFonts(this.documentFonts());
    this.syncPageNumbering();
    this.syncLayoutState();
    this.syncReviewState();
    this.syncZoom();
    this.refreshRailCounts();
    this.refreshRailAnchor();
    this.scheduleStats();
    this.options.onOpen?.(this.live);
    return this.live;
  }

  get isDirty(): boolean {
    this.checkDirty();
    return this.dirty;
  }

  markSaved(): void {
    this.savedVersion = this.sessionVersion();
    this.savedTyped = this.typed;
    this.checkDirty();
  }

  setRecent(entries: RibbonRecentEntry[]): void {
    this.recentEntries = entries.slice();
    const select = this.control<HTMLSelectElement>("recent");
    if (!select) return;
    const doc = select.ownerDocument;
    const options = this.recentEntries.map((entry, i) => {
      const option = doc.createElement("option");
      option.value = String(i);
      option.textContent = entry.name;
      if (entry.title) option.title = entry.title;
      return option;
    });
    select.replaceChildren(select.options[0], ...options);
    select.value = "";
    select.hidden = !this.options.onOpenRecent || options.length === 0;
  }

  setLabels(labels: Record<string, string>): void {
    this.labels = { ...labels };
    // (the ruler writes its tooltips from code, in the language of the moment)
    this.ruler?.relabel();
    const root = this.element;
    const doc = root.ownerDocument;
    const surface = this.surface;
    // A text's English is what it said when first seen, unless the ribbon has written new text
    // there since (the document's name, a count): then that text is the English.
    const english = (now: string, seen: { english: string; written: string } | undefined) =>
      seen && seen.written === now ? seen.english : now;
    const translate = (text: string) => {
      const key = text.trim();
      const to = key ? this.labels[key] : undefined;
      return to ? text.replace(key, to) : text;
    };
    // The document is never touched: the walk skips the editing surface.
    const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
      acceptNode: (node) => node === surface ? NodeFilter.FILTER_REJECT
        : node.nodeName === "STYLE" || node.nodeName === "SCRIPT" ? NodeFilter.FILTER_REJECT
        : NodeFilter.FILTER_ACCEPT,
    });
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node as Text;
        const from = english(text.data, this.englishText.get(text));
        const to = translate(from);
        this.englishText.set(text, { english: from, written: to });
        if (text.data !== to) text.data = to;
        continue;
      }
      const el = node as Element;
      for (const attr of ["title", "aria-label", "placeholder", "label"]) {
        const now = el.getAttribute(attr);
        if (now === null) continue;
        const seen = this.englishAttrs.get(el) ?? {};
        const from = english(now, seen[attr]);
        const to = translate(from);
        seen[attr] = { english: from, written: to };
        this.englishAttrs.set(el, seen);
        if (now !== to) el.setAttribute(attr, to);
      }
    }
  }

  setFileMenu(items: RibbonFileMenuItem[]): void {
    this.fileMenuItems = items.slice();
    this.renderFileMenu();
  }

  openFileMenu(): void {
    const button = this.control("filebutton"), menu = this.control("filemenu");
    if (!button || !menu) return;
    this.closePopovers();
    this.renderFileMenu();
    menu.setAttribute("data-open", "");
    button.setAttribute("aria-expanded", "true");
    this.positionPopover(menu, button);
    menu.querySelector<HTMLButtonElement>("button:not([hidden]):not(:disabled)")?.focus({ preventScroll: true });
  }

  /** The File menu's state and its lists, from the open document, setRecent and fileMenu. */
  private renderFileMenu(): void {
    const menu = this.control("filemenu");
    if (!menu) return;
    const doc = menu.ownerDocument;
    const item = (attrs: Record<string, string>, label: string, shortcut = "", title = "", disabled = false) => {
      const b = doc.createElement("button");
      b.type = "button";
      b.setAttribute("role", "menuitem");
      for (const [k, v] of Object.entries(attrs)) b.setAttribute(k, v);
      const text = doc.createElement("span");
      text.textContent = label;
      const key = doc.createElement("kbd");
      key.textContent = shortcut;
      key.setAttribute("aria-hidden", "true");    // (the item's name is its label; the key goes in aria-keyshortcuts)
      if (shortcut) b.setAttribute("aria-keyshortcuts", shortcut.replace(/\bCtrl\b/g, "Control"));
      b.append(text, key);
      if (title) b.title = title;
      b.disabled = disabled;
      return b;
    };
    const save = menu.querySelector<HTMLButtonElement>('[data-file="save"]');
    if (save) save.disabled = !this.live;
    const saveAs = menu.querySelector<HTMLButtonElement>('[data-file="saveas"]');
    if (saveAs) { saveAs.hidden = !this.options.onSaveAsRequest; saveAs.disabled = !this.live; }

    const recent = this.control("filerecent"), recentItems = this.control("filerecentitems");
    if (recent && recentItems) {
      recentItems.replaceChildren(...this.recentEntries.map((entry, i) =>
        item({ "data-recent": String(i) }, entry.name, "", entry.title ?? "")));
      recent.hidden = !this.options.onOpenRecent || this.recentEntries.length === 0;
    }

    const hostItems = this.control("filehostitems");
    if (hostItems) {
      const nodes: HTMLElement[] = [];
      this.fileMenuItems.forEach((entry, i) => {
        if (entry.separatorBefore || i === 0) {
          const rule = doc.createElement("div");
          rule.className = "dxr-filesep";
          nodes.push(rule);
        }
        nodes.push(item({ "data-host": entry.id }, entry.label, entry.shortcut ?? "", entry.title ?? "", !!entry.disabled));
      });
      hostItems.replaceChildren(...nodes);
    }
  }

  private wireFileMenu(): void {
    const button = this.control("filebutton"), menu = this.control("filemenu");
    if (!button || !menu) return;
    this.keepSelection(button);
    button.addEventListener("click", () => {
      if (menu.hasAttribute("data-open")) this.closePopovers();
      else this.openFileMenu();
    });
    // The document keeps its selection while the menu is used.
    menu.addEventListener("mousedown", (event) => event.preventDefault());
    menu.addEventListener("click", (event) => {
      const chosen = (event.target as HTMLElement).closest<HTMLButtonElement>("button");
      if (!chosen || chosen.disabled) return;
      this.closePopovers();
      const file = chosen.dataset.file;
      if (file === "new") this.requestNew();
      else if (file === "open") this.requestOpen();
      else if (file === "save") this.download();
      else if (file === "saveas") this.options.onSaveAsRequest?.();
      else if (chosen.dataset.recent != null) {
        const entry = this.recentEntries[Number(chosen.dataset.recent)];
        if (entry) this.options.onOpenRecent?.(entry);
      } else if (chosen.dataset.host != null) this.options.onFileMenu?.(chosen.dataset.host);
    });
    // Arrow keys move through the items, Home and End go to the first and last, Escape closes.
    menu.addEventListener("keydown", (event) => {
      const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button"))
        .filter((b) => !b.disabled && b.offsetParent !== null);
      const at = items.indexOf(event.target as HTMLButtonElement);
      let next = -1;
      if (event.key === "ArrowDown") next = (at + 1) % items.length;
      else if (event.key === "ArrowUp") next = (at - 1 + items.length) % items.length;
      else if (event.key === "Home") next = 0;
      else if (event.key === "End") next = items.length - 1;
      else if (event.key === "Escape") { event.preventDefault(); this.closePopovers(); button.focus(); return; }
      if (next >= 0 && items[next]) { event.preventDefault(); items[next].focus(); }
    });
  }

  /** Home > Clipboard: Paste (as Ctrl+V) and its menu of Word's three ways. */
  private wirePaste(): void {
    const button = this.control("paste"), more = this.control("pastemore"), menu = this.control("pastemenu");
    if (!button) return;
    this.keepSelection(button);
    button.addEventListener("click", () => this.pasteClipboard());
    if (!more || !menu) return;
    this.keepSelection(more);
    more.addEventListener("click", () => {
      if (menu.hasAttribute("data-open")) { this.closePopovers(); return; }
      this.closePopovers();
      menu.setAttribute("data-open", "");
      more.setAttribute("aria-expanded", "true");
      this.positionPopover(menu, button);
    });
    // The document keeps its selection while the menu is used.
    menu.addEventListener("mousedown", (event) => event.preventDefault());
    menu.addEventListener("click", (event) => {
      const chosen = (event.target as HTMLElement).closest<HTMLButtonElement>("button[data-paste]");
      if (!chosen) return;
      this.closePopovers();
      this.pasteClipboard(chosen.dataset.paste as PasteFormatting);
    });
    menu.addEventListener("keydown", (event) => {
      const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button"));
      const at = items.indexOf(event.target as HTMLButtonElement);
      let next = -1;
      if (event.key === "ArrowDown") next = (at + 1) % items.length;
      else if (event.key === "ArrowUp") next = (at - 1 + items.length) % items.length;
      else if (event.key === "Escape") { event.preventDefault(); this.closePopovers(); more.focus(); return; }
      if (next >= 0) { event.preventDefault(); items[next].focus(); }
    });
  }

  /** Show/Hide ¶ (see the `showFormattingMarks` option). */
  setShowFormattingMarks(on: boolean): void {
    if (this.live) this.live.setShowFormattingMarks(on);
    else this.marksChanged(on);
  }

  /** The editor turned Show/Hide ¶ on or off: the button shows it, the next document keeps it. */
  private marksChanged(on: boolean): void {
    if (this.showMarks === on) return;
    this.showMarks = on;
    this.control("showmarks")?.setAttribute("aria-pressed", String(on));
    this.options.onFormattingMarksChange?.(on);
  }

  /** Home > Paragraph > ¶. */
  private wireShowMarks(): void {
    const button = this.control("showmarks");
    if (!button) return;
    this.keepSelection(button);
    button.addEventListener("click", () => {
      const on = !this.showMarks;
      this.run("show marks", () => this.setShowFormattingMarks(on), cmd("showMarks", { on }));
    });
  }

  /** Home > Font > Change Case (Aa▾): Word's menu, and Headline Case. */
  private wireChangeCase(): void {
    const button = this.control("changecase"), menu = this.control("casemenu");
    if (!button || !menu) return;
    this.keepSelection(button);
    button.addEventListener("click", () => {
      if (menu.hasAttribute("data-open")) { this.closePopovers(); return; }
      this.closePopovers();
      menu.setAttribute("data-open", "");
      button.setAttribute("aria-expanded", "true");
      this.positionPopover(menu, button);
    });
    menu.addEventListener("mousedown", (event) => event.preventDefault());
    menu.addEventListener("click", (event) => {
      const chosen = (event.target as HTMLElement).closest<HTMLButtonElement>("button[data-case]");
      if (!chosen) return;
      this.closePopovers();
      const mode = chosen.dataset.case as CaseMode;
      const live = this.live;
      if (!live) return;
      this.run("change case", () => { live.changeCase(mode); }, cmd("changeCase", { mode }));
    });
    menu.addEventListener("keydown", (event) => {
      const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button"));
      const at = items.indexOf(event.target as HTMLButtonElement);
      let next = -1;
      if (event.key === "ArrowDown") next = (at + 1) % items.length;
      else if (event.key === "ArrowUp") next = (at - 1 + items.length) % items.length;
      else if (event.key === "Escape") { event.preventDefault(); this.closePopovers(); button.focus(); return; }
      if (next >= 0) { event.preventDefault(); items[next].focus(); }
    });
  }

  /** Headline Case's word lists (see the `headlineCaseLists` option); null for the built-in ones. */
  setHeadlineCaseLists(lists: HeadlineCaseLists | null): void {
    this.headlineCaseLists = lists ?? undefined;
    this.live?.setHeadlineCaseLists(lists);
  }

  /** Home > Clipboard > Format Painter: a click paints once (a second click turns it off), a
   *  double-click keeps painting until Esc or another click. */
  private wireFormatPainter(): void {
    const button = this.control("formatpainter");
    if (!button) return;
    this.keepSelection(button);
    button.addEventListener("click", () => {
      const live = this.live;
      if (!live) return;
      if (live.formatPainterState !== "off") { live.stopFormatPainter(); return; }
      this.run("format painter", () => { live.formatPainter(false); }, cmd("formatPainter"));
    });
    button.addEventListener("dblclick", () => {
      const live = this.live;
      if (!live) return;
      this.run("format painter", () => { live.formatPainter(true); }, cmd("formatPainter", { mode: "sticky" }));
    });
  }

  /** Paste the system clipboard at the caret (the page must be allowed to read it). */
  private pasteClipboard(formatting?: PasteFormatting): void {
    const live = this.live;
    if (!live) return;
    this.run("paste", () => { void live.pasteFromClipboard(formatting); }, cmd("paste", formatting ? { formatting } : {}));
  }

  /** The session's committed version, or null when nothing is open or the engine can't tell. */
  private sessionVersion(): number | null {
    try {
      return this.live?.version ?? null;
    } catch {
      return null; // the session is closed
    }
  }

  /** Recompute `isDirty` after whatever may have changed the document, and report a change of it. */
  private checkDirty(): void {
    if (this.destroyed) return;
    const dirty = !!this.live && (this.typed !== this.savedTyped || this.sessionVersion() !== this.savedVersion);
    if (dirty === this.dirty) return;
    this.dirty = dirty;
    this.options.onDirtyChange?.(dirty);
  }

  /** One check after the current task: a command, a commit and its onEdit come as one change. */
  private scheduleDirtyCheck(): void {
    if (this.dirtyCheckQueued) return;
    this.dirtyCheckQueued = true;
    queueMicrotask(() => {
      this.dirtyCheckQueued = false;
      this.checkDirty();
    });
  }

  openBlank(name = "untitled.docx"): DocxEditor {
    if (!this.exports) throw new Error("Docxodus ribbon: WASM exports are not set yet");
    return this.open(this.exports.DocxSessionBridge.CreateBlankDocx(), name);
  }

  save(): Uint8Array | null {
    return this.live ? this.live.save() : null;
  }

  download(name?: string): void {
    const bytes = this.save();
    if (!bytes) return;
    const filename = name ?? this.documentName ?? "edited.docx";
    if (this.options.onSave) {
      this.options.onSave(bytes, filename); // the host calls markSaved() once the bytes are stored
      return;
    }
    this.markSaved();
    const url = URL.createObjectURL(
      new Blob([bytes as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
    );
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    this.setStatus(`Saved ${filename}`);
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    void this.history?.destroy();
    const doc = this.element.ownerDocument ?? document;
    doc.removeEventListener("selectionchange", this.onSelectionChange);
    doc.removeEventListener("mousedown", this.onDocumentMouseDown);
    this.element.removeEventListener("keydown", this.onKeydown);
    this.element.removeEventListener("input", this.onInput);
    this.resizeObserver?.disconnect();
    this.resizeObserver = null;
    this.stripObserver?.disconnect();
    this.stripObserver = null;
    if (this.selectionFrame != null) cancelAnimationFrame(this.selectionFrame);
    this.selectionFrame = null;
    if (this.rulerFrame != null) cancelAnimationFrame(this.rulerFrame);
    this.rulerFrame = null;
    this.rulerObserver?.disconnect();
    this.ruler?.destroy();
    this.focusRing?.remove();
    this.bookmarkLayer?.remove();
    if (this.statsTimer != null) clearTimeout(this.statsTimer);
    this.stopRotation();
    try {
      this.live?.close();
    } catch {
      /* already closed */
    }
    this.live = null;
    this.element.remove();
  }

  // ── command plumbing ────────────────────────────────────────────────────────

  /**
   * Run one ribbon command and report its real cost on the rail.
   *
   * Every control routes through here, which is what makes the rail's "last op"
   * an honest measurement rather than a label the surface makes up.
   */
  private run(label: string, fn: () => void, command?: CommandOf): void {
    if (!this.live) return;
    const started = performance.now();
    let ok = false;
    try {
      fn();
      ok = true;
    } finally {
      const ms = performance.now() - started;
      const el = this.control("railOp");
      if (el) el.textContent = `${label} ${ms >= 1000 ? `${(ms / 1000).toFixed(2)} s` : `${Math.round(ms)} ms`}`;
      let done: RibbonCommand | undefined;
      if (ok) {
        try { done = typeof command === "function" ? command() : command; } catch { done = undefined; }
      }
      this.options.onCommand?.(label, ms, done);
      this.history?.edited();
      // An engine that has no session version can't say whether the command changed anything:
      // count it as a change, so the document is never taken for saved when it may not be.
      if (this.live && this.sessionVersion() === null) this.typed++;
      this.scheduleDirtyCheck();
      this.refreshRailCounts();
      this.refreshRailAnchor();
      this.scheduleRuler();
      this.scheduleStats();
    }
  }

  /** Format controls must not steal the document selection they are about to act on. */
  private keepSelection(el: HTMLElement): void {
    el.addEventListener("mousedown", (event) => event.preventDefault());
  }

  /** The Paragraph dialog and its launchers (Home and Layout). Its OK is one `paragraphFormat`. */
  private wireParagraphDialog(): void {
    this.paragraphDialog = new ParagraphDialog(this.require<HTMLDialogElement>("paradialog"), {
      read: () => this.live?.paragraphFormats() ?? [],
      apply: (op) => this.run("paragraph", () => this.live!.setParagraphFormat(op), cmd("paragraphFormat", { ...op })),
      openTabs: () => { this.openTabsDialog(); },
    }, this.options.measurementUnit ?? "in");
    for (const name of ["paralaunch", "paralaunchlayout"]) {
      const launcher = this.control(name);
      if (!launcher) continue;
      this.keepSelection(launcher);
      launcher.addEventListener("click", () => this.openParagraphDialog());
    }
  }

  /** The List Levels dialog, the Numbering menu's list commands and the 1.1 button. */
  private wireListDialog(): void {
    this.listDefaults = this.options.listDefaults?.map((l) => ({ ...l }));
    this.headlineCaseLists = this.options.headlineCaseLists;
    this.showMarks = !!this.options.showFormattingMarks;
    this.control("showmarks")?.setAttribute("aria-pressed", String(this.showMarks));
    this.listDialog = new ListDialog(this.require<HTMLDialogElement>("listdialog"), {
      read: () => this.live?.listLevels() ?? null,
      defaults: () => this.listDefaults ?? null,
      apply: (levels) => this.run("list levels", () => this.live!.setListLevels(levels), cmd("setListLevels", { levels })),
      keepDefaults: this.options.onListDefaults
        ? (levels) => { this.setListDefaults(levels); this.options.onListDefaults!(levels); }
        : undefined,
    }, this.options.measurementUnit ?? "in");
    const legal = this.control("legallist");
    if (legal) {
      this.keepSelection(legal);
      legal.addEventListener("click", () => {
        if (this.live) this.run("legal numbering", () => this.live!.setLegalNumbering("top"), cmd("legalNumbering", { periods: "top" }));
      });
    }
  }

  // ── Picture Format (contextual), the picture menu and dialog ─────────────────

  /** The editor selected a picture, dropped the selection, or changed the picture. */
  private pictureChanged(info: PictureInfo | null): void {
    const tab = this.element.querySelector<HTMLElement>('.dxr-tab[data-tab="picture"]');
    if (tab) {
      tab.hidden = !info;
      if (!info && tab.getAttribute("aria-selected") === "true") this.selectTab("home");
    }
    if (!info) {
      this.control("picturemenu")?.removeAttribute("data-open");
      return;
    }
    const wrap = this.require<HTMLSelectElement>("picwrap");
    wrap.value = info.wrap ?? "";
    wrap.disabled = !info.can.wrap;
    const doc = this.element.ownerDocument;
    for (const [name, points] of [["picheight", info.heightPoints], ["picwidth", info.widthPoints]] as const) {
      const input = this.require<HTMLInputElement>(name);
      if (doc.activeElement !== input) input.value = points === null ? "" : formatLength(points * 20, this.pictureUnit);
      input.disabled = !info.can.size;
    }
    this.require<HTMLButtonElement>("picreset").disabled = !info.can.size || !info.originalWidthPoints;
    this.require<HTMLButtonElement>("picposition").disabled = !info.can.layout && !info.can.wrap;
    this.require<HTMLButtonElement>("picalt").disabled = !info.can.altText;
    this.require<HTMLButtonElement>("picdelete").disabled = !info.can.remove;
    const replace = this.require<HTMLInputElement>("picreplace");
    replace.disabled = !info.can.replace;
    replace.closest("label")?.toggleAttribute("aria-disabled", !info.can.replace);
    const menu = this.control("picturemenu");
    if (menu) {
      for (const item of Array.from(menu.querySelectorAll<HTMLButtonElement>("[data-pm-wrap]"))) {
        item.setAttribute("aria-checked", String(item.dataset.pmWrap === info.wrap));
        item.disabled = !info.can.wrap;
      }
      menu.querySelector<HTMLButtonElement>('[data-pm="alt"]')!.disabled = !info.can.altText;
      menu.querySelector<HTMLButtonElement>('[data-pm="replace"]')!.disabled = !info.can.replace;
      menu.querySelector<HTMLButtonElement>('[data-pm="delete"]')!.disabled = !info.can.remove;
    }
  }

  /** Run one picture change as a ribbon command; a refusal says so on the status line. */
  private pictureCommand(label: string, change: (live: DocxEditor) => boolean, command: RibbonCommand): void {
    const live = this.live;
    if (!live) return;
    let refused = false;
    this.run(label, () => {
      refused = !change(live);
      if (refused) this.setStatus(this.translateText(RIBBON_STATUS_TEXTS.pictureRefused));
    }, () => refused ? undefined : command);
  }

  private applyPictureDialog(result: PictureDialogResult): void {
    // Wrapping and position first (a picture made to float keeps its size), then the size, then the words.
    if (result.wrap) this.pictureCommand("picture wrap", (l) => l.setPictureWrap(result.wrap!), cmd("pictureWrap", { wrap: result.wrap }));
    if (result.layout) {
      const layout = result.layout;
      this.pictureCommand("picture layout", (l) => l.setPictureLayout(layout), cmd("pictureLayout", { layout: { ...layout } as unknown as Record<string, unknown> }));
    }
    if (result.size) {
      const { widthPoints, heightPoints } = result.size;
      this.pictureCommand("picture size", (l) => l.setPictureSize(widthPoints, heightPoints, false),
        cmd("pictureSize", { width: widthPoints, height: heightPoints, lockAspect: false }));
    }
    if (result.altText) {
      const { altText, title } = result.altText;
      this.pictureCommand("alt text", (l) => l.setPictureAltText(altText, title), cmd("pictureAltText", { altText, title }));
    }
  }

  openPictureDialog(tab: "size" | "layout" | "alt" = "size"): boolean {
    if (!this.live || this.pictureDialog.open || this.columnsDialog.open || this.listDialog.open || this.paragraphDialog.open) return false;
    this.closePopovers();
    return this.pictureDialog.show(tab as PictureDialogTab);
  }

  openTableDialog(tab: "table" | "row" | "column" | "cell" | "alt" = "table"): boolean {
    if (!this.live || this.tableDialog.open || this.pictureDialog.open || this.columnsDialog.open || this.listDialog.open
      || this.paragraphDialog.open) return false;
    this.closePopovers();
    return this.tableDialog.show(tab as TableDialogTab);
  }

  private applyTableDialog(result: TableDialogResult): void {
    const args: Record<string, unknown> = {};
    if (result.table) args.table = { ...result.table };
    if (result.row) args.row = { ...result.row };
    if (result.columnWidthTwips !== undefined) args.columnWidthTwips = result.columnWidthTwips;
    if (result.cell) args.cell = { ...result.cell };
    this.run("table properties", () => this.live!.applyTableProperties(result), cmd("applyTableProperties", args as EditorCommandArgs));
  }

  private wireTable(): void {
    this.tableDialog = new TableDialog(this.require<HTMLDialogElement>("tabledialog"), {
      read: () => this.live?.tableProperties() ?? null,
      apply: (result) => this.applyTableDialog(result),
    }, this.options.measurementUnit ?? "in");
    this.require("tblprops").addEventListener("click", () => this.openTableDialog("table"));
    for (const name of ["tblheight", "tblwidth"] as const) {
      const input = this.require<HTMLInputElement>(name);
      input.addEventListener("change", () => {
        const text = input.value.trim();
        const twips = text === "" && name === "tblheight" ? 0 : parseLength(text, this.pictureUnit);
        if (twips === null || twips < 0 || (twips === 0 && name === "tblwidth")) {
          this.tableChanged();
          return;
        }
        if (name === "tblheight") {
          this.run("row height", () => this.live!.setTableRowOptions({ heightTwips: twips, heightRule: "atLeast" }),
            cmd("rowHeight", { twips, rule: "atLeast" }));
        } else {
          this.run("column width", () => this.live!.setColumnWidth(twips), cmd("columnWidth", { twips }));
        }
      });
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") { event.preventDefault(); input.dispatchEvent(new Event("change")); }
      });
    }
  }

  /** The Cell Size boxes and the Alignment buttons, from the caret's table. */
  private tableChanged(): void {
    const info = this.live?.tableProperties() ?? null;
    const height = this.control<HTMLInputElement>("tblheight");
    const width = this.control<HTMLInputElement>("tblwidth");
    if (height && document.activeElement !== height)
      height.value = info?.rowHeightTwips ? formatLength(info.rowHeightTwips, this.pictureUnit) : "";
    if (width && document.activeElement !== width) {
      const span = info ? info.columnWidthsTwips.slice(info.columnIndex, info.columnIndex + info.columnSpan) : [];
      const twips = span.reduce((sum, w) => sum + w, 0);
      width.value = twips > 0 ? formatLength(twips, this.pictureUnit) : "";
    }
    const valign = info?.verticalAlignment === "center" || info?.verticalAlignment === "bottom" ? info.verticalAlignment : "top";
    for (const [tt, v] of [["alignTop", "top"], ["alignCenter", "center"], ["alignBottom", "bottom"]] as const)
      this.element.querySelector<HTMLElement>(`.dxr-panel[data-panel="table"] button[data-tt="${tt}"]`)
        ?.classList.toggle("dxr-on", !!info && valign === v);
  }

  private wirePicture(): void {
    this.pictureUnit = this.options.measurementUnit ?? "in";
    this.pictureDialog = new PictureDialog(this.require<HTMLDialogElement>("picturedialog"), {
      read: () => this.live?.selectedPicture() ?? null,
      apply: (result) => this.applyPictureDialog(result),
    }, this.pictureUnit);

    const wrap = this.require<HTMLSelectElement>("picwrap");
    wrap.addEventListener("change", () => {
      const value = wrap.value as PictureWrap;
      if (value) this.pictureCommand("picture wrap", (l) => l.setPictureWrap(value), cmd("pictureWrap", { wrap: value }));
    });
    for (const [name, side] of [["picheight", "height"], ["picwidth", "width"]] as const) {
      const input = this.require<HTMLInputElement>(name);
      input.addEventListener("change", () => {
        const twips = parseLength(input.value, this.pictureUnit);
        if (twips === null || twips <= 0) {
          this.pictureChanged(this.live?.selectedPicture() ?? null);
          return;
        }
        const points = Math.round((twips / 20) * 100) / 100;
        this.pictureCommand("picture size",
          (l) => side === "height" ? l.setPictureSize(undefined, points) : l.setPictureSize(points, undefined),
          cmd("pictureSize", { [side]: points }));
      });
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") { event.preventDefault(); input.dispatchEvent(new Event("change")); }
      });
    }
    this.require("picreset").addEventListener("click", () =>
      this.pictureCommand("reset picture size", (l) => l.resetPictureSize(), cmd("resetPictureSize")));
    this.require("picmore").addEventListener("click", () => this.openPictureDialog("size"));
    this.require("picposition").addEventListener("click", () => this.openPictureDialog("layout"));
    this.require("picalt").addEventListener("click", () => this.openPictureDialog("alt"));
    this.require("picdelete").addEventListener("click", () =>
      this.pictureCommand("delete picture", (l) => l.removePicture(), cmd("deletePicture")));
    const replace = this.require<HTMLInputElement>("picreplace");
    replace.addEventListener("change", async () => {
      const chosen = replace.files?.[0];
      replace.value = "";
      const live = this.live;
      if (!chosen || !live) return;
      const ok = await live.replacePictureFile(chosen);
      if (ok) { this.history?.edited(); this.scheduleDirtyCheck(); }
      this.setStatus(ok ? `${this.translateText("Changed the picture to")} ${chosen.name}` : this.translateText(RIBBON_STATUS_TEXTS.pictureRefused));
    });

    // The right-click menu on a picture (the editor selected it on the button press).
    const menu = this.require("picturemenu");
    const surface = this.require("editor");
    surface.addEventListener("contextmenu", (event) => {
      const target = event.target instanceof Element ? event.target : null;
      if (!target?.closest("img, [data-docx-vector-frame]") || !this.live?.selectedPicture()) return;
      event.preventDefault();
      this.closePopovers();
      this.pictureChanged(this.live.selectedPicture());
      menu.setAttribute("data-open", "");
      const view = this.element.ownerDocument.defaultView;
      const width = view?.innerWidth ?? 1024, height = view?.innerHeight ?? 768;
      const box = menu.getBoundingClientRect();
      menu.style.left = `${Math.max(4, Math.min(event.clientX, width - box.width - 4))}px`;
      menu.style.top = `${Math.max(4, Math.min(event.clientY, height - box.height - 4))}px`;
      menu.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus();
    });
    menu.addEventListener("mousedown", (event) => event.preventDefault());
    menu.addEventListener("keydown", (event) => {
      if (event.key === "Escape") { event.preventDefault(); menu.removeAttribute("data-open"); return; }
      if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return;
      event.preventDefault();
      const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button:not(:disabled)"));
      const at = items.indexOf(event.target as HTMLButtonElement);
      items[(at + (event.key === "ArrowDown" ? 1 : items.length - 1)) % items.length]?.focus();
    });
    menu.addEventListener("click", (event) => {
      const item = (event.target as HTMLElement).closest<HTMLButtonElement>("button");
      if (!item || item.disabled) return;
      menu.removeAttribute("data-open");
      const wrapTo = item.dataset.pmWrap as PictureWrap | undefined;
      if (wrapTo) {
        this.pictureCommand("picture wrap", (l) => l.setPictureWrap(wrapTo), cmd("pictureWrap", { wrap: wrapTo }));
        return;
      }
      switch (item.dataset.pm) {
        case "size": this.openPictureDialog("size"); break;
        case "alt": this.openPictureDialog("alt"); break;
        case "replace": replace.click(); break;
        case "delete": this.pictureCommand("delete picture", (l) => l.removePicture(), cmd("deletePicture")); break;
      }
    });
  }

  /** Layout > Page setup > Columns (One, Two, Three, More Columns…) and its dialog. */
  private wireColumns(): void {
    this.columnsDialog = new ColumnsDialog(this.require<HTMLDialogElement>("columnsdialog"), {
      read: () => this.live?.sectionInfo() ?? null,
      hasSelection: () => !!this.columnsRange && !this.columnsRange.collapsed,
      apply: (op, applyTo) => {
        const range = this.columnsRange;
        this.columnsRange = null;
        if (range) {
          const sel = window.getSelection();
          sel?.removeAllRanges();
          sel?.addRange(range);
        }
        this.applyColumns(op, applyTo);
      },
    }, this.options.measurementUnit ?? "in");
    const menu = this.require<HTMLSelectElement>("columns");
    menu.addEventListener("change", () => {
      const v = menu.value;
      menu.value = "";
      if (!this.live || !v) return;
      if (v === "more") { this.openColumnsDialog(); return; }
      this.applyColumns({ count: Number(v) });
    });
  }

  /** Set columns as the Columns list or dialog asked; a refusal says why on the status line and
   *  is not recorded. */
  private applyColumns(op: ColumnsOp, applyTo?: ColumnsApplyTo): void {
    let refused = false;
    this.run("columns", () => {
      refused = !this.live!.setColumns(op, applyTo);
      if (refused) this.setStatus(this.translateText(RIBBON_STATUS_TEXTS.columnsRefused));
    }, () => refused ? undefined : cmd("columns", { ...op, ...(applyTo ? { applyTo } : {}) }));
    this.syncLayoutState();
  }

  private wireLineNumbers(): void {
    this.lineNumbersDialog = new LineNumbersDialog(this.require<HTMLDialogElement>("linenumbersdialog"), {
      read: () => this.live?.sectionInfo() ?? null,
      apply: (lineNumbering, applyTo) => this.applyLineNumbering(lineNumbering, applyTo),
    }, this.options.measurementUnit ?? "in");
    const menu = this.require<HTMLSelectElement>("linenumbers");
    menu.addEventListener("change", () => {
      const v = menu.value;
      menu.value = "";
      if (!this.live || !v) return;
      if (v === "more") { this.openLineNumbersDialog(); return; }
      if (v === "suppress") {
        this.run("suppress line numbers", () => this.live!.setSuppressLineNumbers(true), cmd("suppressLineNumbers", { on: true }));
        return;
      }
      // The menu numbers the whole document: columns on selected text and continuous breaks make
      // sections nobody sees, and numbering only the caret's one looked arbitrary. (The options
      // dialog still offers This section.)
      if (v === "none") { this.applyLineNumbering({ countBy: 0 }, "document"); return; }
      // As Word, keep the count-by, start and distance and change only when numbering restarts.
      const current = this.live.lineNumbering();
      this.applyLineNumbering({ countBy: current?.countBy ?? 1, start: current?.start ?? 1,
        ...(current?.distanceTwips !== undefined ? { distanceTwips: current.distanceTwips } : {}),
        restart: v as "newPage" | "newSection" | "continuous" }, "document");
    });
  }

  /** Set line numbering as the menu or dialog asked; in the continuous view, say where they show. */
  private applyLineNumbering(lineNumbering: LineNumbering, applyTo: "section" | "document" = "section"): void {
    this.run("line numbers", () => this.live!.setLineNumbering(lineNumbering, applyTo),
      cmd("lineNumbers", { ...lineNumbering, ...(applyTo === "document" ? { applyTo } : {}) }));
    if (lineNumbering.countBy > 0 && !this.require<HTMLInputElement>("paginated").checked)
      this.setStatus(this.translateText(RIBBON_STATUS_TEXTS.lineNumbersPageView));
  }

  private wireBookmarks(): void {
    this.bookmarkDialog = new BookmarkDialog(this.require<HTMLDialogElement>("bookmarkdialog"), {
      list: (hidden) => this.live?.bookmarks(hidden) ?? [],
      add: (name) => {
        let result: { ok: boolean; code?: string } = { ok: false };
        // (the dialog is not modal: the bookmark goes on the document's selection as it is now)
        this.run("bookmark", () => { result = this.live!.addBookmark(name, this.live!.selectionDocumentRange()); },
          () => result.ok ? cmd("addBookmark", { name }) : undefined);
        return result;
      },
      remove: (name) => {
        let result: { ok: boolean; code?: string } = { ok: false };
        this.run("delete bookmark", () => { result = this.live!.removeBookmark(name); },
          () => result.ok ? cmd("deleteBookmark", { name }) : undefined);
        if (result.code === "bookmark_in_use") this.setStatus(this.translateText(RIBBON_STATUS_TEXTS.bookmarkInUse));
        return result;
      },
      goTo: (name) => {
        const ok = !!this.live?.goToBookmark(name);
        if (ok) this.options.onCommand?.("go to bookmark", 0, cmd("goToBookmark", { name }));
        this.scheduleRuler();
        return ok;
      },
      closed: () => {
        const live = this.live;
        const at = live?.selectionDocumentRange();
        if (live && at) live.placeCaretAt(at);
        this.scheduleRuler();
      },
    });
    const button = this.require("bookmark");
    this.keepSelection(button);
    button.addEventListener("click", () => { this.openBookmarkDialog(); });
    const box = this.require<HTMLInputElement>("showbookmarks");
    // (a way of looking, as the Ruler box: not a command a recording replays)
    box.addEventListener("change", () => this.setShowBookmarks(box.checked));
    this.wireCrossReferences();
    this.wireFields();
  }

  /** Update Table (Insert), F9 (see handleShortcut) and the right-click Update Field on a field. */
  private wireFields(): void {
    const button = this.require("updatetoc");
    this.keepSelection(button);
    button.addEventListener("click", () => { this.updateFields("tables"); });
    const menu = this.require("fieldmenu");
    const surface = this.require("editor");
    surface.addEventListener("contextmenu", (event) => {
      const target = event.target instanceof Element ? event.target : null;
      if (!this.live || !target || target.closest("img, [data-docx-vector-frame]")) return;
      // (on a field, or anywhere in a table of contents entry: it links to its heading's _Toc bookmark)
      const block = target.closest<HTMLElement>('[data-anchor][contenteditable="true"]');
      if (!target.closest("[data-field]") && !block?.querySelector('a[href^="#_Toc"]')) return;
      event.preventDefault();
      this.closePopovers();
      menu.setAttribute("data-open", "");
      const view = this.element.ownerDocument.defaultView;
      const width = view?.innerWidth ?? 1024, height = view?.innerHeight ?? 768;
      const box = menu.getBoundingClientRect();
      menu.style.left = `${Math.max(4, Math.min(event.clientX, width - box.width - 4))}px`;
      menu.style.top = `${Math.max(4, Math.min(event.clientY, height - box.height - 4))}px`;
      menu.querySelector<HTMLButtonElement>("button")?.focus();
    });
    menu.addEventListener("mousedown", (event) => event.preventDefault());
    menu.addEventListener("keydown", (event) => {
      if (event.key === "Escape") { event.preventDefault(); menu.removeAttribute("data-open"); }
    });
    menu.addEventListener("click", (event) => {
      const item = (event.target as HTMLElement).closest<HTMLButtonElement>("button");
      if (!item) return;
      menu.removeAttribute("data-open");
      if (item.dataset.fm === "update") this.updateFields("selection");
    });
    this.element.ownerDocument.addEventListener("mousedown", (event) => {
      if (menu.hasAttribute("data-open") && !(event.target instanceof Node && menu.contains(event.target))) menu.removeAttribute("data-open");
    }, true);
  }

  updateFields(scope: "selection" | "all" | "tables" = "selection"): boolean {
    const live = this.live;
    if (!live) return false;
    let ok = false;
    this.run(scope === "tables" ? "update table" : "update fields", () => { ok = live.updateFields(scope); },
      () => ok ? cmd("updateFields", scope === "selection" ? {} : { scope }) : undefined);
    this.setStatus(this.translateText(ok ? RIBBON_STATUS_TEXTS.fieldsUpdated : RIBBON_STATUS_TEXTS.fieldsRefused));
    return ok;
  }

  private wireCrossReferences(): void {
    this.crossRefDialog = new CrossReferenceDialog(this.require<HTMLDialogElement>("crossrefdialog"), {
      targets: (type) => {
        const live = this.live;
        if (!live) return [];
        if (type === "bookmark") return live.bookmarks(false).map((b) => ({ key: b.name, label: b.name }));
        if (type === "numbered") return live.numberedItems().map((n) => ({ key: n.anchorId, label: `${n.number} ${n.text}`.trim() }));
        return live.outline().map((h) => {
          const number = live.numberOf(h.anchorId);
          return { key: h.anchorId, label: number ? `${number} ${h.text}` : h.text, level: h.level };
        });
      },
      insert: (type, key, insert, hyperlink, position) => {
        let result: { ok: boolean; code?: string; after?: DocumentRange } = { ok: false };
        const live = this.live!;
        // (the recorded command names the target as a reader would: a heading's text, an item's number and text)
        const target = type === "bookmark" ? key
          : type === "heading" ? live.outline().find((h) => h.anchorId === key)?.text ?? key
          : (() => { const n = live.numberedItems().find((i) => i.anchorId === key); return n ? `${n.number} ${n.text}`.trim() : key; })();
        const options = crossReferenceOptionsFor(insert, hyperlink, position);
        this.run("cross-reference", () => {
          // (the dialog is not modal: the reference goes at the document's caret as it is now)
          result = live.insertCrossReference(live.selectionDocumentRange(), type === "bookmark" ? { bookmark: key } : { anchorId: key }, options);
        }, () => result.ok ? cmd("crossReference", { type, target, insert, ...(hyperlink ? { hyperlink } : {}), ...(position && options.includePosition ? { position } : {}) }) : undefined);
        return result;
      },
      closed: () => {
        const live = this.live;
        const at = live?.selectionDocumentRange();
        if (live && at) live.placeCaretAt(at);
        this.scheduleRuler();
      },
    });
    const button = this.require("crossref");
    this.keepSelection(button);
    button.addEventListener("click", () => { this.openCrossReferenceDialog(); });
  }

  openCrossReferenceDialog(): boolean {
    if (!this.live || this.paragraphDialog.open || this.listDialog.open) return false;
    this.closePopovers();
    this.crossRefDialog.show();
    return true;
  }

  private wireNavPane(): void {
    this.navPane = new NavPane(this.require("navpane"), this.require("navmenu"), {
      outline: () => this.live?.outline() ?? [],
      current: () => this.live?.headingAtCaret() ?? null,
      goTo: (entry) => {
        if (!this.live?.goToHeading(entry.anchorId)) return;
        this.options.onCommand?.("go to heading", 0, cmd("goToHeading", { heading: entry.text }));
        this.scheduleRuler();
      },
      move: (entry, before) => this.moveHeading(entry, before),
      command: (entry, command) => this.headingCommand(entry, command),
      close: () => this.setNavigationPaneVisible(false, true),
    });
    const box = this.require<HTMLInputElement>("shownav");
    box.addEventListener("change", () => this.setNavigationPaneVisible(box.checked, true));
    if (this.options.navigationPane) this.setNavigationPaneVisible(true);
  }

  setNavigationPaneVisible(on: boolean, report = false): void {
    const box = this.control<HTMLInputElement>("shownav");
    if (box) box.checked = on;
    const was = this.navPane.visible;
    this.navPane.setVisible(on);
    this.scheduleRuler();
    if (report && was !== on) this.options.onNavigationPaneChange?.(on);
  }

  /** A heading dragged in the Navigation Pane: before heading `before`, or to the end of the
   *  document. The command records the headings by their text. */
  private moveHeading(entry: OutlineEntry, before: OutlineEntry | null): void {
    const outline = this.live?.outline() ?? [];
    const last = outline[outline.length - 1];
    const target = before ?? last;
    if (!target) return;
    let result: { ok: boolean; code?: string } = { ok: false };
    this.run("move heading", () => {
      result = before ? this.live!.moveHeading(entry.anchorId, before.anchorId, "before")
        : this.live!.moveHeading(entry.anchorId, last.lastAnchorId, "after");
    }, () => result.ok ? cmd("moveHeading", { heading: entry.text, to: target.text, position: before ? "before" : "after" }) : undefined);
    if (!result.ok) this.setStatus(this.translateText(RIBBON_STATUS_TEXTS.headingMoveRefused));
  }

  /** The Navigation Pane's right-click menu (Expand, Collapse and Show Heading Levels stay in the pane). */
  private headingCommand(entry: OutlineEntry, command: NavPaneCommand): void {
    const heading = entry.text;
    let ok = false;
    const say = () => { if (!ok) this.setStatus(this.translateText(RIBBON_STATUS_TEXTS.headingRefused)); };
    switch (command) {
      case "promote":
      case "demote": {
        const level = entry.level + (command === "promote" ? -1 : 1);
        this.run(command, () => { ok = this.live!.setHeadingLevel(entry.anchorId, level).ok; },
          () => ok ? cmd("headingLevel", { heading, level }) : undefined);
        say();
        return;
      }
      case "before":
      case "after":
      case "sub":
        this.run("new heading", () => { ok = this.live!.insertHeading(entry.anchorId, command); },
          () => ok ? cmd("newHeading", { heading, where: command }) : undefined);
        say();
        return;
      case "delete":
        this.run("delete heading", () => { ok = this.live!.deleteHeadingSection(entry.anchorId); },
          () => ok ? cmd("deleteHeading", { heading }) : undefined);
        say();
        return;
      case "select":
        ok = !!this.live?.selectHeadingSection(entry.anchorId);
        if (ok) this.options.onCommand?.("select heading", 0, cmd("selectHeading", { heading }));
        this.scheduleRuler();
        return;
    }
  }

  openBookmarkDialog(): boolean {
    if (!this.live || this.paragraphDialog.open || this.listDialog.open) return false;
    this.closePopovers();
    this.bookmarkDialog.show();
    return true;
  }

  setShowBookmarks(on: boolean): void {
    this.showBookmarks = on;
    const box = this.control<HTMLInputElement>("showbookmarks");
    if (box) box.checked = on;
    this.scheduleRuler();
  }

  /** Word's Show bookmarks: a gray bracket where each bookmark starts and ends, an I-beam for an
   *  empty one. Drawn over the document (its text and offsets stay as they are), clipped to the
   *  scroll area, again after every edit, scroll and resize. */
  private drawBookmarks(): void {
    const layer = this.bookmarkLayer;
    if (!this.showBookmarks || !this.live) {
      if (layer) layer.hidden = true;
      return;
    }
    const doc = this.element.ownerDocument;
    let host = layer;
    if (!host) {
      host = doc.createElement("div");
      host.className = "dxr-bkmarks";
      host.dataset.dxr = "bookmarkmarks";
      this.element.appendChild(host);
      this.bookmarkLayer = host;
    }
    host.hidden = false;
    const area = this.element.querySelector<HTMLElement>("[data-dxr-scroll]")?.getBoundingClientRect();
    host.style.clipPath = area
      ? `inset(${Math.max(0, area.top)}px ${Math.max(0, (doc.defaultView?.innerWidth ?? area.right) - area.right)}px ${Math.max(0, (doc.defaultView?.innerHeight ?? area.bottom) - area.bottom)}px ${Math.max(0, area.left)}px)`
      : "none";
    const marks: HTMLElement[] = [];
    const add = (side: "start" | "end" | "empty", rect: DOMRect, name: string) => {
      const mark = doc.createElement("span");
      mark.className = "dxr-bkmark";
      mark.dataset.side = side;
      mark.dataset.name = name;
      const left = side === "start" ? rect.left - 1 : side === "end" ? rect.left - 3 : rect.left - 2.5;
      Object.assign(mark.style, { left: `${left}px`, top: `${rect.top}px`, height: `${Math.max(rect.height, 8)}px` });
      marks.push(mark);
    };
    for (const mark of this.live.bookmarkPositions()) {
      if (mark.empty) { add("empty", mark.start, mark.name); continue; }
      add("start", mark.start, mark.name);
      add("end", mark.end, mark.name);
    }
    host.replaceChildren(...marks);
  }

  private wireRuler(): void {
    this.rulerUnit = this.options.measurementUnit ?? "in";
    this.tabsDialog = new TabsDialog(this.require<HTMLDialogElement>("tabsdialog"), {
      read: () => {
        this.live?.focusRulerParagraph();
        const state = this.live?.rulerState();
        return state ? { stops: state.tabs, styleStops: [], defaultTabStopTwips: state.defaultTabStopTwips } : null;
      },
      apply: (stops, defaultTwips) => {
        this.run("tabs", () => { this.live!.focusRulerParagraph(); this.live!.setTabStops(stops); }, cmd("tabStops", { stops }));
        if (defaultTwips !== null)
          this.run("default tab stops", () => this.live!.setDefaultTabStop(defaultTwips), cmd("defaultTabStop", { twips: defaultTwips }));
      },
    }, this.rulerUnit);
    const scroll = this.element.querySelector<HTMLElement>("[data-dxr-scroll]");
    this.ruler = new Ruler(this.require("ruler"), {
      state: () => this.live?.rulerState() ?? null,
      unit: () => this.rulerUnit,
      t: (text) => this.translateText(text),
      setIndents: (op) => this.run("ruler indent", () => { this.live!.focusRulerParagraph(); this.live!.setParagraphFormat(op); }, cmd("paragraphFormat", { ...op })),
      setTabStops: (stops) => this.run("ruler tab", () => { this.live!.focusRulerParagraph(); this.live!.setTabStops(stops); }, cmd("tabStops", { stops })),
      setColumnWidths: (widths) => this.run("ruler column", () => { this.live!.focusRulerParagraph(); this.live!.setTableColumnWidths(widths); }),
      setMargins: (op) => this.run("ruler margin", () => { this.live!.focusRulerParagraph(); this.live!.setPageSetup(op); }, cmd("pageSetup", { ...op })),
      openTabsDialog: () => { this.openTabsDialog(); },
      guideArea: () => scroll,
    });
    const box = this.require<HTMLInputElement>("showruler");
    box.addEventListener("change", () => this.setRulerVisible(box.checked));
    scroll?.addEventListener("scroll", () => this.scheduleRuler(), { passive: true });
    if (typeof ResizeObserver !== "undefined" && scroll) {
      this.rulerObserver = new ResizeObserver(() => this.scheduleRuler());
      this.rulerObserver.observe(scroll);
      // (a view switch or a page layout changes the document's size, not the scroll area's)
      this.rulerObserver.observe(this.surface);
    }
  }

  setRulerVisible(on: boolean): void {
    const box = this.control<HTMLInputElement>("showruler");
    if (box) box.checked = on;
    this.ruler.element.hidden = !on;
    this.scheduleRuler();
  }

  /** Redraw the ruler and the hanging paragraph's focus ring on the next frame (caret moves,
   *  edits, scrolling, zoom all ask). */
  private scheduleRuler(): void {
    if (!this.ruler || this.rulerFrame != null || this.destroyed) return;
    this.rulerFrame = requestAnimationFrame(() => {
      this.rulerFrame = null;
      this.syncFocusRing();
      this.drawBookmarks();
      if (this.navPane?.visible) this.navPane.refresh();
      if (!this.ruler.element.hidden) this.ruler.refresh();
    });
  }

  /** The focus outline of a paragraph whose first line hangs left of its box: around the box and
   *  that line (a CSS outline follows the box only, and the ribbon's CSS hides it for these). */
  private syncFocusRing(): void {
    const doc = this.element.ownerDocument;
    const active = doc.activeElement as HTMLElement | null;
    const block = active?.closest?.<HTMLElement>('[data-anchor][contenteditable="true"]') ?? null;
    const hang = block && this.surface.contains(block) && typeof getComputedStyle === "function"
      ? parseFloat(getComputedStyle(block).textIndent) : 0;
    if (!block || !(hang < 0)) {
      if (this.focusRing) this.focusRing.hidden = true;
      return;
    }
    if (!this.focusRing) {
      this.focusRing = doc.createElement("div");
      this.focusRing.className = "dxr-focusring";
      this.focusRing.dataset.dxr = "focusring";
      this.element.appendChild(this.focusRing);
    }
    const r = block.getBoundingClientRect();
    const scale = block.offsetWidth > 0 ? r.width / block.offsetWidth : 1;
    // (the CSS outline's 2px offset and 2px width)
    const left = Math.min(r.left, r.left + hang * scale) - 4, top = r.top - 4;
    const right = r.right + 4, bottom = r.bottom + 4;
    const area = this.element.querySelector<HTMLElement>("[data-dxr-scroll]")?.getBoundingClientRect();
    const clip = area
      ? `inset(${Math.max(0, area.top - top)}px ${Math.max(0, right - area.right)}px ${Math.max(0, bottom - area.bottom)}px ${Math.max(0, area.left - left)}px)`
      : "none";
    Object.assign(this.focusRing.style, {
      left: `${left}px`, top: `${top}px`, width: `${right - left}px`, height: `${bottom - top}px`, clipPath: clip,
    });
    this.focusRing.hidden = false;
  }

  openTabsDialog(): boolean {
    if (!this.live || this.tabsDialog.open || this.paragraphDialog.open) return false;
    this.closePopovers();
    return this.tabsDialog.show();
  }

  openLineNumbersDialog(): boolean {
    if (!this.live || this.lineNumbersDialog.open || this.columnsDialog.open || this.listDialog.open
      || this.paragraphDialog.open) return false;
    this.closePopovers();
    return this.lineNumbersDialog.show();
  }

  openColumnsDialog(): boolean {
    if (!this.live || this.columnsDialog.open || this.listDialog.open || this.paragraphDialog.open) return false;
    this.closePopovers();
    const sel = window.getSelection();
    const range = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const surface = this.control("editor");
    this.columnsRange = range && surface?.contains(range.commonAncestorContainer) ? range.cloneRange() : null;
    return this.columnsDialog.show();
  }

  openListDialog(): boolean {
    if (!this.live || this.listDialog.open || this.paragraphDialog.open) return false;
    this.closePopovers();
    this.listDialog.show();
    return true;
  }

  setListDefaults(levels: ListLevelSpec[]): void {
    this.listDefaults = levels.map((l) => ({ ...l }));
    this.live?.setListDefaults(this.listDefaults);
  }

  openParagraphDialog(): boolean {
    if (!this.live || this.paragraphDialog.open) return false;
    this.closePopovers();
    return this.paragraphDialog.show();
  }

  setMeasurementUnit(unit: MeasurementUnit): void {
    this.paragraphDialog.setUnit(unit);
    this.listDialog.setUnit(unit);
    this.columnsDialog.setUnit(unit);
    this.pictureDialog.setUnit(unit);
    this.tableDialog.setUnit(unit);
    this.lineNumbersDialog.setUnit(unit);
    this.tabsDialog.setUnit(unit);
    this.rulerUnit = unit;
    this.scheduleRuler();
    this.pictureUnit = unit;
    this.pictureChanged(this.live?.selectedPicture() ?? null);
  }

  private wire(): void {
    this.wireParagraphDialog();
    this.wireListDialog();
    this.wireColumns();
    this.wirePicture();
    this.wireTable();
    this.wireLineNumbers();
    this.wireBookmarks();
    this.wireNavPane();
    this.wireRuler();
    const ribbon = this.require("ribbon");

    // (the File tab opens a menu instead: wireFileMenu)
    for (const tab of Array.from(this.element.querySelectorAll<HTMLElement>(".dxr-tab[data-tab]"))) {
      this.keepSelection(tab);
      tab.addEventListener("click", () => this.selectTab(tab.dataset.tab ?? "home"));
    }

    const delegate = <T extends HTMLElement>(selector: string, label: (el: T) => string, fn: (el: T) => void,
      command?: (el: T) => CommandOf | undefined) => {
      for (const el of Array.from(ribbon.querySelectorAll<T>(selector))) {
        this.keepSelection(el);
        el.addEventListener("click", () => this.run(label(el), () => fn(el), command?.(el)));
      }
    };

    delegate<HTMLElement>("button[data-cmd]", (b) => b.dataset.cmd!, (b) =>
      this.live!.format(b.dataset.cmd as FormatKey), (b) => cmd(b.dataset.cmd!));
    delegate<HTMLElement>("button[data-align]", (b) => `align ${b.dataset.align}`, (b) =>
      this.live!.setAlignment(b.dataset.align as EditorAlignment), (b) => cmd("align", { alignment: b.dataset.align! }));
    delegate<HTMLElement>("button[data-indent]", () => "indent", (b) =>
      this.live!.indent(parseInt(b.dataset.indent ?? "720", 10)), (b) => cmd("indent", { twips: parseInt(b.dataset.indent ?? "720", 10) }));
    delegate<HTMLElement>("button[data-list]", (b) => `list ${b.dataset.list}`, (b) =>
      this.live!.toggleList(b.dataset.list as "bullet" | "decimal"), (b) => cmd("list", { kind: b.dataset.list! }));
    delegate<HTMLElement>("button[data-pagebreak]", () => "page break", () =>
      this.insertBreakAt("page"), () => () => this.breakCommand("page"));
    const tableCommands: Record<string, RibbonCommand> = {
      rowAbove: cmd("insertRow", { where: "above" }),
      rowBelow: cmd("insertRow", { where: "below" }),
      colLeft: cmd("insertColumn", { where: "left" }),
      colRight: cmd("insertColumn", { where: "right" }),
      delRow: cmd("deleteRow"),
      delCol: cmd("deleteColumn"),
      mergeRight: cmd("mergeCells", { rows: 1, columns: 2 }),
      mergeDown: cmd("mergeCells", { rows: 2, columns: 1 }),
      unmerge: cmd("unmergeCells"),
      borders: cmd("tableBorders", { spec: { scope: "all", style: "single", size: 4 } }),
      noborders: cmd("tableBorders", { spec: { scope: "all", style: "none" } }),
      noshade: cmd("cellShading", { fill: "", scope: "cell" }),
      alignTop: cmd("cellAlignment", { align: "top" }),
      alignCenter: cmd("cellAlignment", { align: "center" }),
      alignBottom: cmd("cellAlignment", { align: "bottom" }),
      delTable: cmd("deleteTable"),
    };
    delegate<HTMLElement>('.dxr-panel[data-panel="table"] button[data-tt]', (b) => b.dataset.tt!, (b) => {
      const ops: Record<string, () => void> = {
        rowAbove: () => this.live!.insertTableRow("above"),
        rowBelow: () => this.live!.insertTableRow("below"),
        colLeft: () => this.live!.insertTableColumn("left"),
        colRight: () => this.live!.insertTableColumn("right"),
        delRow: () => this.live!.deleteTableRow(),
        delCol: () => this.live!.deleteTableColumn(),
        mergeRight: () => this.live!.mergeCells(1, 2),
        mergeDown: () => this.live!.mergeCells(2, 1),
        unmerge: () => this.live!.unmergeCells(),
        borders: () => this.live!.setTableBorders({ scope: "all", style: "single", size: 4 }),
        noborders: () => this.live!.setTableBorders({ scope: "all", style: "none" }),
        noshade: () => this.live!.setCellShading("", "cell"),
        alignTop: () => this.live!.setCellProperties({ verticalAlignment: "top" }),
        alignCenter: () => this.live!.setCellProperties({ verticalAlignment: "center" }),
        alignBottom: () => this.live!.setCellProperties({ verticalAlignment: "bottom" }),
        delTable: () => this.live!.deleteTable(),
      };
      ops[b.dataset.tt ?? ""]?.();
    }, (b) => tableCommands[b.dataset.tt ?? ""]);
    const revisionCommands: Record<string, RibbonCommand> = { acceptall: cmd("acceptAllChanges"), rejectall: cmd("rejectAllChanges") };
    delegate<HTMLElement>("button[data-rev]", (b) => `revision ${b.dataset.rev}`, (b) =>
      this.runRevisionCommand(b.dataset.rev ?? ""), (b) => revisionCommands[b.dataset.rev ?? ""]);
    const headerFooterCommands: Record<string, RibbonCommand> = {
      pagenum: cmd("pageNumber", { field: "currentPage" }),
      totalpages: cmd("pageNumber", { field: "totalPages" }),
      pageof: cmd("pageNumber", { field: "pageOfTotal" }),
      header: cmd("goToHeaderFooter", { which: "header" }),
      footer: cmd("goToHeaderFooter", { which: "footer" }),
      close: cmd("closeHeaderFooter"),
    };
    delegate<HTMLElement>("button[data-hf]", (b) => `header/footer ${b.dataset.hf}`, (b) =>
      this.runHeaderFooterCommand(b.dataset.hf ?? ""), (b) => headerFooterCommands[b.dataset.hf ?? ""]);

    const rulepos = () => (this.require<HTMLSelectElement>("rulepos").value === "above" ? "above" : "below");
    const rule = (weight: number, style: string) => () => cmd("horizontalRule", { weight, style, position: rulepos() });
    const simple: Array<[string, string, () => void, CommandOf]> = [
      ["undo", "undo", () => this.live!.undo(), cmd("undo")],
      ["redo", "redo", () => this.live!.redo(), cmd("redo")],
      ["hr", "rule", () => this.live!.insertHorizontalRule(12, "single", rulepos()), rule(12, "single")],
      ["hrThick", "thick rule", () => this.live!.insertHorizontalRule(24, "single", rulepos()), rule(24, "single")],
      ["hrDouble", "double rule", () => this.live!.insertHorizontalRule(12, "double", rulepos()), rule(12, "double")],
      ["hrClear", "clear border", () => this.live!.clearParagraphBorders(), cmd("clearBorders")],
      ["delblock", "delete block", () => this.live!.deleteBlock(), cmd("deleteParagraph")],
      ["footnote", "footnote", () => this.live!.insertFootnote(), cmd("footnote")],
      ["endnote", "endnote", () => this.live!.insertEndnote(), cmd("endnote")],
      ["pagenum", "page number", () => this.live!.insertPageNumber("currentPage"), cmd("pageNumber", { field: "currentPage" })],
      ["totalpages", "total pages", () => this.live!.insertPageNumber("totalPages"), cmd("pageNumber", { field: "totalPages" })],
      ["fontgrow", "grow font", () => this.live!.adjustFontSize(1), cmd("growFont")],
      ["fontshrink", "shrink font", () => this.live!.adjustFontSize(-1), cmd("shrinkFont")],
      ["clearformat", "clear formatting", () => this.live!.clearFormatting(), cmd("clearFormatting")],
      ["toc", "table of contents", () => this.live!.insertTableOfContents(), cmd("tableOfContents")],
      ["unlink", "remove link", () => this.live!.removeHyperlink(), cmd("removeHyperlink")],
      ["gotoheader", "go to header", () => this.live!.goToHeaderFooter("header"), cmd("goToHeaderFooter", { which: "header" })],
      ["gotofooter", "go to footer", () => this.live!.goToHeaderFooter("footer"), cmd("goToHeaderFooter", { which: "footer" })],
    ];
    for (const [name, label, fn, command] of simple) {
      const el = this.control(name);
      if (!el) continue;
      this.keepSelection(el);
      el.addEventListener("click", () => this.run(label, fn, command));
    }
    // (these run their command themselves, so it is reported once)
    const own: Array<[string, () => void]> = [
      ["viewpage", () => this.setPaginated(true)],
      ["viewweb", () => this.setPaginated(false)],
      ["zoomfit", () => this.fitPageWidth()],
      ["zoomonepage", () => this.fitOnePage()],
      ["zoomin", () => this.stepZoom(1)],
      ["zoomout", () => this.stepZoom(-1)],
    ];
    for (const [name, fn] of own) {
      const el = this.control(name);
      if (!el) continue;
      this.keepSelection(el);
      el.addEventListener("click", fn);
    }

    const smallcaps = this.control("smallcaps");
    if (smallcaps) {
      this.keepSelection(smallcaps);
      smallcaps.addEventListener("click", () => {
        const on = smallcaps.getAttribute("aria-pressed") !== "true";
        this.run("small caps", () => this.live!.setSmallCaps(on), cmd("smallCaps", { on }));
        smallcaps.setAttribute("aria-pressed", String(on));
      });
    }

    // Font size: `change` only (it fires on Enter and on blur). Binding keydown as
    // well fired twice and stacked two undo snapshots for one size change.
    const fontsize = this.require<HTMLInputElement>("fontsize");
    fontsize.addEventListener("change", () => {
      const pts = parseFloat(fontsize.value);
      if (this.live && pts > 0) this.run("font size", () => this.live!.setFontSize(pts), cmd("fontSize", { pts }));
    });
    fontsize.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        fontsize.blur();
      }
    });

    this.wireSelect("fontfamily", "font family", (v) => this.live!.setFontFamily(v), (v) => cmd("fontFamily", { name: v }));
    const styleBox = this.require<HTMLSelectElement>("style");
    styleBox.addEventListener("change", () => {
      const value = styleBox.value;
      if (this.live && value) this.run("style", () => this.live!.setParagraphStyle(value), cmd("style", { styleId: value }));
      // (then back to what the caret's paragraph has, the new style when it took)
      styleBox.blur();
      this.syncStyleBox();
    });
    const listMenu = this.require<HTMLSelectElement>("listmenu");
    listMenu.addEventListener("change", () => {
      const v = listMenu.value;
      listMenu.value = "";
      if (!this.live || !v) return;
      if (v === "levels") { this.openListDialog(); return; }
      if (v === "legal") this.run("legal numbering", () => this.live!.setLegalNumbering("top"), cmd("legalNumbering", { periods: "top" }));
      else if (v === "restart") this.run("restart numbering", () => this.live!.restartNumbering(1), cmd("restartNumbering", { start: 1 }));
      else if (v === "continue") this.run("continue numbering", () => this.live!.continueNumbering(), cmd("continueNumbering"));
      else this.run("list style", () => this.live!.setListFormat(v as ListFormat), cmd("listFormat", { format: v }));
    });
    this.wireSelect("linespacing", "line spacing", (v) => this.live!.setLineSpacing(parseFloat(v)),
      (v) => cmd("lineSpacing", { multiple: parseFloat(v) }));
    this.wireSelect("highlight", "highlight", (v) => this.live!.setHighlight(v === "none" ? "" : v), (v) => cmd("highlight", { color: v }));
    const pageNumberField = (v: string) => (v === "pageOf" ? "pageOfTotal" : (v as "currentPage" | "totalPages"));
    this.wireSelect("pagenummenu", "page number", (v) => this.live!.insertPageNumber(pageNumberField(v)),
      (v) => cmd("pageNumber", { field: pageNumberField(v) }));

    // Colour pickers: the native input steals focus, so the engine's last-selection
    // bookmark is what the command applies to.
    const fontcolor = this.require<HTMLInputElement>("fontcolor");
    fontcolor.addEventListener("input", () => {
      this.require("fontcolorbar").style.setProperty("--dxr-swatch", fontcolor.value);
    });
    const fontColor = () => this.run("font color", () => this.live!.setFontColor(fontcolor.value), cmd("fontColor", { color: fontcolor.value }));
    fontcolor.addEventListener("change", () => {
      this.require("fontcolorbar").style.setProperty("--dxr-swatch", fontcolor.value);
      fontColor();
    });
    const fontcolorbtn = this.control("fontcolorbtn");
    if (fontcolorbtn) {
      this.keepSelection(fontcolorbtn);
      fontcolorbtn.addEventListener("click", fontColor);
    }
    const shade = this.require<HTMLInputElement>("shade");
    const cellShading = () => this.run("cell shading", () => this.live!.setCellShading(shade.value, "cell"),
      cmd("cellShading", { fill: shade.value, scope: "cell" }));
    shade.addEventListener("input", () => this.require("shadebar").style.setProperty("--dxr-swatch", shade.value));
    shade.addEventListener("change", () => {
      this.require("shadebar").style.setProperty("--dxr-swatch", shade.value);
      cellShading();
    });
    const shadebtn = this.control("shadebtn");
    if (shadebtn) {
      this.keepSelection(shadebtn);
      shadebtn.addEventListener("click", cellShading);
    }
    const repeatheader = this.require<HTMLInputElement>("repeatheader");
    repeatheader.addEventListener("change", () =>
      this.run("repeat header row", () => this.live!.setRepeatHeaderRow(repeatheader.checked), cmd("repeatHeaderRow", { on: repeatheader.checked })));

    // Comments: New/Insert open a draft bubble in the gutter (Word's flow); the rest act on the
    // active thread.
    for (const name of ["comment", "insertcomment"]) {
      const button = this.control(name);
      if (!button) continue;
      this.keepSelection(button);
      button.addEventListener("click", () => this.beginComment());
    }
    this.control("commentdelete")?.addEventListener("click", () => {
      const active = this.live?.activeComment;
      if (!active) return;
      this.run("delete comment", () => {
        for (const id of threadMembers(this.live!.listComments(), active)) this.live!.removeComment(id);
        this.live!.removeComment(active);
      });
    });
    this.control("commentresolve")?.addEventListener("click", () => {
      const active = this.live?.activeComment;
      if (!active) return;
      const entry = this.live!.listComments().find((c) => c.anchorId === active);
      if (!entry) return;
      this.run(entry.resolved ? "reopen comment" : "resolve comment", () =>
        this.live!.setCommentResolved(active, !entry.resolved));
      this.live!.activateComment(active);
    });
    this.control("commentprev")?.addEventListener("click", () => this.live?.stepComment(-1));
    this.control("commentnext")?.addEventListener("click", () => this.live?.stepComment(1));
    const showcomments = this.require<HTMLInputElement>("showcomments");
    showcomments.addEventListener("change", () => {
      this.live?.showComments(showcomments.checked);
      this.surface.dataset.comments = showcomments.checked && this.options.comments !== false ? "on" : "off";
    });

    // Tracking.
    const trackchanges = this.require<HTMLInputElement>("trackchanges");
    trackchanges.addEventListener("change", () => {
      if (!this.live) return;
      this.run(trackchanges.checked ? "track changes on" : "track changes off", () =>
        this.live!.setTrackedChanges(trackchanges.checked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept),
        cmd("trackChanges", { on: trackchanges.checked }));
      this.syncReviewState();
    });
    const markup = this.require<HTMLSelectElement>("markup");
    markup.addEventListener("change", () => {
      if (!this.live) return;
      // "No markup" shows the document as if every change were accepted; edits keep recording.
      const visible = markup.value !== "none";
      this.showRevisions = visible;
      this.run(`markup ${markup.value}`, () => this.live!.showRevisions(visible), cmd("showMarkup", { on: visible }));
    });
    const author = this.require<HTMLInputElement>("author");
    author.addEventListener("change", () => {
      const name = author.value.trim() || "Reviewer";
      // Initials given for one name do not belong to another: a new name takes its own.
      if (name !== this.author) this.initials = "";
      this.author = name;
      this.live?.setRevisionAuthor(this.author);
      this.live?.setCommentAuthor(this.author, this.initials);
    });

    this.wireFileActions();
    this.wireLayout();
    this.wirePicker();
    this.wireFileMenu();
    this.wirePaste();
    this.wireFormatPainter();
    this.wireChangeCase();
    this.wireShowMarks();
    this.wireLinkPopover();
    this.wireFindBar();
    this.wireHeaderFooterTab();
    this.wireZoom();
    this.wireViewToggles();
  }

  /** New: the host's (onNewRequest), the history's, or a blank document at once. */
  private requestNew(): void {
    if (this.options.onNewRequest) { this.options.onNewRequest(); return; }
    if (this.history) { void this.history.newDocument(); return; }
    if (this.exports) this.openBlank("untitled.docx");
  }

  /** Open a document: the host's picker (onOpenRequest), otherwise the browser's. */
  private requestOpen(): void {
    if (this.options.onOpenRequest) { this.options.onOpenRequest(); return; }
    this.control<HTMLInputElement>("file")?.click();
  }

  /** A `<select>` that acts on change and snaps back to its placeholder. */
  private wireSelect(name: string, label: string, apply: (value: string) => void, command?: (value: string) => CommandOf): void {
    const select = this.require<HTMLSelectElement>(name);
    select.addEventListener("change", () => {
      const value = select.value;
      if (this.live && value) this.run(label, () => apply(value), command?.(value));
      select.value = "";
    });
  }

  private wireFileActions(): void {
    const file = this.control<HTMLInputElement>("file");
    file?.addEventListener("change", async () => {
      const chosen = file.files?.[0];
      if (!chosen) return;
      file.value = '';
      if (this.history) { await this.history.openFile(chosen); return; }
      this.setStatus(`Loading ${chosen.name}…`);
      try { this.open(new Uint8Array(await chosen.arrayBuffer()), chosen.name); }
      catch (error) { this.setStatus(`Could not open this document. ${String(error)}`); }
      // Reset so re-picking the same file fires `change` again.
      file.value = "";
    });
    // Open is a label around the hidden file input: a click on it opens the browser's picker,
    // unless the host has its own (onOpenRequest). The label is focusable, so Enter and Space
    // work it too.
    const openLabel = file?.closest("label");
    openLabel?.addEventListener("click", (event) => {
      if (!this.options.onOpenRequest) return;
      event.preventDefault();
      this.options.onOpenRequest();
    });
    openLabel?.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      this.requestOpen();
    });
    this.control("new")?.addEventListener("click", () => this.requestNew());
    this.control("save")?.addEventListener("click", () => this.download());
    const recent = this.control<HTMLSelectElement>("recent");
    recent?.addEventListener("change", () => {
      const entry = recent.value === "" ? undefined : this.recentEntries[Number(recent.value)];
      recent.value = ""; // back to the placeholder, so the same entry can be picked again
      if (entry) this.options.onOpenRecent?.(entry);
    });

    const picture = this.control<HTMLInputElement>("picturefile");
    picture?.addEventListener("change", async () => {
      const chosen = picture.files?.[0];
      picture.value = "";
      if (!chosen || !this.live) return;
      const started = performance.now();
      const ok = await this.live.insertImageFile(chosen, { altText: chosen.name });
      if (ok) { this.history?.edited(); this.scheduleDirtyCheck(); }
      const ms = performance.now() - started;
      const el = this.control("railOp");
      if (el) el.textContent = `picture ${Math.round(ms)} ms`;
      this.setStatus(ok ? `Inserted ${chosen.name}` : "Could not insert the picture here");
      this.refreshRailCounts();
    });
  }

  private wireLayout(): void {
    // Pagination toggles re-render the LIVE session, so edits survive the switch.
    const paginated = this.require<HTMLInputElement>("paginated");
    paginated.addEventListener("change", () => this.setPaginated(paginated.checked));

    // The region is chosen at open time, so toggling it re-opens the LIVE
    // session's bytes rather than the originally loaded file — edits so far survive.
    const headerFooter = this.require<HTMLInputElement>("headerfooter");
    headerFooter.addEventListener("change", () => {
      this.headerFooter = headerFooter.checked;
      if (!this.live) return;
      this.open(this.live.save(), this.documentName);
    });

    const pgfmt = this.require<HTMLSelectElement>("pgfmt");
    pgfmt.addEventListener("change", () => {
      if (!this.live || !pgfmt.value) return;
      this.run("page format", () => this.live!.setPageNumbering({ format: pgfmt.value as NumberFormat }),
        cmd("pageNumbering", { format: pgfmt.value }));
      this.syncPageNumbering();
    });
    const pgstart = this.require<HTMLInputElement>("pgstart");
    pgstart.addEventListener("change", () => {
      const value = parseInt(pgstart.value, 10);
      if (!this.live || !(value > 0)) return;
      this.run("page start", () => this.live!.setPageNumbering({ start: value }), cmd("pageNumbering", { start: value }));
      this.syncPageNumbering();
    });
    this.control("pgclear")?.addEventListener("click", () => {
      this.run("clear numbering", () => this.live!.clearPageNumbering(), cmd("clearPageNumbering"));
      this.syncPageNumbering();
    });

    // Page setup (Word's Layout > Page Setup): margins presets, orientation, size.
    const pageSetup = (op: Parameters<DocxEditor["setPageSetup"]>[0]) => {
      if (!this.live!.setPageSetup(op)) this.setStatus("This engine build cannot change the page setup");
      this.syncLayoutState();
    };
    const margins = (v: string) => {
      const [top, bottom, left, right] = v.split(",").map((n) => parseInt(n, 10));
      return { marginTopTwips: top, marginBottomTwips: bottom, marginLeftTwips: left, marginRightTwips: right };
    };
    this.wireSelect("breaks", "break", (v) => this.insertBreakAt(v), (v) => () => this.breakCommand(v));
    this.wireSelect("sectionstart", "section start", (v) => pageSetup({ sectionStart: v as SectionStart }),
      (v) => cmd("pageSetup", { sectionStart: v }));
    this.wireSelect("margins", "margins", (v) => pageSetup(margins(v)), (v) => cmd("pageSetup", margins(v)));
    this.wireSelect("orientation", "orientation", (v) => pageSetup({ landscape: v === "landscape" }),
      (v) => cmd("pageSetup", { landscape: v === "landscape" }));
    // (a size keeps the orientation the section had before it: worked out when the choice is made)
    let size: EditorPageSetup = {};
    this.wireSelect("pagesize", "page size", (v) => {
      const [w, h] = v.split(",").map((n) => parseInt(n, 10));
      const landscape = !!this.live!.sectionInfo()?.landscape;
      size = landscape
        ? { pageWidthTwips: h, pageHeightTwips: w, landscape: true }
        : { pageWidthTwips: w, pageHeightTwips: h };
      pageSetup(size);
    }, () => () => cmd("pageSetup", { ...size }));

    // Paragraph spacing / special indent.
    const before = this.require<HTMLInputElement>("spacebefore");
    const after = this.require<HTMLInputElement>("spaceafter");
    const spacing = (which: "before" | "after", input: HTMLInputElement) => {
      const pts = parseFloat(input.value);
      if (!this.live || !(pts >= 0)) return;
      this.run(`space ${which}`, () =>
        this.live!.setParagraphSpacing(which === "before" ? { beforePt: pts } : { afterPt: pts }),
        cmd("paragraphSpacing", { [which]: pts }));
    };
    before.addEventListener("change", () => spacing("before", before));
    after.addEventListener("change", () => spacing("after", after));
    const special = this.require<HTMLSelectElement>("specialindent");
    const specialVal = this.require<HTMLInputElement>("specialindentval");
    const applySpecial = () => {
      if (!this.live || !special.value) return;
      const inches = parseFloat(specialVal.value);
      const twips = Math.round((inches >= 0 ? inches : 0.5) * 1440);
      const kind = special.value;
      this.run("special indent", () => {
        if (kind === "firstLine") this.live!.setFirstLineIndent(twips);
        else if (kind === "hanging") this.live!.setHangingIndent(twips);
        else this.live!.setFirstLineIndent(0);
      }, kind === "hanging" ? cmd("hangingIndent", { twips }) : cmd("firstLineIndent", { twips: kind === "firstLine" ? twips : 0 }));
    };
    special.addEventListener("change", applySpecial);
    specialVal.addEventListener("change", applySpecial);
  }

  private setPaginated(on: boolean): void {
    const paginated = this.require<HTMLInputElement>("paginated");
    paginated.checked = on;
    this.surface.dataset.view = on ? "paginated" : "continuous";
    if (!this.live) return;
    this.run(on ? "page view" : "continuous view", () => this.live!.setPaginated(on), cmd("pageView", { on }));
    this.syncViewButtons();
  }

  private syncViewButtons(): void {
    const on = this.require<HTMLInputElement>("paginated").checked;
    this.control("viewpage")?.classList.toggle("dxr-on", on);
    this.control("viewweb")?.classList.toggle("dxr-on", !on);
  }

  private wireViewToggles(): void {
    const showrail = this.require<HTMLInputElement>("showrail");
    showrail.addEventListener("change", () => {
      const rail = this.element.querySelector<HTMLElement>("[data-dxr-rail]");
      if (rail) rail.hidden = !showrail.checked;
    });
    const showhint = this.require<HTMLInputElement>("showhint");
    showhint.addEventListener("change", () => {
      const hint = this.element.querySelector<HTMLElement>("[data-dxr-hint]");
      if (hint) hint.hidden = !showhint.checked;
    });
    this.syncViewButtons();
  }

  /** The bands own the same setting and read the live session, so both stay in step. */
  private syncPageNumbering(): void {
    if (!this.live) return;
    const numbering = this.live.pageNumbering() ?? {};
    this.require<HTMLSelectElement>("pgfmt").value = numbering.format ?? "";
    this.require<HTMLInputElement>("pgstart").value = numbering.start != null ? String(numbering.start) : "";
  }

  /** Layout tab: describe the section under the caret. */
  private syncLayoutState(): void {
    const note = this.control("pagesetupnote");
    const info = this.live?.sectionInfo() ?? null;
    if (!note) return;
    if (!info) {
      note.textContent = "";
      return;
    }
    const inches = (twips: number) => (twips / 1440).toFixed(2).replace(/\.?0+$/, "");
    note.textContent =
      `${inches(info.pageWidthTwips)}″ × ${inches(info.pageHeightTwips)}″ ${info.landscape ? "landscape" : "portrait"}, ` +
      `margins ${inches(info.marginTopTwips)}″ / ${inches(info.marginRightTwips)}″ / ${inches(info.marginBottomTwips)}″ / ${inches(info.marginLeftTwips)}″`;
  }

  // ── review ──────────────────────────────────────────────────────────────────

  private beginComment(): void {
    if (!this.live) return;
    if (this.density === "compact") {
      // No gutter on a phone: post directly on the selection with a prompt.
      const text = window.prompt("Comment");
      if (!text?.trim()) return;
      this.run("comment", () => this.live!.addComment(text.trim(), this.author), cmd("comment", { text: text.trim() }));
      return;
    }
    if (!this.live.beginComment()) this.setStatus("Click into a paragraph (or select text) to comment on it");
  }

  private onCommentsChange(info: { threads: number; open: number; active: string | null }): void {
    const count = this.control("commentcount");
    if (count) count.textContent = info.threads === 0 ? "No comments" : `${info.open} open · ${info.threads} total`;
    const resolve = this.control<HTMLButtonElement>("commentresolve");
    const del = this.control<HTMLButtonElement>("commentdelete");
    const entry = info.active ? this.live?.listComments().find((c) => c.anchorId === info.active) : null;
    if (resolve) {
      resolve.disabled = !entry;
      resolve.textContent = entry?.resolved ? "Reopen" : "Resolve";
    }
    if (del) del.disabled = !entry;
  }

  private syncReviewState(): void {
    if (!this.live) return;
    this.require<HTMLInputElement>("trackchanges").checked =
      this.live.trackedChanges === TrackedChangeMode.RenderInline;
    this.require<HTMLSelectElement>("markup").value = this.live.revisionsVisible ? "all" : "none";
    const revs = this.live.listRevisions().filter(r =>
      r.diagnostic?.code !== "inert_move_range" && r.diagnostic?.code !== "empty_wrapper_range");
    const count = this.control("revcount");
    if (count) count.textContent = revs.length === 0 ? "No tracked changes" : `${revs.length} change${revs.length === 1 ? "" : "s"}`;
    this.require<HTMLInputElement>("showcomments").checked = this.live.commentsVisible || this.options.comments === false;
  }

  private runRevisionCommand(op: string): void {
    if (!this.live) return;
    const marks = this.live.revisionElements();
    const current = this.currentRevisionMark(marks);
    switch (op) {
      case "accept":
      case "reject": {
        const mark = current ?? marks[0];
        const entry = mark ? this.live.revisionAt(mark) : null;
        if (!entry) { this.setStatus("Put the caret in a tracked change first"); return; }
        if (op === "accept") this.live.acceptRevision(entry.id);
        else this.live.rejectRevision(entry.id);
        break;
      }
      case "acceptall": this.live.acceptAllRevisions(); break;
      case "rejectall": this.live.rejectAllRevisions(); break;
      case "next":
      case "prev": {
        if (marks.length === 0) { this.setStatus("No tracked changes"); return; }
        const index = current ? marks.indexOf(current) : -1;
        const target = marks[(index + (op === "next" ? 1 : -1) + marks.length) % marks.length];
        this.selectRevisionMark(target);
        return;
      }
      default: return;
    }
    this.syncReviewState();
  }

  private currentRevisionMark(marks: HTMLElement[]): HTMLElement | null {
    const sel = (this.element.ownerDocument ?? document).getSelection();
    const node = sel?.anchorNode ?? null;
    if (!node) return null;
    const el = node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
    const mark = el?.closest<HTMLElement>('ins, del, tr[class*="row-"]') ?? null;
    return mark && marks.includes(mark) ? mark : null;
  }

  private selectRevisionMark(mark: HTMLElement): void {
    const doc = this.element.ownerDocument ?? document;
    const block = mark.closest<HTMLElement>('[data-anchor][contenteditable="true"]');
    block?.focus({ preventScroll: true });
    const range = doc.createRange();
    range.selectNodeContents(mark);
    const sel = doc.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    mark.scrollIntoView({ block: "center", behavior: "smooth" });
  }

  // ── header & footer tab ─────────────────────────────────────────────────────

  private wireHeaderFooterTab(): void {
    for (const kind of ["first", "even"] as const) {
      const box = this.require<HTMLInputElement>(kind === "first" ? "hfFirst" : "hfEven");
      box.addEventListener("change", () => {
        if (!this.live) return;
        const ok = this.live.setHeaderFooterKindEnabled(kind, box.checked);
        if (!ok) {
          box.checked = !box.checked;
          this.setStatus("This engine build cannot change that option");
        }
        this.run(`${kind === "first" ? "different first page" : "different odd & even"} ${box.checked ? "on" : "off"}`, () => {},
          ok ? cmd("headerFooterOption", { kind, on: box.checked }) : undefined);
        this.syncHeaderFooterState();
      });
    }
  }

  private runHeaderFooterCommand(op: string): void {
    if (!this.live) return;
    switch (op) {
      case "pagenum": this.live.insertPageNumber("currentPage"); break;
      case "totalpages": this.live.insertPageNumber("totalPages"); break;
      case "pageof": this.live.insertPageNumber("pageOfTotal"); break;
      case "header": this.live.goToHeaderFooter("header"); break;
      case "footer": this.live.goToHeaderFooter("footer"); break;
      case "close": this.live.closeHeaderFooter(); break;
    }
    this.syncHeaderFooterState();
  }

  private onStoryChange(which: BandWhich | null): void {
    const tab = this.element.querySelector<HTMLElement>('.dxr-tab[data-tab="headerfooter"]');
    if (!tab) return;
    tab.hidden = which === null;
    if (which) {
      this.selectTab("headerfooter");
    } else if (tab.getAttribute("aria-selected") === "true") {
      this.selectTab("home");
    }
    this.syncHeaderFooterState();
  }

  private syncHeaderFooterState(): void {
    if (!this.live) return;
    this.require<HTMLInputElement>("hfFirst").checked = this.live.headerFooterKindEnabled("first");
    this.require<HTMLInputElement>("hfEven").checked = this.live.headerFooterKindEnabled("even");
    const story = this.control("hfstory");
    if (story) {
      const which = this.live.activeStoryKind ?? "header";
      story.textContent = this.live.headerFooterStoryLabel(which);
    }
  }

  // ── zoom ────────────────────────────────────────────────────────────────────

  private wireZoom(): void {
    for (const name of ["zoom", "zoomlevel"]) {
      const select = this.control<HTMLSelectElement>(name);
      select?.addEventListener("change", () => this.setZoom(parseFloat(select.value)));
    }
  }

  private setZoom(scale: number): void {
    if (!this.live || !(scale > 0)) return;
    this.run(`zoom ${Math.round(scale * 100)}%`, () => this.live!.setZoom(scale), cmd("zoom", { scale }));
    this.syncZoom();
  }

  /** Page width: the page as wide as the window, as Word's Page Width (kept as the window resizes). */
  private fitPageWidth(): void {
    if (!this.live) return;
    this.run("page width", () => this.live!.fitPageWidth(), cmd("pageWidth"));
    this.syncZoom();
  }

  /** One page: a whole page in view, as Word's One Page (kept as the window resizes). */
  private fitOnePage(): void {
    if (!this.live) return;
    this.run("one page", () => this.live!.fitOnePage(), cmd("onePage"));
    this.syncZoom();
  }

  private stepZoom(direction: 1 | -1): void {
    if (!this.live) return;
    const current = this.live.requestedZoom;
    const next = direction > 0
      ? ZOOM_STEPS.find((z) => z > current + 1e-6) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1]
      : [...ZOOM_STEPS].reverse().find((z) => z < current - 1e-6) ?? ZOOM_STEPS[0];
    this.setZoom(next);
  }

  private syncZoom(): void {
    if (!this.live) return;
    // (page width's zoom is any number: shown, and offered, as a whole percent)
    const requested = Math.round(this.live.requestedZoom * 100) / 100;
    const value = String(ZOOM_STEPS.find((z) => Math.abs(z - requested) < 1e-6) ?? requested);
    for (const name of ["zoom", "zoomlevel"]) {
      const select = this.control<HTMLSelectElement>(name);
      if (!select) continue;
      if (!Array.from(select.options).some((o) => o.value === value)) {
        const opt = document.createElement("option");
        opt.value = value;
        opt.textContent = `${Math.round(requested * 100)}%`;
        select.appendChild(opt);
      }
      select.value = value;
    }
  }

  // ── find & replace ──────────────────────────────────────────────────────────

  private wireFindBar(): void {
    const bar = this.require("findbar");
    const text = this.require<HTMLInputElement>("findtext");
    const replace = this.require<HTMLInputElement>("replacetext");
    this.control("findtoggle")?.addEventListener("click", () => this.openFindBar(false));
    this.control("replacetoggle")?.addEventListener("click", () => this.openFindBar(true));
    this.control("findclose")?.addEventListener("click", () => this.closeFindBar());
    text.addEventListener("input", () => this.refreshFind(true));
    // Special inserts its code into the field last typed in.
    text.addEventListener("focus", () => { this.findSpecialTarget = "find"; });
    replace.addEventListener("focus", () => { this.findSpecialTarget = "replace"; });
    this.wireFindOptions();
    text.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.stepFind(event.shiftKey ? -1 : 1);
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.closeFindBar();
      }
    });
    replace.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.replaceCurrent();
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.closeFindBar();
      }
    });
    this.control("findprev")?.addEventListener("click", () => this.stepFind(-1));
    this.control("findnext")?.addEventListener("click", () => this.stepFind(1));
    this.control("replaceone")?.addEventListener("click", () => this.replaceCurrent());
    this.control("replaceall")?.addEventListener("click", () => {
      if (!this.live || !text.value || this.live.findError(text.value, this.findOptions)) return;
      let count = 0;
      const options = normalizeFindOptions(this.findOptions);
      this.run("replace all", () => {
        count = this.live!.replaceAll(text.value, replace.value, { ...options, focus: false });
      }, cmd("replaceAll", { text: text.value, with: replace.value, ...options }));
      this.setStatus(`Replaced ${count} occurrence${count === 1 ? "" : "s"}`);
      this.refreshFind(true);
    });
    // Every find-bar button acts on the search, not on the caret: swallowing mousedown keeps the
    // keyboard in the field that was being typed into, so a click on Next does not end the query.
    for (const id of ["findprev", "findnext", "replaceone", "replaceall"]) {
      const button = this.control(id);
      if (button) this.keepSelection(button);
    }
    bar.hidden = true;
  }

  private openFindBar(withReplace: boolean): void {
    const bar = this.require("findbar");
    bar.hidden = false;
    this.require("replacegroup").hidden = !withReplace;
    this.syncStripFade(bar);
    const text = this.require<HTMLInputElement>("findtext");
    text.focus();
    text.select();
    // Painting is focus-safe, so a reopened bar can show its hits without taking the keyboard.
    this.refreshFind(true);
  }

  /**
   * Close the bar and hand the current hit over to the caret: the document is where the user is
   * going next, and Word leaves the insertion point on the match it stopped at. Until this runs
   * the match is only PAINTED (see `showFindMatches`), so the search field keeps the keyboard for
   * as long as the bar is open.
   */
  private closeFindBar(): void {
    const bar = this.control("findbar");
    if (bar) bar.hidden = true;
    const match = this.findMatches[this.findIndex];
    if (match) this.live?.selectMatch(match);
    else this.live?.clearFindMatches();
    this.findMatches = [];
    this.findIndex = -1;
  }

  /**
   * Re-scan for the current query. `show` paints the hits and rides to the first one — it never
   * focuses the document, so typing the second character of a query that already matched goes on
   * refining the search instead of being typed into the document.
   */
  private refreshFind(show: boolean): void {
    const text = this.require<HTMLInputElement>("findtext");
    const error = this.live && text.value ? this.live.findError(text.value, this.findOptions) : null;
    this.findMatches = this.live && text.value && !error ? this.live.find(text.value, this.findOptions) : [];
    this.findIndex = this.findMatches.length > 0 ? 0 : -1;
    this.findErrorText = error;
    this.updateFindCount();
    if (!show) return;
    if (this.findIndex >= 0) this.live?.showFindMatches(this.findMatches, this.findIndex);
    else this.live?.clearFindMatches();
  }

  private stepFind(direction: 1 | -1): void {
    if (!this.live) return;
    if (this.findMatches.length === 0) this.refreshFind(false);
    if (this.findMatches.length === 0) return;
    this.findIndex = (this.findIndex + direction + this.findMatches.length) % this.findMatches.length;
    this.live.showFindMatches(this.findMatches, this.findIndex);
    this.updateFindCount();
  }

  private replaceCurrent(): void {
    if (!this.live) return;
    if (this.findMatches.length === 0 || this.findIndex < 0) this.refreshFind(true);
    const match = this.findMatches[this.findIndex];
    if (!match) return;
    const replacement = this.require<HTMLInputElement>("replacetext").value;
    this.run("replace", () => this.live!.replaceMatch(match, replacement, { focus: false }));
    // Offsets after the replacement shifted; re-scan and continue from the same slot.
    const keep = this.findIndex;
    this.refreshFind(false);
    if (this.findMatches.length > 0) {
      this.findIndex = Math.min(keep, this.findMatches.length - 1);
      this.live.showFindMatches(this.findMatches, this.findIndex);
      this.updateFindCount();
    }
  }

  private updateFindCount(): void {
    const el = this.control("findcount");
    if (!el) return;
    const field = this.require<HTMLInputElement>("findtext");
    const error = this.findErrorText;
    el.textContent = error
      ? this.translateText("Invalid")
      : this.findMatches.length === 0
        ? (field.value ? "0 of 0" : "")
        : `${this.findIndex + 1} of ${this.findMatches.length}`;
    el.title = error ?? "";
    if (error) field.setAttribute("aria-invalid", "true");
    else field.removeAttribute("aria-invalid");
    field.title = error ?? "";
  }

  /**
   * A page, column or section break at the caret (Layout > Breaks, the Page break button,
   * Ctrl+Enter). The engine refuses one in a table, header or footer, or while changes are
   * tracked: the status line says so, and {@link breakCommand} keeps it out of a recording.
   */
  private insertBreakAt(kind: string): void {
    const live = this.live!;
    if (kind === "removeSection") {
      this.breakRefused = !live.removeSectionBreak();
      if (this.breakRefused) this.setStatus(this.translateText(RIBBON_STATUS_TEXTS.noSectionBreak));
      return;
    }
    this.breakRefused = !(kind === "page" || kind === "column"
      ? live.insertBreak(kind)
      : live.insertSectionBreak(kind as SectionStart));
    if (this.breakRefused) this.setStatus(this.translateText(RIBBON_STATUS_TEXTS.breakRefused));
  }

  /** The recorded command for the break {@link insertBreakAt} just made (none when it was refused). */
  private breakCommand(kind: string): RibbonCommand | undefined {
    if (this.breakRefused) return undefined;
    if (kind === "removeSection") return cmd("removeSectionBreak");
    return kind === "page" || kind === "column" ? cmd("insertBreak", { kind }) : cmd("sectionBreak", { start: kind });
  }

  /** A text in the ribbon's language (see `setLabels`), for text the ribbon writes later. */
  private translateText(text: string): string {
    return this.labels[text] ?? text;
  }

  /** The find bar's Options ▾ (Word's More >> checkboxes) and Special ▾ menus. */
  private wireFindOptions(): void {
    this.findOptions = normalizeFindOptions(this.options.findOptions);
    const button = this.control("findoptions"), menu = this.control("findoptmenu");
    const special = this.control("findspecial"), specialMenu = this.control("findspecialmenu");
    if (button && menu) {
      this.keepSelection(button);
      button.addEventListener("click", () => {
        if (menu.hasAttribute("data-open")) { this.closePopovers(); return; }
        this.closePopovers();
        menu.setAttribute("data-open", "");
        button.setAttribute("aria-expanded", "true");
        this.positionPopover(menu, button);
      });
      menu.addEventListener("change", (event) => {
        const box = (event.target as HTMLElement).closest<HTMLInputElement>("input[data-findopt]");
        if (!box) return;
        const key = box.dataset.findopt as keyof FindOptions;
        const next: FindOptions = { ...this.findOptions, [key]: box.checked };
        // Wildcards and regular expressions are two readings of the query: one at a time.
        if (box.checked && key === "wildcards") next.regex = false;
        if (box.checked && key === "regex") next.wildcards = false;
        this.applyFindOptions(next, true);
        this.require<HTMLInputElement>(this.findSpecialTarget === "replace" ? "replacetext" : "findtext").focus();
      });
      menu.addEventListener("keydown", (event) => {
        if (event.key === "Escape") { event.preventDefault(); this.closePopovers(); button.focus(); }
      });
    }
    if (special && specialMenu) {
      this.keepSelection(special);
      special.addEventListener("click", () => {
        if (specialMenu.hasAttribute("data-open")) { this.closePopovers(); return; }
        this.closePopovers();
        this.renderSpecialMenu(specialMenu);
        specialMenu.setAttribute("data-open", "");
        special.setAttribute("aria-expanded", "true");
        this.positionPopover(specialMenu, special);
      });
      specialMenu.addEventListener("mousedown", (event) => event.preventDefault());
      specialMenu.addEventListener("click", (event) => {
        const chosen = (event.target as HTMLElement).closest<HTMLButtonElement>("button[data-code]");
        if (!chosen) return;
        this.closePopovers();
        this.insertFindCode(chosen.dataset.code ?? "");
      });
      specialMenu.addEventListener("keydown", (event) => {
        const items = Array.from(specialMenu.querySelectorAll<HTMLButtonElement>("button"));
        const at = items.indexOf(event.target as HTMLButtonElement);
        let next = -1;
        if (event.key === "ArrowDown") next = (at + 1) % items.length;
        else if (event.key === "ArrowUp") next = (at - 1 + items.length) % items.length;
        else if (event.key === "Escape") { event.preventDefault(); this.closePopovers(); special.focus(); return; }
        if (next >= 0) { event.preventDefault(); items[next].focus(); }
      });
    }
    this.syncFindOptionControls();
  }

  /** The codes Special offers for the field it will insert into, in the current mode. */
  private renderSpecialMenu(menu: HTMLElement): void {
    const codes = this.findSpecialTarget === "replace"
      ? specialReplaceCodes(this.findOptions)
      : specialFindCodes(this.findOptions);
    const doc = menu.ownerDocument;
    menu.replaceChildren(...codes.map((c) => {
      const item = doc.createElement("button");
      item.type = "button";
      item.setAttribute("role", "menuitem");
      item.dataset.code = c.code;
      const label = doc.createElement("span");
      label.textContent = this.translateText(c.label);
      const code = doc.createElement("kbd");
      code.textContent = c.code;
      item.append(label, code);
      return item;
    }));
  }

  /** Put a Special code into the field at its caret (replacing what is selected there). */
  private insertFindCode(code: string): void {
    const field = this.require<HTMLInputElement>(this.findSpecialTarget === "replace" ? "replacetext" : "findtext");
    const start = field.selectionStart ?? field.value.length;
    const end = field.selectionEnd ?? start;
    field.focus();
    field.setRangeText(code, start, end, "end");
    // A group's "( )" leaves the caret between the brackets.
    if (code === "( )") field.setSelectionRange(start + 1, start + 2);
    field.dispatchEvent(new Event("input", { bubbles: true }));
  }

  private applyFindOptions(options: FindOptions, report: boolean): void {
    const next = normalizeFindOptions(options);
    const changed = JSON.stringify(next) !== JSON.stringify(this.findOptions);
    this.findOptions = next;
    this.syncFindOptionControls();
    if (changed && report) this.options.onFindOptionsChange?.({ ...next });
    if (changed && !this.require("findbar").hidden) this.refreshFind(true);
  }

  /** The checkboxes show the options; those the mode leaves out are greyed, as in Word. */
  private syncFindOptionControls(): void {
    const menu = this.control("findoptmenu");
    if (!menu) return;
    for (const box of Array.from(menu.querySelectorAll<HTMLInputElement>("input[data-findopt]"))) {
      const key = box.dataset.findopt as keyof FindOptions;
      const applies = findOptionApplies(key, this.findOptions);
      box.checked = !!this.findOptions[key] || (key === "matchCase" && !!this.findOptions.wildcards);
      box.disabled = !applies;
      box.closest("label")?.toggleAttribute("data-disabled", !applies);
    }
    const active = FIND_OPTION_KEYS.some((k) => this.findOptions[k] && findOptionApplies(k, this.findOptions));
    this.control("findoptions")?.toggleAttribute("data-active", active);
  }

  /** Find's options (see the `findOptions` option). */
  setFindOptions(options: FindOptions): void {
    this.applyFindOptions(options, false);
  }

  /** Find's options as they are now. */
  getFindOptions(): FindOptions {
    return { ...this.findOptions };
  }

  // ── keyboard shortcuts the engine does not own ──────────────────────────────

  private handleShortcut(event: KeyboardEvent): void {
    // A dialog owns its keys while it is open.
    if (this.paragraphDialog?.open || this.listDialog?.open) return;
    // (the Bookmark and Cross-reference dialogs stay open beside the document: only the keys typed in them are their own)
    const keyTarget = event.target instanceof Node ? event.target : null;
    if (this.bookmarkDialog?.open && this.bookmarkDialog.contains(keyTarget)) return;
    if (this.crossRefDialog?.open && this.crossRefDialog.contains(keyTarget)) return;
    // Alt+F opens the File menu, as in Word.
    if (event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey && (event.key ?? "").toLowerCase() === "f") {
      event.preventDefault();
      this.openFileMenu();
      return;
    }
    // Ctrl+Shift+8 (Ctrl+*): Word's Show/Hide ¶, wherever the focus is.
    if ((event.ctrlKey || event.metaKey) && event.shiftKey && !event.altKey && event.code === "Digit8") {
      event.preventDefault();
      const on = !this.showMarks;
      this.setShowFormattingMarks(on);
      this.options.onCommand?.("show marks", 0, cmd("showMarks", { on }));
      return;
    }
    // Shift+F3: the editor has cycled the case (Word's Change Case cycle).
    if (this.live && event.key === "F3" && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey
      && event.composedPath().includes(this.surface)) {
      this.scheduleDirtyCheck();
      this.options.onCommand?.("change case", 0, cmd("changeCase", { mode: "cycle" }));
      return;
    }
    // F9: Word's Update Field, for the fields in the selection (Ctrl+A first: every field).
    if (this.live && event.key === "F9" && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey
      && event.composedPath().includes(this.surface)) {
      event.preventDefault();
      this.updateFields("selection");
      return;
    }
    if (!this.live || !(event.ctrlKey || event.metaKey)) return;
    // Ctrl+Shift+F5: Word's Bookmark dialog.
    if (event.key === "F5" && event.shiftKey && !event.altKey) {
      event.preventDefault();
      this.openBookmarkDialog();
      return;
    }
    const key = (event.key ?? "").toLowerCase();    // (no key: an autofill pick in a text box)
    const target = event.target instanceof Element ? event.target : null;
    // (by its path, not by where the target is now: the editor's own chords re-render the block they
    // were pressed in, which has left the document by the time the key bubbles up here)
    const inDocument = event.composedPath().includes(this.surface);
    // Ctrl+Alt+M is Word's New Comment; every other chord here is plain Ctrl.
    if (event.altKey) {
      if (key === "m" && inDocument) { event.preventDefault(); this.beginComment(); }
      // Ctrl+Alt+Shift+V: the editor pastes the text only.
      if (key === "v" && event.shiftKey && inDocument && target?.closest('[data-anchor][contenteditable="true"]')) {
        this.scheduleDirtyCheck();
        this.options.onCommand?.("paste", 0, cmd("paste", { formatting: "text" }));
      }
      return;
    }
    // The editor has already acted on its own chords (Undo, Redo, Bold, ...) by now: the key
    // reaches the block before it bubbles up here.
    this.scheduleDirtyCheck();
    // ...and they are reported as the commands they are (onCommand), as the buttons' are: the same
    // chords the editor takes (DocxEditor.onKeydown), in one of its blocks.
    if (inDocument && target?.closest('[data-anchor][contenteditable="true"]')) {
      const own = ({ b: "bold", i: "italic", u: "underline", y: "redo" } as Record<string, string>)[key]
        ?? (key === "z" ? (event.shiftKey ? "redo" : "undo") : undefined);
      if (own) this.options.onCommand?.(own, 0, cmd(own));
      if (key === "v") {
        if (event.shiftKey) this.options.onCommand?.("paste formatting", 0, cmd("pasteFormatting"));
        else this.options.onCommand?.("paste", 0, cmd("paste"));
      }
      if (key === "c" && event.shiftKey) this.options.onCommand?.("copy formatting", 0, cmd("copyFormatting"));
    }
    // Save and Open, as in Word, while the surface shows them (fileActions).
    if (!event.shiftKey && key === "s" && this.control("save")) { event.preventDefault(); this.download(); return; }
    if (!event.shiftKey && key === "o" && this.control("file")) { event.preventDefault(); this.requestOpen(); return; }
    if (key === "f") { event.preventDefault(); this.openFindBar(false); return; }
    if (key === "h") { event.preventDefault(); this.openFindBar(true); return; }
    if (!inDocument) return;
    // Below here every chord edits the focused BLOCK, so a text box inside the surface — the
    // comment gutter's reply and edit boxes — owns its keys: Ctrl+Enter posts a reply there and
    // must not also drop a page break on the commented paragraph, nor Ctrl+E/L/R/J re-align it
    // mid-reply. Find and replace stay reachable from a comment box, as they are in Word.
    if (target?.closest(".docx-comment-gutter, textarea, input")) return;
    const align = (alignment: EditorAlignment) => {
      event.preventDefault();
      this.run(`align ${alignment}`, () => this.live!.setAlignment(alignment), cmd("align", { alignment }));
    };
    switch (key) {
      case "k": event.preventDefault(); this.openLinkPopover(); return;
      case "e": align("center"); return;
      case "l": align("left"); return;
      case "r": align("right"); return;
      case "j": align("justify"); return;
      case "]": event.preventDefault(); this.run("grow font", () => this.live!.adjustFontSize(1), cmd("growFont")); return;
      case "[": event.preventDefault(); this.run("shrink font", () => this.live!.adjustFontSize(-1), cmd("shrinkFont")); return;
      case "enter": event.preventDefault(); this.run("page break", () => this.insertBreakAt("page"), () => this.breakCommand("page")); return;
      default: return;
    }
  }

  // ── popovers ────────────────────────────────────────────────────────────────

  private positionPopover(popover: HTMLElement, anchor: HTMLElement): void {
    if (this.density !== "full") return;
    const rect = anchor.getBoundingClientRect();
    const host = this.element.getBoundingClientRect();
    popover.style.left = `${rect.left - host.left}px`;
    popover.style.top = `${rect.bottom - host.top + 5}px`;
  }

  private wirePicker(): void {
    const button = this.require("table");
    const picker = this.require("gridpicker");
    const cells = this.require("gridcells");

    const highlight = (rows: number, cols: number) => {
      for (const cell of Array.from(cells.children) as HTMLElement[]) {
        const on = Number(cell.dataset.r) < rows && Number(cell.dataset.c) < cols;
        if (on) cell.setAttribute("data-on", "");
        else cell.removeAttribute("data-on");
      }
      this.require("gridlabel").textContent = `${rows} × ${cols}`;
    };

    // `pointerover` covers mouse hover; touch users get the same feedback from the
    // press itself, since a tap fires pointerover immediately before pointerdown.
    cells.addEventListener("pointerover", (event) => {
      const cell = (event.target as HTMLElement).closest<HTMLElement>("[data-r]");
      if (cell) highlight(Number(cell.dataset.r) + 1, Number(cell.dataset.c) + 1);
    });
    cells.addEventListener("mousedown", (event) => {
      const cell = (event.target as HTMLElement).closest<HTMLElement>("[data-r]");
      if (!cell || !this.live) return;
      event.preventDefault();
      const rows = Number(cell.dataset.r) + 1;
      const cols = Number(cell.dataset.c) + 1;
      this.closePopovers();
      const borderless = this.require<HTMLInputElement>("gridborderless").checked;
      const align = this.require<HTMLSelectElement>("gridalign").value as EditorAlignment;
      this.run(`table ${rows}×${cols}`, () =>
        this.live!.insertTable(rows, cols, { borderless, cellAlignment: align }),
        cmd("table", { rows, columns: cols, align, borderless }));
    });

    this.keepSelection(button);
    button.addEventListener("click", () => {
      if (!this.live) return;
      if (picker.hasAttribute("data-open")) {
        this.closePopovers();
        return;
      }
      this.closePopovers();
      highlight(0, 0);
      picker.setAttribute("data-open", "");
      this.positionPopover(picker, button);
    });
  }

  private wireLinkPopover(): void {
    const button = this.require("link");
    const pop = this.require("linkpop");
    const form = this.require<HTMLFormElement>("linkform");
    const url = this.require<HTMLInputElement>("linkurl");
    this.keepSelection(button);
    button.addEventListener("click", () => {
      if (pop.hasAttribute("data-open")) this.closePopovers();
      else this.openLinkPopover();
    });
    this.control("linkcancel")?.addEventListener("click", () => this.closePopovers());
    const place = this.require<HTMLSelectElement>("linkbookmark");
    // A place in the document (a bookmark) or a web address: choosing one sets the other aside.
    place.addEventListener("change", () => { url.disabled = !!place.value; if (!place.value) url.focus(); });
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const bookmark = place.value;
      const target = bookmark || url.value.trim();
      if (!target || !this.live) return;
      this.closePopovers();
      let ok = false;
      const kind = bookmark ? "internal" : "external";
      this.run("link", () => { ok = this.live!.insertHyperlink(target, kind); },
        () => (ok ? cmd("hyperlink", bookmark ? { target, kind } : { target }) : undefined));
      if (!ok) this.setStatus("Select the text to link first");
    });
    url.addEventListener("keydown", (event) => {
      if (event.key === "Escape") { event.preventDefault(); this.closePopovers(); }
    });
  }

  private openLinkPopover(): void {
    if (!this.live) return;
    this.closePopovers();
    const pop = this.require("linkpop");
    const url = this.require<HTMLInputElement>("linkurl");
    const existing = this.live.hyperlinkAtCaret();
    const internal = existing?.kind === "internal";
    url.value = internal ? "" : existing?.target ?? "";
    // (the document's bookmarks, Word's own hidden ones left out, as Word's Link dialog lists them)
    const place = this.require<HTMLSelectElement>("linkbookmark");
    const none = place.options[0];
    place.replaceChildren(none);
    for (const mark of this.live.bookmarks(false)) {
      const option = place.ownerDocument.createElement("option");
      option.value = mark.name;
      option.textContent = mark.name;
      option.title = mark.text;
      place.appendChild(option);
    }
    place.value = internal && existing?.target && Array.from(place.options).some((o) => o.value === existing.target) ? existing.target : "";
    place.disabled = place.options.length < 2;
    url.disabled = !!place.value;
    pop.setAttribute("data-open", "");
    this.positionPopover(pop, this.require("link"));
    if (url.disabled) place.focus();
    else { url.focus(); url.select(); }
  }

  private closePopovers(): void {
    this.control("picturemenu")?.removeAttribute("data-open");
    this.control("fieldmenu")?.removeAttribute("data-open");
    this.control("pastemenu")?.removeAttribute("data-open");
    this.control("pastemore")?.setAttribute("aria-expanded", "false");
    this.control("casemenu")?.removeAttribute("data-open");
    this.control("changecase")?.setAttribute("aria-expanded", "false");
    this.control("findoptmenu")?.removeAttribute("data-open");
    this.control("findoptions")?.setAttribute("aria-expanded", "false");
    this.control("findspecialmenu")?.removeAttribute("data-open");
    this.control("findspecial")?.setAttribute("aria-expanded", "false");
    this.control("gridpicker")?.removeAttribute("data-open");
    this.control("linkpop")?.removeAttribute("data-open");
    this.control("filemenu")?.removeAttribute("data-open");
    this.control("filebutton")?.setAttribute("aria-expanded", "false");
  }

  private maybeClosePopovers(event: MouseEvent): void {
    const target = event.target as Node;
    const picMenu = this.control("picturemenu");
    if (picMenu?.hasAttribute("data-open") && !picMenu.contains(target)) picMenu.removeAttribute("data-open");
    for (const [popName, buttonName] of [["gridpicker", "table"], ["linkpop", "link"], ["filemenu", "filebutton"], ["pastemenu", "pastemore"], ["casemenu", "changecase"], ["findoptmenu", "findoptions"], ["findspecialmenu", "findspecial"]] as const) {
      const pop = this.control(popName);
      if (!pop?.hasAttribute("data-open")) continue;
      if (pop.contains(target) || this.control(buttonName)?.contains(target)) continue;
      pop.removeAttribute("data-open");
      if (popName === "filemenu") this.control("filebutton")?.setAttribute("aria-expanded", "false");
      if (popName === "pastemenu") this.control("pastemore")?.setAttribute("aria-expanded", "false");
      if (popName === "casemenu") this.control("changecase")?.setAttribute("aria-expanded", "false");
      if (popName === "findoptmenu") this.control("findoptions")?.setAttribute("aria-expanded", "false");
      if (popName === "findspecialmenu") this.control("findspecial")?.setAttribute("aria-expanded", "false");
    }
  }

  // ── tabs ────────────────────────────────────────────────────────────────────

  exec(name: string, args: EditorCommandArgs = {}): unknown {
    const live = this.live;
    if (!live) throw new EditorCommandError("closed", name, `${name}: no document is open`);
    // A query only reads: not a command for the rail, the history or a recorder.
    if (QUERY_COMMANDS.has(name)) return live.exec(name, args);
    let result: unknown = null;
    this.run(name, () => { result = live.exec(name, args); }, { name, args: { ...args } });
    // The controls that show what these commands changed (the commands checked their arguments).
    const on = args.on === true || args.on === 1;
    if (name === "pageView") {
      this.require<HTMLInputElement>("paginated").checked = on;
      this.surface.dataset.view = on ? "paginated" : "continuous";
      this.syncViewButtons();
    } else if (name === "zoom" || name === "pageWidth" || name === "onePage") {
      this.syncZoom();
    } else if (name === "trackChanges") {
      this.require<HTMLInputElement>("trackchanges").checked = on;
      this.syncReviewState();
    } else if (name === "showComments") {
      this.require<HTMLInputElement>("showcomments").checked = on;
      this.surface.dataset.comments = on && this.options.comments !== false ? "on" : "off";
    } else if (name === "showMarkup") {
      this.showRevisions = on;
      this.require<HTMLSelectElement>("markup").value = on ? "all" : "none";
    } else if (name === "smallCaps") {
      this.control("smallcaps")?.setAttribute("aria-pressed", String(on));
    } else if (name === "pageSetup") {
      this.syncLayoutState();
    } else if (name === "pageNumbering" || name === "clearPageNumbering") {
      this.syncPageNumbering();
    } else if (name === "headerFooterOption" || name === "goToHeaderFooter" || name === "closeHeaderFooter") {
      this.syncHeaderFooterState();
    }
    return result;
  }

  commands(): EditorCommandInfo[] {
    return DocxEditor.commands();
  }

  selectTab(name: string): void {
    for (const tab of Array.from(this.element.querySelectorAll<HTMLElement>(".dxr-tab[data-tab]"))) {
      tab.setAttribute("aria-selected", String(tab.dataset.tab === name));
    }
    for (const panel of Array.from(this.element.querySelectorAll<HTMLElement>(".dxr-panel"))) {
      if (panel.dataset.panel === name) panel.setAttribute("data-active", "");
      else panel.removeAttribute("data-active");
    }
    if (name === "layout") { this.syncPageNumbering(); this.syncLayoutState(); }
    if (name === "review") this.syncReviewState();
    if (name === "headerfooter") this.syncHeaderFooterState();
    if (name === "view") this.syncZoom();
  }

  // ── selection-driven state ──────────────────────────────────────────────────

  private selectionElement(): HTMLElement | null {
    const selection = (this.element.ownerDocument ?? document).getSelection();
    const node = selection?.anchorNode ?? null;
    if (!node) return null;
    const el = node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
    // Ignore selections belonging to another editor (or another ribbon) on the page.
    return el && this.surface.contains(el) ? el : null;
  }

  private syncSelection(): void {
    if (!this.live || this.destroyed) return;
    this.scheduleRuler();
    const el = this.selectionElement();

    const state = this.live.queryFormatState();
    for (const button of Array.from(this.element.querySelectorAll<HTMLElement>("button[data-cmd]"))) {
      button.classList.toggle("dxr-on", !!state[button.dataset.cmd as FormatKey]);
    }

    const doc = this.element.ownerDocument ?? document;
    if (el && typeof getComputedStyle === "function") {
      const cs = getComputedStyle(el);
      // Reflect the caret's font size, unless the field itself is being edited.
      const fontsize = this.control<HTMLInputElement>("fontsize");
      if (fontsize && doc.activeElement !== fontsize) {
        const px = parseFloat(cs.fontSize);
        if (px) fontsize.value = String(Math.round(px * 0.75 * 2) / 2);
      }
      const hex = toHex(cs.color);
      if (hex) {
        this.control("fontcolorbar")?.style.setProperty("--dxr-swatch", hex);
        const input = this.control<HTMLInputElement>("fontcolor");
        if (input && doc.activeElement !== input) input.value = hex;
      }
      this.control("smallcaps")?.setAttribute("aria-pressed", String(cs.fontVariant.includes("small-caps")));
      const align = cs.textAlign === "start" ? "left" : cs.textAlign === "end" ? "right" : cs.textAlign;
      for (const button of Array.from(this.element.querySelectorAll<HTMLElement>("button[data-align]"))) {
        button.classList.toggle("dxr-on", button.dataset.align === align);
      }
    }
    this.syncStyleBox();
    const listFormat = this.live.listFormatAtCaret();
    for (const button of Array.from(this.element.querySelectorAll<HTMLElement>("button[data-list]"))) {
      const kind = button.dataset.list;
      button.classList.toggle(
        "dxr-on",
        !!listFormat && (kind === "bullet" ? listFormat.startsWith("bullet") : !listFormat.startsWith("bullet")),
      );
    }

    // Reveal the contextual Table tab only while the caret is inside a table. Typing in one of the
    // ribbon's own boxes (the Table tab's Height and Width) takes the selection out of the document:
    // the tab stays as it is then, or it would hide the box being typed in.
    const inTable = !!el?.closest("table");
    const focus = doc.activeElement;
    const inChrome = !el && !!focus && focus !== doc.body && this.element.contains(focus) && !this.surface.contains(focus);
    const tableTab = this.element.querySelector<HTMLElement>('.dxr-tab[data-tab="table"]');
    if (tableTab && !inChrome) {
      tableTab.hidden = !inTable;
      if (!inTable && tableTab.getAttribute("aria-selected") === "true") this.selectTab("home");
      if (inTable) {
        const row = el?.closest("tr");
        this.require<HTMLInputElement>("repeatheader").checked = !!row?.classList.toString().includes("header");
        this.tableChanged();
      }
    }
    this.refreshRailAnchor();
    this.scheduleStats();
  }

  /** Scope is read from where the block lives, which is exactly what the anchor encodes. */
  private scopeOf(el: HTMLElement): string {
    const band = el.closest<HTMLElement>("[data-hf-band]");
    if (band) return band.getAttribute("data-hf-band") === "header" ? "hdr" : "ftr";
    if (el.closest(".footnotes")) return "fn";
    if (el.closest(".endnotes")) return "en";
    return "body";
  }

  private kindOf(el: HTMLElement): string {
    if (/^H[1-6]$/.test(el.tagName)) return "h";
    if (el.tagName === "LI" || el.hasAttribute("data-list-marker") || el.querySelector("[data-list-marker]")) {
      return "li";
    }
    return "p";
  }

  /**
   * Block count and session handle — an O(blocks) query, so it runs only when
   * something could have changed them (a command, or a document opening), never on
   * the selectionchange path that fires per keystroke.
   */
  private refreshRailCounts(): void {
    const blocks = this.control("railBlocks");
    if (!blocks) return; // rail disabled
    blocks.textContent = this.live ? String(this.surface.querySelectorAll("[data-anchor]").length) : "—";
    const session = this.control("railSession");
    if (session) session.textContent = this.live ? `#${this.live.sessionHandle}` : "—";
  }

  /** The focused anchor — cheap, and the only part the caret can change. */
  private refreshRailAnchor(): void {
    const anchorEl = this.control("railAnchor");
    if (!anchorEl) return;
    const block = this.selectionElement()?.closest<HTMLElement>("[data-anchor]") ?? null;
    if (!block) {
      anchorEl.textContent = this.live ? "none" : "—";
      this.lastAnchorText = "";
      return;
    }
    const unid = block.getAttribute("data-anchor") ?? "";
    const text = `${this.kindOf(block)}:${this.scopeOf(block)}:${unid.slice(0, 10)}…`;
    if (text === this.lastAnchorText) return;
    anchorEl.textContent = text;
    this.lastAnchorText = text;
    anchorEl.classList.remove("dxr-flash");
    void anchorEl.offsetWidth;
    anchorEl.classList.add("dxr-flash");
  }

  /** Page position and word count are O(document); debounced off the hot paths. */
  private scheduleStats(): void {
    if (this.statsTimer != null) clearTimeout(this.statsTimer);
    this.statsTimer = setTimeout(() => {
      this.statsTimer = null;
      this.refreshStats();
    }, 150);
  }

  private refreshStats(): void {
    if (!this.live || this.destroyed) return;
    const pageinfo = this.control("pageinfo");
    if (pageinfo) {
      const info = this.live.pageInfo();
      pageinfo.textContent = info ? `Page ${info.page} of ${info.total}` : "Continuous";
    }
    const words = this.control("wordcount");
    if (words) {
      const n = this.live.wordCount();
      words.textContent = `${n.toLocaleString()} word${n === 1 ? "" : "s"}`;
    }
  }

  /** Font families the document's rendered text actually uses, most common first. */
  private documentFonts(): string[] {
    const counts = new Map<string, number>();
    for (const el of Array.from(this.surface.querySelectorAll<HTMLElement>("[data-anchor] [style*='font-family'], [data-anchor][style*='font-family']"))) {
      const family = el.style.fontFamily.split(",")[0]?.trim().replace(/^['"]|['"]$/g, "");
      if (!family || /^(serif|sans-serif|monospace)$/i.test(family)) continue;
      counts.set(family, (counts.get(family) ?? 0) + 1);
    }
    return Array.from(counts.entries()).sort((a, b) => b[1] - a[1]).map(([name]) => name).slice(0, 12);
  }

  // ── loading overlay ─────────────────────────────────────────────────────────

  private createLoaderController(): RibbonLoader {
    const overlay = this.control("loader");
    if (!overlay) {
      const noop = () => {};
      return { show: noop, stage: noop, progress: noop, done: noop, fail: noop };
    }

    const options = this.loaderOptions ?? {};
    const eyebrow = this.control("loaderEyebrow");
    if (eyebrow && options.eyebrow) eyebrow.textContent = options.eyebrow;
    const meta = this.control("loaderMeta");
    if (meta && options.meta) meta.textContent = options.meta;
    this.control("loaderRetry")?.addEventListener("click", () => {
      if (options.onRetry) options.onRetry();
      else location.reload();
    });
    if (this.features.length === 0) this.control("loaderAd")?.remove();

    const setProgress = (percent: number, label?: string) => {
      const bar = this.control("loaderBar");
      if (bar) bar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
      if (label) {
        const el = this.control("loaderLabel");
        if (el) el.textContent = label;
      }
    };

    return {
      show: () => {
        overlay.hidden = false;
        overlay.removeAttribute("data-done");
        overlay.removeAttribute("data-error");
        this.setState("loading");
        this.showFeature(0);
        this.startRotation();
        this.loaderStage(0);
      },
      stage: (step) => this.loaderStage(step),
      progress: setProgress,
      done: () => {
        this.stopRotation();
        this.setState("ready");
        // Two-step so the fade actually plays before the overlay leaves the layout.
        setTimeout(() => overlay.setAttribute("data-done", ""), 420);
        setTimeout(() => {
          overlay.hidden = true;
        }, 1050);
      },
      fail: (error) => {
        this.stopRotation();
        this.setState("error");
        overlay.hidden = false;
        overlay.setAttribute("data-error", "");
        overlay.removeAttribute("data-done");
        const message = error instanceof Error ? error.message : String(error);
        const title = this.control("loaderTitle");
        if (title) title.textContent = "The local engine did not start";
        const copy = this.control("loaderCopy");
        if (copy) copy.textContent = message.slice(0, 220);
        setProgress(100, "Load failed");
        this.setStatus(message.slice(0, 180));
      },
    };
  }

  private loaderStage(step: number | Partial<RibbonLoaderStage>): void {
    const stage = typeof step === "number" ? this.stages[Math.max(0, Math.min(this.stages.length - 1, step))] : step;
    if (!stage) return;
    const title = this.control("loaderTitle");
    if (title && stage.title) title.textContent = stage.title;
    const copy = this.control("loaderCopy");
    if (copy && stage.copy) copy.textContent = stage.copy;
    const bar = this.control("loaderBar");
    if (bar && stage.progress != null) bar.style.width = `${stage.progress}%`;
    const label = this.control("loaderLabel");
    if (label && stage.label) label.textContent = stage.label;
  }

  private showFeature(index: number): void {
    const feature = this.features[index];
    if (!feature) return;
    const card = this.control("loaderAd");
    const number = this.control("loaderNumber");
    const title = this.control("loaderAdTitle");
    const copy = this.control("loaderAdCopy");
    if (number) number.textContent = feature.number;
    if (title) title.textContent = feature.title;
    if (copy) copy.textContent = feature.copy;
    if (card) {
      card.removeAttribute("data-swap");
      void card.offsetWidth;
      card.setAttribute("data-swap", "");
    }
  }

  private startRotation(): void {
    this.stopRotation();
    if (this.features.length < 2) return;
    const every = this.loaderOptions?.rotateMs ?? 1750;
    this.featureTimer = setInterval(() => {
      this.featureIndex = (this.featureIndex + 1) % this.features.length;
      this.showFeature(this.featureIndex);
    }, every);
  }

  private stopRotation(): void {
    if (this.featureTimer == null) return;
    clearInterval(this.featureTimer);
    this.featureTimer = null;
  }
}
