/**
 * DocxEditor — a framework-agnostic, in-browser DOCX block editor.
 *
 * Architecture (see docs/architecture/ir_editor_feasibility.md, "Option B"):
 *   - model-of-record: a live DocxSession in WASM (lossless save);
 *   - rendering: WmlToHtmlConverter HTML (faithful) stamped with data-anchor;
 *   - editing: each block is contenteditable; on commit, the edit goes through
 *     DocxSession by anchor, then ONLY that block is re-rendered from the live
 *     session (session-attached RenderBlockHtml) and patched into the DOM.
 *
 * The IR/anchor system is the addressing spine; the live OOXML is the truth.
 * This is the pure-TypeScript core; a React wrapper can sit on top.
 *
 * MVP scope: per-block, commit-on-blur editing of paragraphs/headings. An edited
 * block's content is replaced from its plain text (inline formatting within an
 * edited block is not preserved — a documented MVP limit); UNTOUCHED blocks keep
 * full fidelity, and save() is lossless for them.
 */

import { liftedDrawingHomes, paginateHtml } from "./pagination.js";
import { drawLineNumbers, hasLineNumbers } from "./line-numbers.js";
import { applyUnroundedNormalLineHeights } from "./line-metrics.js";
import { DocumentViewport } from "./viewport.js";
import type { ColumnWidth } from "./viewport.js";
import { parseSectionDimensions } from "./page-geometry.js";
import { HeaderFooterRegion } from "./editor-headerfooter.js";
import type { BandWhich } from "./editor-headerfooter.js";
import { CommentGutter, COMMENT_GUTTER_CSS } from "./editor-comments.js";
import type { CommentTarget } from "./editor-comments.js";
import {
  draggable,
  dropTargetForElements,
  monitorForElements,
} from "@atlaskit/pragmatic-drag-and-drop/element/adapter";
import { setCustomNativeDragPreview } from "@atlaskit/pragmatic-drag-and-drop/element/set-custom-native-drag-preview";
import { pointerOutsideOfPreview } from "@atlaskit/pragmatic-drag-and-drop/element/pointer-outside-of-preview";
import {
  autoScrollForElements,
  autoScrollWindowForElements,
} from "@atlaskit/pragmatic-drag-and-drop-auto-scroll/element";
import { TrackedChangeMode } from "./types.js";
import { requireRevisionIdentitySupport } from "./session.js";
import type {
  CommentListEntry,
  EditErrorCode,
  EditResult,
  FormattingInspection,
  InlineSpan,
  HeaderFooterKind,
  HyperlinkInfo,
  ImageInsertOptions,
  ImageOccurrence,
  FloatingImageLayout,
  ListFormat,
  ListLevelInfo,
  ListLevelSpec,
  NumberFormat,
  PageSetupOp,
  TabStop,
  TabStopsInfo,
  ColumnsOp,
  ColumnsApplyTo,
  ParagraphFormatting,
  RevisionListEntry,
  RevisionIdentity,
  RevisionAuthorNameInfo,
  SectionInfo,
  LineNumbering,
  SectionStart,
  StyleInfo,
  TableBorderSpec,
  TablePropertiesOp,
  TablePropertiesInfo,
  TableRowOptions,
  CellPropertiesOp,
  TableCellScope,
  TableOfContentsOptions,
} from "./types.js";
import { diffUnits, needsRemount, tokenOf, unidOf } from "./editor-reconcile.js";
import { EditorCommandError, editorCommands, runEditorCommand } from "./editor-commands.js";
import type { EditorCommandArgs, EditorCommandInfo } from "./editor-commands.js";
import { imageOnlyDelta, patchImageAttributes } from "./editor-image-patch.js";
import { PictureOverlay } from "./editor-picture.js";
import { cleanPastedHtml, pastedPlainText, serializePaste } from "./paste-html.js";
import { FORMAT_PAINTER_CSS, pickRunFormat, runOpFromDirect, wordSpanAt } from "./format-painter.js";
import { compileFind, expandReplacement, type FindOptions } from "./find-query.js";
import { changeCase, DEFAULT_HEADLINE_CASE_LISTS, nextCycleCase, type CaseMode, type HeadlineCaseLists } from "./change-case.js";
import type { RenderPlan, RenderUnit, UnitDiff } from "./editor-reconcile.js";

/** The subset of WASM bridge exports the editor needs (as exposed on `window.Docxodus`). */
export interface DocxEditorExports {
  DocxSessionBridge: {
    GetVersion?: (handle: number) => string;
    OpenSession: (bytes: Uint8Array, settingsJson: string) => number;
    CloseSession: (handle: number) => void;
    CreateBlankDocx: () => Uint8Array;
    Project: (handle: number) => string;
    ReplaceText: (handle: number, anchor: string, md: string) => string;
    ReplaceTextAtSpan: (
      handle: number,
      anchor: string,
      spanStart: number,
      spanLength: number,
      replace: string,
      deletedTextSide?: string,
    ) => string;
    SplitParagraph: (handle: number, anchor: string, offset: number) => string;
    MergeParagraphs: (handle: number, first: string, second: string) => string;
    /** An anchor's live facts (`AnchorInfo` JSON); its visibleText checks a merge's join. */
    GetAnchorInfo?: (handle: number, anchorId: string) => string;
    DeleteBlock: (handle: number, anchor: string) => string;
    /** Delete the sibling blocks from `fromAnchor` up to (not including) `toAnchorExclusive`. */
    DeleteRange?: (handle: number, fromAnchor: string, toAnchorExclusive: string) => string;
    BeginTransaction?: (handle: number) => number;
    CommitTransaction?: (transactionHandle: number) => void;
    RollbackTransaction?: (transactionHandle: number) => void;
    MoveBlock?: (handle: number, sourceAnchor: string, targetAnchor: string, pos: string) => string;
    /** JSON `{anchorId, before, after}[]`: where a block may legally move, per side. Optional —
     *  without it the drag UI offers every block and lets the engine refuse, as it did before. */
    ValidMoveTargets?: (handle: number, sourceAnchor: string) => string;
    InsertHorizontalRule: (handle: number, anchor: string, pos: string, ruleJson: string) => string;
    InsertTable: (
      handle: number,
      anchor: string,
      pos: string,
      rows: number,
      cols: number,
      optionsJson: string,
    ) => string;
    InsertTableRow: (handle: number, cellAnchor: string, pos: string) => string;
    InsertTableColumn: (handle: number, cellAnchor: string, pos: string) => string;
    DeleteTableRow: (handle: number, cellAnchor: string) => string;
    DeleteTableColumn: (handle: number, cellAnchor: string) => string;
    ApplyFormat: (handle: number, anchor: string, spanJson: string, opJson: string) => string;
    SetParagraphStyle: (handle: number, anchor: string, styleId: string) => string;
    SetParagraphFormat: (handle: number, anchor: string, opJson: string) => string;
    ApplyListFormat: (handle: number, anchor: string, kind: string) => string;
    SetListLevel: (handle: number, anchor: string, delta: number) => string;
    GetListMembership: (handle: number, anchor: string) => string;
    GetListLevels?: (handle: number, anchor: string) => string;
    SetListLevels?: (handle: number, anchor: string, levelsJson: string) => string;
    ContinueNumbering?: (handle: number, anchor: string) => string;
    SetListMembership?: (handle: number, anchor: string, numId: number, level: number) => string;
    SetListStartOverride?: (handle: number, anchor: string, value: number) => string;
    RenderBlockHtml: (
      handle: number,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean,
    ) => string;
    RenderBlockHtmlForReview?: (
      handle: number,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      renderTrackedChanges: boolean,
    ) => string;
    /** Session-attached full-document render (optional: older WASM bundles lack it). */
    RenderHtml?: (
      handle: number,
      cssPrefix: string,
      fabricateClasses: boolean,
      paginated: boolean,
      scale: number,
    ) => string;
    RenderHtmlForReview?: (
      handle: number,
      cssPrefix: string,
      fabricateClasses: boolean,
      paginated: boolean,
      scale: number,
      renderTrackedChanges: boolean,
    ) => string;
    Save: (handle: number) => Uint8Array;
    /** Save keeping the projector's Unid bookkeeping — remount only, never a user download.
     *  Optional so a bridge predating it still satisfies this type. */
    SaveWithAnchorIds?: (handle: number) => Uint8Array;
    Undo: (handle: number) => boolean;
    Redo: (handle: number) => boolean;
    /** Header/footer region: the section a body anchor belongs to (kind → part mapping). */
    GetSectionInfo: (handle: number, anchorId: string) => string;
    SetHeaderText: (handle: number, anchor: string, kind: string, markdown: string) => string;
    SetFooterText: (handle: number, anchor: string, kind: string, markdown: string) => string;
    InsertPageNumberField: (handle: number, anchor: string, field: string, format: string) => string;
    EnsureHeaderFooterVisible: (handle: number, anchor: string, kind: string) => string;
    SetPageNumbering: (handle: number, anchor: string, opJson: string) => string;
    ClearPageNumbering: (handle: number, anchor: string) => string;
    /** Note authoring (optional: older WASM bundles predate it). */
    InsertFootnote?: (
      handle: number,
      anchor: string,
      characterOffset: number,
      markdown: string,
    ) => string;
    InsertEndnote?: (
      handle: number,
      anchor: string,
      characterOffset: number,
      markdown: string,
    ) => string;
    /** Native comment authoring (optional: older WASM bundles predate it; issue #580). */
    AddComment?: (
      handle: number,
      anchor: string,
      spanJson: string,
      author: string,
      initials: string,
      date: string,
      markdown: string,
    ) => string;
    SetCommentResolved?: (handle: number, commentAnchor: string, resolved: boolean) => string;
    ListComments?: (handle: number) => string;
    AddCommentReply?: (
      handle: number,
      parentCommentAnchor: string,
      author: string,
      initials: string,
      date: string,
      markdown: string,
    ) => string;
    UpdateComment?: (handle: number, commentAnchor: string, markdown: string) => string;
    RemoveComment?: (handle: number, commentAnchor: string) => string;
    /** The editor's own render profile as one JSON object — the comment-aware twin of
     *  RenderHtml / RenderBlockHtml / RenderBlocksHtml (optional: older bundles lack it). */
    RenderEditorHtml?: (handle: number, optionsJson: string) => string;
    RenderEditorBlockHtml?: (handle: number, anchorId: string, optionsJson: string) => string;
    RenderEditorBlocksHtml?: (handle: number, anchorIdsJson: string, optionsJson: string) => string;
    RenderEditorChromeHtml?: (handle: number, optionsJson: string) => string;
    RenderEditorRangeHtml?: (handle: number, anchorIdsJson: string, optionsJson: string) => string;
    /** Review-mode controls (optional). */
    SetTrackedChanges?: (handle: number, mode: number) => void;
    SetRevisionAuthor?: (handle: number, author: string) => void;
    SetRevisionIdentities?: (handle: number, identitiesJson: string) => void;
    GetRevisionAuthorNameStatus?: (handle: number, author: string) => string;
    ListRevisions?: (handle: number) => string;
    AcceptRevision?: (handle: number, revisionId: string) => string;
    RejectRevision?: (handle: number, revisionId: string) => string;
    AcceptAllRevisions?: (handle: number) => string;
    RejectAllRevisions?: (handle: number) => string;
    /** Links, images, references (optional). */
    ListHyperlinks?: (handle: number, scopes: number) => string;
    AddHyperlink?: (
      handle: number,
      anchor: string,
      start: number,
      length: number,
      kind: string,
      target: string,
    ) => string;
    RemoveHyperlink?: (handle: number, hyperlinkId: string) => string;
    InsertImage?: (
      handle: number,
      anchor: string,
      characterOffset: number,
      imageBase64: string,
      optionsJson: string,
    ) => string;
    ListImages?: (handle: number, scopes: number) => string;
    SetImageDimensions?: (handle: number, imageId: string, dimensionsJson: string) => string;
    SetImageMetadata?: (handle: number, imageId: string, altText: string | null, title: string | null) => string;
    SetImageFloatingLayout?: (handle: number, imageId: string, layoutJson: string) => string;
    SetImagePlacement?: (handle: number, imageId: string, placement: string, layoutJson: string | null) => string;
    ReplaceImage?: (handle: number, imageId: string, imageBase64: string) => string;
    RemoveImage?: (handle: number, imageId: string) => string;
    InsertTableOfContents?: (handle: number, anchor: string, pos: string, optionsJson: string) => string;
    /** Paste (optional): the clipboard's HTML as XHTML, or plain text, at a caret. */
    InsertHtml?: (
      handle: number,
      anchor: string,
      characterOffset: number,
      replaceLength: number,
      html: string,
      optionsJson: string,
    ) => string;
    InsertPastedText?: (handle: number, anchor: string, characterOffset: number, replaceLength: number, text: string) => string;
    /** Table cell structure and appearance (optional). */
    MergeCells?: (handle: number, cellAnchor: string, rowSpan: number, colSpan: number, content: string) => string;
    UnmergeCells?: (handle: number, cellAnchor: string) => string;
    SetTableBorders?: (handle: number, cellAnchor: string, specJson: string) => string;
    SetCellShading?: (handle: number, cellAnchor: string, fill: string, scope: string) => string;
    SetRepeatHeaderRow?: (handle: number, cellAnchor: string, repeat: boolean) => string;
    SetTableRowOptions?: (handle: number, cellAnchor: string, repeatHeader: boolean | null,
      allowBreakAcrossPages: boolean | null, heightTwips: number | null, heightRule: string) => string;
    SetColumnWidths?: (handle: number, cellAnchor: string, widthsJson: string) => string;
    /** Table Properties (optional: older engine builds lack them). */
    SetTableProperties?: (handle: number, cellAnchor: string, opJson: string) => string;
    SetCellProperties?: (handle: number, cellAnchor: string, opJson: string, scope: string) => string;
    GetTableProperties?: (handle: number, cellAnchor: string) => string;
    /** Section page setup and header/footer flags (optional). */
    SetPageSetup?: (handle: number, anchor: string, opJson: string) => string;
    /** Tab stops (optional: older engine builds lack them). */
    SetTabStops?: (handle: number, anchorId: string, stopsJson: string) => string;
    GetTabStops?: (handle: number, anchorId: string) => string;
    SetDefaultTabStop?: (handle: number, twips: number) => string;
    /** Word's Layout > Breaks (optional: absent on older bundles). */
    InsertSectionBreak?: (handle: number, anchor: string, offset: number, sectionStart: string) => string;
    RemoveSectionBreak?: (handle: number, anchor: string) => string;
    InsertBreak?: (handle: number, anchor: string, offset: number, kind: string) => string;
    /** Word's Layout > Columns (optional: absent on older bundles). */
    SetColumns?: (handle: number, anchor: string, offset: number, opJson: string, applyTo: string, endAnchor: string, endOffset: number) => string;
    SetHeaderFooterKindEnabled?: (handle: number, anchor: string, kind: string, enabled: boolean) => string;
    /** Introspection (optional). */
    ListStyles?: (handle: number) => string;
    GetFormatting?: (handle: number, anchorId: string) => string;
    ReplaceTextRange?: (handle: number, anchor: string, find: string, replace: string, optionsJson: string) => string;
    /** Incremental-reconcile endpoints (optional: older WASM bundles predate them;
     *  the editor falls back to full remounts / full projections without them). */
    ListBlocks?: (handle: number) => string;
    ListRenderedBlocks?: (handle: number, renderTrackedChanges: boolean) => string;
    ListNotes?: (handle: number, endnotes: boolean) => string;
    ListAnchors?: (handle: number) => string;
    RenderBlocksHtml?: (
      handle: number,
      anchorIdsJson: string,
      cssPrefix: string,
      fabricateClasses: boolean,
    ) => string;
    RenderBlocksHtmlForReview?: (
      handle: number,
      anchorIdsJson: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      renderTrackedChanges: boolean,
    ) => string;
  };
  DocumentConverter: {
    ConvertDocxToHtmlComplete: (...args: any[]) => string;
  };
}

export interface DocxEditorOptions {
  /** CSS class prefix for rendered HTML. Default "docx-". */
  cssPrefix?: string;
  /**
   * Fabricate CSS classes (vs inline styles). Default FALSE for the editor: a per-block
   * re-render must be self-contained, but fabricated class names are per-conversion and
   * have no matching stylesheet on the page, so re-rendered blocks would lose styling.
   * Inline styles keep every block's formatting intact on incremental re-render.
   */
  fabricateClasses?: boolean;
  /** Make paragraph/heading blocks editable. Default true. */
  editable?: boolean;
  /** Render block-flow pages (page boxes via pagination.ts) vs a continuous view. Default false. */
  paginated?: boolean;
  /** Author-pinned page render scale (1.0 = 100%). Default 1. */
  scale?: number;
  /**
   * How the continuous view sizes its text column. `"section"` (default) uses the width the
   * document's `w:sectPr` defines, so line breaking matches Word on every screen; `"fluid"`
   * lets the column follow the host, the pre-geometry behavior.
   */
  columnWidth?: ColumnWidth;
  /**
   * Zoom the page down to fit a host narrower than it, the way a word processor's
   * fit-to-width does, instead of letting it overflow. Default true.
   */
  fitToWidth?: boolean;
  /**
   * Render docked Header/Footer editing bands around the body flow. Default FALSE — with it off
   * the editor's DOM is unchanged, so existing consumers are unaffected. When on, the body flow
   * is wrapped in a `.docx-body-flow` element that becomes the edit root, keeping band blocks out
   * of the body's block list (which indexes remount focus).
   */
  headerFooter?: boolean;
  /** Enable the editor-owned block drag handle. Default false (the ribbon defaults it to true). */
  blockDrag?: boolean;
  /** How editor mutations are recorded. RenderInline enables native Word revisions. */
  trackedChanges?: TrackedChangeMode;
  /**
   * Whether the document's tracked revisions render inline (`ins`/`del`) or as if every change
   * were accepted, independently of how `trackedChanges` records edits. Unset, it follows
   * `trackedChanges`: shown under RenderInline, hidden otherwise.
   */
  showRevisions?: boolean;
  /** Author stamped on native Word revisions. Default "docxodus". */
  revisionAuthor?: string;
  /** Account identifiers for ownership; absent, null or empty retains display-name ownership. */
  revisionIdentities?: readonly RevisionIdentity[] | null;
  /**
   * Render comments Word-style: the commented range highlighted inline and each thread as a
   * bubble in a gutter beside the page. Default true. With it off the document renders without
   * comment markup at all (the pre-gutter behaviour).
   */
  comments?: boolean;
  /** Author stamped on comments posted from the gutter. Defaults to `revisionAuthor`. */
  commentAuthor?: string;
  /**
   * How typing becomes undo steps. `"block"` (the default): a block's typing commits when the
   * block loses focus or a command runs, so one Undo takes back everything typed in it since.
   * `"word"`, as Word does: the typing also commits as each word ends (a space typed after it) and
   * after a pause of a second, so Undo takes it back a word at a time. Each commit is one session
   * edit, so `"word"` suits an engine that answers in a few milliseconds.
   */
  typingUndo?: "block" | "word";
  /**
   * The levels every NEW list starts from (Word's "set as default" for a multilevel list): when
   * the Numbering button or a numbering format makes paragraphs a list, these are written into
   * it. A numbered list (`decimal`) takes them whole; any other format (bullets, letters) takes
   * only their positions. Unset, a new list has the engine's own levels. `setListDefaults`
   * changes them later.
   */
  listDefaults?: ListLevelSpec[];
  /**
   * Headline Case's word lists (small words, always lower case, always as written); unset, the
   * built-in ones ({@link DEFAULT_HEADLINE_CASE_LISTS}). A host that keeps them in a file reads
   * it with {@link parseHeadlineCaseLists}. `setHeadlineCaseLists` changes them later.
   */
  headlineCaseLists?: HeadlineCaseLists;
  /**
   * Word's Show/Hide ¶: a ¶ at the end of every paragraph and an arrow in every tab, drawn over
   * the text (nothing moves, and nothing is saved). `setShowFormattingMarks` changes it later.
   */
  showFormattingMarks?: boolean;
  /**
   * Initials stamped on `commentAuthor`'s comments and replies (`w:initials`, shown in Word's
   * balloons). Defaults to the first letter of each word of the author's name, as Word makes them.
   */
  commentInitials?: string;
  /** Called after a block edit commits (with the affected anchor). */
  onEdit?: (info: { anchorId: string; unid: string }) => void;
  /**
   * Called when the engine rejects an edit. The editor has already put the affected block back
   * to the session's content, so the screen never shows text the document does not hold. When no
   * handler is set, the editor logs the rejection with `console.warn` instead.
   */
  onEditFailed?: (info: EditFailure) => void;
  /**
   * What Ctrl+V does with the source's formatting: `"merge"` (the default, Word's Merge Formatting:
   * paragraphs, headings, lists, tables, links and emphasis come along, fonts, sizes and colors are
   * the document's), `"keep"` (Keep Source Formatting) or `"text"` (Keep Text Only).
   * Ctrl+Alt+Shift+V always pastes the text only (Ctrl+Shift+V pastes copied formatting, as in Word).
   */
  pasteFormatting?: PasteFormatting;
  /** Called when Format Painter turns on (`"once"`, `"sticky"`) or off, so a button can show it. */
  onFormatPainterChange?: (state: FormatPainterState) => void;
  /** Called when Show/Hide ¶ turns on or off, so a button can show it. */
  onFormattingMarksChange?: (on: boolean) => void;
  /** Called when a picture is selected or deselected, and after the selected picture changes. */
  onPictureChange?: (picture: PictureInfo | null) => void;
  /**
   * Find the bytes of a pasted picture the page cannot fetch itself, such as Word's
   * `file:///…/clip_image001.png`: a `data:` URL, or null to leave the picture out. Pictures on
   * the web are fetched by the page; without this, local ones are left out.
   */
  resolvePasteImage?: (src: string) => Promise<string | null>;
  /** Called after a successful block move. */
  onMove?: (info: { sourceAnchorId: string; destinationAnchorId: string }) => void;
  /** Called when the caret enters or leaves a header/footer story (`null` = back in the body). */
  onStoryChange?: (which: BandWhich | null) => void;
  /** Called after every comment-gutter layout with thread counts and the active thread. */
  onCommentsChange?: (info: { threads: number; open: number; active: string | null }) => void;
}

/** Options for {@link DocxEditor.openAsync}: the mount's window size and a progress callback. */
export interface DocxEditorOpenAsyncOptions extends DocxEditorOptions {
  /**
   * Body units (top-level paragraphs and tables) mounted per task before yielding to the event
   * loop. Default 24 — on the reference 17-page document that is a ~50 ms task. A window is
   * extended past the size so it never ends inside a border box the renderer groups adjacent
   * paragraphs into.
   */
  windowSize?: number;
  /** Called after each window lands with the count of body units mounted so far. */
  onProgress?: (mounted: number, total: number) => void;
}

/**
 * Word's Page Setup, as {@link DocxEditor.setPageSetup} takes it — the session's own
 * {@link PageSetupOp} (all twips; omit = unchanged), so the two cannot drift.
 */
export type EditorPageSetup = PageSetupOp;

/**
 * What a ruler shows for the caret's paragraph ({@link DocxEditor.rulerState}). Positions are
 * client pixels (as `getBoundingClientRect` gives them); lengths are twips.
 */
export interface EditorRulerState {
  /** The page's edges: the page box in page view, the section's sheet in continuous view. */
  pageLeft: number;
  pageRight: number;
  /** The section's text column (inside the left and right margins). */
  textLeft: number;
  textRight: number;
  /** Where the paragraph's indent 0 lies, left and right: the text column, a table cell's text
   *  area or a newspaper column. Tab positions count from `originLeft` too. */
  originLeft: number;
  originRight: number;
  /** How far the paragraph's indents may reach: a table cell's edges, else the page's. */
  boundsLeft: number;
  boundsRight: number;
  /** In a table cell: the caret row's column boundaries (client x of each cell's left edge after
   *  the first, then the last cell's right edge), each with the grid column it starts (the last:
   *  the grid's column count), and the table grid's widths in twips. */
  table?: { boundaries: { x: number; grid: number }[]; widthsTwips: number[] };
  /** Screen pixels per twip at the current zoom. */
  pxPerTwip: number;
  /** True when the paragraph's origin is the section's text column, so dragging a margin edge
   *  changes the section's page margins. */
  inSectionColumn: boolean;
  marginLeftTwips: number;
  marginRightTwips: number;
  leftIndent: number;
  rightIndent: number;
  /** The first line's indent from `leftIndent`: positive for a first-line indent, negative for a
   *  hanging one. */
  firstLine: number;
  /** The tab stops in effect (from the active paragraph), and the default interval after them. */
  tabs: Required<TabStop>[];
  defaultTabStopTwips: number;
}

/** One hit from {@link DocxEditor.find}: a content-offset span inside an editable block. */
export interface EditorMatch {
  block: HTMLElement;
  start: number;
  /** Length in `block` (to the block's end when the hit runs on into `endBlock`). */
  length: number;
  /** A hit over paragraph marks ends in this later block of the same story... */
  endBlock?: HTMLElement;
  /** ...at this content offset. */
  endOffset?: number;
  /** The blocks wholly inside a hit that runs over paragraph marks. */
  between?: HTMLElement[];
  /** The text found, in stream characters (a paragraph mark is a carriage return). */
  text?: string;
  /** The replacement text for this hit, read as the find that made it reads one (^p, ^&, $1...). */
  expand?: (replacement: string) => string;
}

interface AnchorTargetLite {
  unid: string;
  kind: string;
  scope: string;
  textPreview?: string;
}

const EDITABLE_TAGS = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6"]);

const BLOCK_DRAG_TYPE = "docxodus-block";
const blockDragStyledDocuments = new WeakSet<Document>();

/** One-line description of a block, for the handle's label and the drag preview chip. */
function blockPreviewText(unit: HTMLElement): string {
  if (unit.tagName === "TABLE") return "Table";
  const text = (unit.textContent ?? "").trim().replace(/\s+/g, " ");
  return text.length > 48 ? `${text.slice(0, 48)}…` : text;
}

/**
 * A capture-time snapshot of one movable block's box, in viewport coordinates.
 *
 * Taken once per drag: nothing reflows while a drag is in flight, so re-measuring every block on
 * every `dragover` would buy nothing. Scrolling (including drag autoscroll) moves the whole flow
 * rigidly, which `dropZoneShift` corrects for with one subtraction.
 */
interface BlockDropZone {
  unit: HTMLElement;
  anchorId: string;
  /** Position in `dropZones`, so a neighbour is one index away — see `dropEdgeY`. */
  index: number;
  top: number;
  bottom: number;
  left: number;
  width: number;
  /** Own outer margin, read only for the first/last zone — see `dropEdgeY`. */
  marginBefore: number;
  marginAfter: number;
}

/** The floating handle's own box and its gap to the block, as `ensureBlockDragStyles` sizes it. */
const BLOCK_HANDLE_WIDTH = 26;
const BLOCK_HANDLE_HEIGHT = 28;
const BLOCK_HANDLE_GAP = 6;

/** `value` pulled into [`low`, `high`]; `low` wins when the range is narrower than the box. */
function clampToRange(value: number, low: number, high: number): number {
  return Math.max(low, Math.min(value, high));
}

function ensureBlockDragStyles(doc: Document): void {
  if (blockDragStyledDocuments.has(doc)) return;
  blockDragStyledDocuments.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusBlockDrag = "true";
  style.textContent = `
.docx-block-handle {
  position: fixed; z-index: 2147483000; display: none; width: 26px; height: 28px;
  align-items: center; justify-content: center; padding: 0; border: 1px solid #d0d5dd;
  border-radius: 6px; background: #fff; color: #667085; box-shadow: 0 1px 3px rgba(16,24,40,.12);
  cursor: grab; font: 700 17px/1 system-ui, sans-serif; user-select: none; touch-action: none;
}
.docx-block-handle:hover, .docx-block-handle:focus-visible { color: #344054; border-color: #98a2b3; outline: none; }
.docx-block-handle[aria-pressed="true"] { color: #175cd3; border-color: #84adff; background: #eff8ff; }
.docx-block-handle.docx-block-dragging { cursor: grabbing; opacity: .35; }
/* The block being carried. Dimming it is the "a drag is happening" signal that survives the
   pointer being anywhere on screen — the drop line only says where, not what. */
.docx-block-drag-source { opacity: .38; transition: opacity 120ms ease-out; }
/* Positioned by transform so tracking the pointer costs no layout. Flipping display none→block
   restarts the fade — one cheap entry animation per appearance, none while it tracks. */
.docx-block-drop-indicator {
  position: fixed; top: 0; left: 0; z-index: 2147482999; display: none; height: 0;
  pointer-events: none; border-top: 2px solid #2e90fa;
  filter: drop-shadow(0 1px 2px rgba(46,144,250,.5));
  animation: docx-block-drop-in 110ms ease-out;
}
.docx-block-drop-indicator::before {
  content: ""; position: absolute; top: -5px; left: -2px; width: 8px; height: 8px;
  border-radius: 50%; background: #2e90fa;
}
@keyframes docx-block-drop-in { from { opacity: 0; } to { opacity: 1; } }
.docx-block-drag-preview {
  max-width: 320px; padding: 6px 10px; border: 1px solid #b2ddff; border-radius: 6px;
  background: #eff8ff; color: #175cd3; box-shadow: 0 6px 16px rgba(16,24,40,.18);
  font: 500 13px/1.35 system-ui, sans-serif; white-space: nowrap; overflow: hidden;
  text-overflow: ellipsis;
}
.docx-block-move-menu {
  position: fixed; z-index: 2147483001; display: none; min-width: 150px; padding: 5px;
  border: 1px solid #d0d5dd; border-radius: 8px; background: #fff;
  box-shadow: 0 8px 24px rgba(16,24,40,.18); font: 13px/1.35 system-ui, sans-serif;
}
.docx-block-move-menu button { display: block; width: 100%; padding: 7px 9px; border: 0; border-radius: 5px; background: transparent; color: #344054; text-align: left; }
.docx-block-move-menu button:hover, .docx-block-move-menu button:focus-visible { background: #f2f4f7; outline: none; }
.docx-block-move-menu button:disabled { color: #98a2b3; background: transparent; }
.docx-block-move-flash { animation: docx-block-move-flash 800ms ease-out; }
@keyframes docx-block-move-flash { from { box-shadow: 0 0 0 3px rgba(46,144,250,.55); } to { box-shadow: none; } }
.docx-block-live-region { position: fixed; width: 1px; height: 1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; }
`;
  (doc.head ?? doc.documentElement).appendChild(style);
}

/**
 * Let the event loop run — input, rendering, and any other task — before the next window
 * of a mount. `scheduler.yield` where the browser has it (it keeps the continuation ahead of
 * other tasks), otherwise a frame followed by a macrotask.
 */
/**
 * Give the page a turn between mount windows: a frame, so a paint and any input handling
 * happen, then a macrotask, so the host's own timers run too. `scheduler.yield()` is not used
 * on purpose — its continuation is scheduled ahead of timer tasks, so a progress spinner or a
 * poll the host drives with `setInterval` would not tick until the whole mount was done.
 */
/** The anchor of the empty paragraphs a chrome render leaves where body content was
 *  (`HtmlConversionOps.ChromeCarrierAnchor`); a windowed mount removes them before it fills
 *  the sections. */
const CHROME_CARRIER_ANCHOR = "chrome-carrier";

function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => {
    const settle = (): void => { setTimeout(resolve, 0); };
    if (typeof requestAnimationFrame === "function") requestAnimationFrame(settle);
    else settle();
  });
}

function trackedChangeWireName(mode: TrackedChangeMode): "accept" | "render_inline" | "strip_deletions" {
  if (mode === TrackedChangeMode.RenderInline) return "render_inline";
  if (mode === TrackedChangeMode.StripDeletions) return "strip_deletions";
  return "accept";
}

// ─── M1: inline HTML → markdown (preserve formatting on edit) ───────────────

interface InlineSeg {
  text: string;
  bold: boolean;
  italic: boolean;
  href: string | null;
}

function fontWeightIsBold(w: string): boolean {
  if (w === "bold" || w === "bolder") return true;
  const n = parseInt(w, 10);
  return !Number.isNaN(n) && n >= 600;
}

function escapeInlineMarkdown(text: string): string {
  // Escape the markdown the projector subset is sensitive to; keep it minimal.
  return text.replace(/([\\`*_[\]])/g, "\\$1");
}

/**
 * Escape block-construct markers at the start of every line of the assembled
 * payload (#571). The projector escapes literal markers on the way out
 * (`\#\# x`) and the parser unescapes `\X` generically, so this is the
 * serializer's half of that contract: a typed leading `## `, `> `, `- `,
 * `+ `, `1. `, or `|` must commit as literal text, not be consumed as a
 * heading/quote/list marker — and a list marker after a soft break must not
 * be split into a second block the session then discards. Line starts only:
 * the parser detects block constructs at line start, and `*` needs no rule
 * here because a literal `*` is already inline-escaped, so an unescaped
 * leading `*` can only be our own emphasis emission.
 */
function escapeLeadingBlockMarkers(md: string): string {
  return md
    .split("\n")
    .map((line) => {
      if (/^[#>|]/.test(line) || /^[-+] /.test(line)) return "\\" + line;
      if (/^\d+\. /.test(line)) return line.replace(".", "\\.");
      return line;
    })
    .join("\n");
}

function collectInlineSegments(node: Node, out: InlineSeg[]): void {
  node.childNodes.forEach((child) => {
    // Skip converter-generated chrome (list markers, note citation markers, note backrefs) —
    // it isn't part of the paragraph's content and must never be committed as text.
    if (isGeneratedChrome(child)) return;
    if (isField(child)) {
      // A field serializes as its cached result (the session's own text for it).
      const text = fieldText(child);
      if (text) out.push({ text, bold: false, italic: false, href: null });
      return;
    }
    const atom = atomText(child);
    if (atom !== null) {
      if (atom) out.push(styledSegment(atom, child.parentElement));
      return;
    }
    if (child.nodeType === 3 /* TEXT_NODE */) {
      const text = child.textContent ?? "";
      if (!text) return;
      out.push(styledSegment(text, child.parentElement));
    } else if (child.nodeType === 1 /* ELEMENT_NODE */) {
      const el = child as HTMLElement;
      collectInlineSegments(el, out);
    }
  });
}

/** `text` with the formatting its parent element renders it in. */
function styledSegment(text: string, parent: HTMLElement | null): InlineSeg {
  let bold = false;
  let italic = false;
  let href: string | null = null;
  if (parent && typeof getComputedStyle === "function") {
    const cs = getComputedStyle(parent);
    bold = fontWeightIsBold(cs.fontWeight);
    italic = cs.fontStyle === "italic" || cs.fontStyle === "oblique";
    const a = parent.closest("a");
    href = a ? a.getAttribute("href") : null;
  }
  return { text, bold, italic, href };
}

/**
 * Trim placeholder padding, retaining native tabs and line breaks in the session's text.
 */
const LEADING_SPACES = /^[^\S\t\n]+/;
const TRAILING_SPACES = /[^\S\t\n]+$/;
function trimSpaces(s: string): string {
  return s.replace(LEADING_SPACES, "").replace(TRAILING_SPACES, "");
}

function segToMarkdown(seg: InlineSeg): string {
  // A <br> segment is a hard line break → the canonical GFM "  \n", which the
  // DocxSession markdown parser turns into a real w:br (Word's intra-paragraph
  // line break) instead of a literal newline in w:t.
  if (seg.text === "\n") return "  \n";
  let md = escapeInlineMarkdown(seg.text).replace(/[ \t]*\n/g, "  \n");
  if (/\S/.test(seg.text)) {
    // Don't wrap pure whitespace — `** **` is not valid emphasis.
    if (seg.bold && seg.italic) md = `***${md}***`;
    else if (seg.bold) md = `**${md}**`;
    else if (seg.italic) md = `*${md}*`;
  }
  if (seg.href) md = `[${md}](${seg.href})`;
  return md;
}

/**
 * Serialize a block's inline content to the projector's markdown subset, preserving
 * bold / italic / links (emphasis detected via computed style). Used so an edit keeps
 * the block's formatting instead of flattening it to plain text. Formatting the markdown
 * subset cannot express (font size/color) is still dropped on an edited block.
 */
export function serializeInlineMarkdown(block: HTMLElement): string {
  const segs: InlineSeg[] = [];
  collectInlineSegments(block, segs);
  // Merge adjacent segments with identical formatting to avoid `**a****b**`.
  const merged: InlineSeg[] = [];
  for (const s of segs) {
    const prev = merged[merged.length - 1];
    if (
      prev &&
      prev.text !== "\n" &&
      s.text !== "\n" &&
      prev.bold === s.bold &&
      prev.italic === s.italic &&
      prev.href === s.href
    ) {
      prev.text += s.text;
    } else {
      merged.push({ ...s });
    }
  }
  // Leading whitespace is never content (a rendered empty paragraph carries a placeholder
  // space, and typing lands after it). Trailing whitespace usually is not either — the
  // browser's bogus trailing <br> serializes as a hard break — EXCEPT the single space a user
  // types before a page-number field goes in: "Page " + PAGE must not commit as "Page". The
  // engine writes it with xml:space="preserve"; `trimBounds` counts it the same way.
  const md = trimSpaces(merged.map(segToMarkdown).join(""));
  const content = md.startsWith("\n") ? "  " + md : md;
  return escapeLeadingBlockMarkers(content + (keepsTrailingSpace(block) ? " " : ""));
}

/**
 * Whether a block's content text ends in a space the commit should keep: a plain or no-break
 * space (contenteditable turns a typed trailing space into U+00A0) after some non-whitespace
 * text. A whitespace-only block (the placeholder an empty paragraph renders as) keeps nothing.
 */
function keepsTrailingSpace(block: HTMLElement): boolean {
  const text = blockContentText(block);
  if (trimSpaces(text).length === 0) return false;
  // An ordinary trailing space can only have come from the document (the converter renders
  // it from a preserved w:t), so it is content. A trailing NBSP is ambiguous: the browser
  // writes one for a typed space, but the placeholder an empty paragraph renders as is one
  // too, and typing at the start of such a paragraph leaves it dangling at the end. Only the
  // typed one counts, which `wireBlock` records from the input event.
  if (/ $/.test(text)) return true;
  return /\u00a0$/.test(text) && block.dataset.typedTrailingSpace === "1";
}

/** Record whether the last input left a typed space at the end of the block \u2014 see
 *  {@link keepsTrailingSpace}. */
function noteTypedSpace(block: HTMLElement, event: Event): void {
  const input = event as InputEvent;
  const typedSpace =
    input.inputType === "insertText" && typeof input.data === "string" && /[ \u00a0]$/.test(input.data);
  // At the end, or before nothing but whitespace: the placeholder an empty paragraph renders as
  // stays after the caret when the caret was put at the paragraph's start, and is not content.
  const caret = typedSpace ? caretOffsetIn(block) : null;
  const atEnd = caret !== null && trimSpaces(blockContentText(block).slice(caret)).length === 0;
  if (atEnd) block.dataset.typedTrailingSpace = "1";
  else delete block.dataset.typedTrailingSpace;
}

// ─── M2: structural editing (split / merge) ─────────────────────────────────

interface AnchorRef {
  id: string;
  kind: string;
  scope: string;
  unid: string;
}

/**
 * An `EditResult` as the editor reads it straight off the bridge. It is derived from the public
 * type the typed session wrapper (`session.ts`) returns, so a change to the wire shape surfaces
 * in both at compile time; the lists are optional because a malformed reply may omit them.
 */
type EditResultLite = Pick<EditResult, "success">
  & Partial<Pick<EditResult, "created" | "removed" | "modified" | "error" | "caret" | "imageId">>;

/**
 * How text flows around a picture, as Word's Wrap Text names it: in line with the text, or
 * floating with square, tight, through or top-and-bottom wrapping, or behind or in front of the
 * text (no wrapping).
 */
export type PictureWrap = "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront";

/** The selected picture, as the Picture Format tab and dialog show it. */
export interface PictureInfo {
  /** The engine's image id. */
  id: string;
  /** How text wraps it; null when the document says something the editor cannot change. */
  wrap: PictureWrap | null;
  /** Its size on the page, in points (null when the document gives none). */
  widthPoints: number | null;
  heightPoints: number | null;
  /** Its picture's own size at 96 dpi, in points (Reset), when known. */
  originalWidthPoints: number | null;
  originalHeightPoints: number | null;
  altText: string;
  title: string;
  /** Where a floating picture sits and how text wraps round it; null for an inline one. */
  layout: FloatingImageLayout | null;
  /** What can be changed. */
  can: { size: boolean; wrap: boolean; layout: boolean; altText: boolean; replace: boolean; remove: boolean };
}

/** What a picture can be selected as: a rendered picture (an inline SVG for a vector one). */
const PICTURE_SELECTOR = "img, [data-docx-vector-frame]";

/** Word's clearance at the sides of a picture newly made to float (0.125"). */
const FLOAT_SIDE_DISTANCE_EMU = 114300;

function wrapOf(occurrence: ImageOccurrence): PictureWrap | null {
  if (occurrence.placement === "inline") return "inline";
  const layout = occurrence.floatingLayout;
  if (occurrence.placement !== "floating" || !layout) return null;
  switch (layout.wrapMode) {
    case "none": return layout.behindDocument ? "behind" : "inFront";
    case "square": return "square";
    case "tight": return "tight";
    case "through": return "through";
    case "top_and_bottom": return "topAndBottom";
    default: return null;
  }
}

function pictureInfoOf(occurrence: ImageOccurrence): PictureInfo {
  const can = (name: string): boolean =>
    !!occurrence.operations?.find((op) => op.operation === name)?.canMutate;
  const floating = occurrence.placement === "floating";
  const layoutWritable = floating && occurrence.floatingLayoutSupported && can("set_floating_layout");
  const px = (n: number | undefined): number | null => (n && n > 0 ? n * 0.75 : null);
  return {
    id: occurrence.id,
    wrap: wrapOf(occurrence),
    widthPoints: occurrence.renderedWidthPoints ?? null,
    heightPoints: occurrence.renderedHeightPoints ?? null,
    originalWidthPoints: px(occurrence.intrinsicWidthPixels),
    originalHeightPoints: px(occurrence.intrinsicHeightPixels),
    altText: occurrence.altText ?? "",
    title: occurrence.title ?? "",
    layout: floating ? occurrence.floatingLayout ?? null : null,
    can: {
      size: can("set_dimensions"),
      wrap: can("set_placement") || layoutWritable,
      layout: layoutWritable,
      altText: can("set_metadata"),
      replace: can("replace"),
      remove: can("remove"),
    },
  };
}

/** A layout as the engine takes it back: the read-only raw tokens left out. */
function writableLayout(layout: FloatingImageLayout): FloatingImageLayout {
  const copy: Record<string, unknown> = { ...layout };
  for (const key of Object.keys(copy)) if (key.startsWith("raw")) delete copy[key];
  return copy as FloatingImageLayout;
}

/** What {@link DocxEditorOptions.onEditFailed} receives when the engine rejects an edit. */
export interface EditFailure {
  /** The engine's error code, when it sent one. */
  code?: EditErrorCode;
  /** The engine's error message. */
  message: string;
  /** The anchor the engine named in its error, else the block being edited, when there is one. */
  anchorId: string | null;
}

interface DomSelectionPoint {
  node: Node;
  offset: number;
}

interface CrossBlockSelectionBookmark {
  startUnid: string;
  startOffset: number;
  endUnid: string;
  endOffset: number;
}

interface CrossBlockDragSelection {
  anchor: DomSelectionPoint;
  origin: HTMLElement;
  crossedBlockBoundary: boolean;
  focus: DomSelectionPoint | null;
  frame: number | null;
}

/** True if `block` renders as a list item (has a generated marker as its first child). */
function isListBlock(block: HTMLElement): boolean {
  return !!block.querySelector(":scope > [data-list-marker]");
}

/**
 * The converter's tab box must stay text-empty because any child changes its inline baseline.
 * Editable list items still need a textContent separator between generated marker chrome and an
 * empty body so typing produces "2. BB", not "2.BB". Keep that editor-only separator hidden and
 * generated, outside the converter-owned marker wrapper.
 */
function ensureListMarkerSeparator(block: HTMLElement): void {
  const marker = block.querySelector<HTMLElement>(":scope > [data-list-marker]");
  if (!marker) return;
  const next = marker.nextElementSibling as HTMLElement | null;
  if (next?.hasAttribute("data-editor-list-separator")) return;

  const separator = block.ownerDocument.createElement("span");
  separator.setAttribute("data-editor-list-separator", "");
  separator.setAttribute("data-list-marker", "true");
  separator.setAttribute("aria-hidden", "true");
  separator.style.display = "none";
  separator.textContent = " ";
  marker.after(separator);
}

/**
 * Inline chrome the CONVERTER generates that is not part of a paragraph's run text: list
 * number/bullet markers, footnote/endnote citation markers
 * (`<a class="footnote-ref"><sup>1</sup></a>`), the paginated notes' own numbers, and the note
 * backrefs (`↩`).
 *
 * The session's run text contains none of it. A citation is a zero-width
 * `w:footnoteReference` — the displayed number is computed by the renderer from document order —
 * so every character of chrome the editor fails to exclude shifts its content-offset space away
 * from the session's. Each omission has its own failure mode: excluded from offsets but not from
 * serialization and the display number gets COMMITTED as literal text (destroying the citation
 * run); left editable and the user can delete a marker outright, orphaning the note.
 */
/**
 * The paginated view's note numbers open their note's first paragraph, as Word's reference mark
 * run does (issue #1130); the session holds that mark as a zero-width w:footnoteRef/w:endnoteRef.
 */
const NOTE_NUMBER_SELECTOR = ".footnote-number, .endnote-number";

/**
 * A tracked deletion rendered inline (`<del>`) and the pilcrow shown for a deleted paragraph mark
 * are not run text either: the session reads only live `w:t`, so counting them put every offset to
 * their right past the session's (issue #1278). A move source (`<del class="rev-move-from">`) is
 * removed text too. One directly in the paragraph makes the block read-only (see `wireBlock`); the
 * session leaves one nested in a hyperlink or other inline container out of its text whether its
 * runs hold `w:t`, as Word writes it, or `w:delText` (issue #1282), so the editor does as well.
 */
const RENDERED_TEXT_DELETION_SELECTOR = 'del[class*="-del"], del[class*="move-from"]';
/** The pilcrow the renderer draws at the end of a paragraph whose mark is a pending deletion. */
const DELETED_PARAGRAPH_MARK_SELECTOR = 'span[class*="para-mark-del"]';
const RENDERED_DELETION_SELECTOR = RENDERED_TEXT_DELETION_SELECTOR + ", " + DELETED_PARAGRAPH_MARK_SELECTOR;

const GENERATED_CHROME_SELECTOR =
  '[data-list-marker], a.footnote-ref, a.endnote-ref, a[class$="-backref"], a.comment-marker, '
  + NOTE_NUMBER_SELECTOR + ", " + RENDERED_DELETION_SELECTOR;

function isGeneratedChrome(node: Node | null | undefined): boolean {
  return node?.nodeType === 1 && !!(node as Element).matches?.(GENERATED_CHROME_SELECTOR);
}

/**
 * A paginated note's number is chrome inside an editable paragraph, and the browser deletes a
 * non-editable island as a unit when a caret deletion reaches it: Backspace just after it, Delete
 * just before it, a word deletion that spans it. Nothing in the session re-creates the number (it
 * is the renderer's, not run text), so refuse those deletions, as Word refuses to backspace over a
 * reference mark. A deletion that replaces a selection is allowed; the repaint that commits it
 * restores the number (see `DocxEditor.replaceNode`). A rendered deletion is not removed either
 * (the session already holds it as deleted, so that would only drop it from view); as in Word, a
 * character Backspace or Delete steps over it and removes the live character beyond it instead,
 * where that is safe (see `deleteLiveCharacterBeyond`). A word or line deletion that reaches one
 * is refused.
 */
function guardChromeDeletion(block: HTMLElement, ev: InputEvent): void {
  if (!ev.inputType.startsWith("delete")) return;
  const islands = block.querySelectorAll(NOTE_NUMBER_SELECTOR + ", " + RENDERED_DELETION_SELECTOR);
  if (islands.length === 0) return;
  const selection = block.ownerDocument.defaultView?.getSelection();
  if (selection && !selection.isCollapsed) return;
  const targets = typeof ev.getTargetRanges === "function" ? ev.getTargetRanges() : [];
  for (const target of targets) {
    const range = block.ownerDocument.createRange();
    range.setStart(target.startContainer, target.startOffset);
    range.setEnd(target.endContainer, target.endOffset);
    const reached = Array.from(islands).filter((island) => range.intersectsNode(island));
    if (reached.length > 0) {
      ev.preventDefault();
      if (!reached.some((island) => island.matches(NOTE_NUMBER_SELECTOR))) {
        deleteLiveCharacterBeyond(block, ev.inputType);
      }
      return;
    }
  }
}

/**
 * Content that is not plain text and not chrome either (a tab is one character of the content
 * offsets, the rest take none): stepping over a deletion must not also cross one of these to
 * reach a character on its far side.
 */
const STEP_BLOCKER_SELECTOR =
  'br, img, svg, picture, canvas, video, object, [data-docx-tab], [data-field], [contenteditable="false"]';

/**
 * Remove the live character before (Backspace) or after (Delete) the collapsed caret, measured in
 * content offsets, which rendered deletions do not occupy, so the step skips over them. Only a
 * single-character deletion is stepped, and only across chrome: when a line break, tab, image or
 * field also lies between the caret and that character, or there is no character beyond the
 * deletion, nothing is removed. The caret stays where the removed character was.
 */
function deleteLiveCharacterBeyond(block: HTMLElement, inputType: string): void {
  const backward = inputType === "deleteContentBackward";
  if (!backward && inputType !== "deleteContentForward") return;
  const caret = caretOffsetIn(block);
  if (caret === null) return;
  const end = backward ? caret : caret + 1;
  if (end < 1 || end > blockContentText(block).length) return;
  // contentPositionIn resolves an offset to the earliest text node reaching it, so the
  // character just before `end` is the one before this position in this node.
  const { node, offset } = contentPositionIn(block, end);
  if (node.nodeType !== 3 /* TEXT_NODE */ || offset < 1) return;
  const text = node.textContent ?? "";
  let from = offset - 1;
  let to = offset;
  if (/[\uDC00-\uDFFF]/.test(text[from]) && /[\uD800-\uDBFF]/.test(text[from - 1] ?? "")) from--;
  else if (/[\uD800-\uDBFF]/.test(text[from]) && /[\uDC00-\uDFFF]/.test(text[to] ?? "")) to++;
  const selection = block.ownerDocument.defaultView?.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const at = selection.getRangeAt(0);
  const gap = block.ownerDocument.createRange();
  if (backward) {
    gap.setStart(node, to);
    gap.setEnd(at.startContainer, at.startOffset);
  } else {
    gap.setStart(at.startContainer, at.startOffset);
    gap.setEnd(node, from);
  }
  const between = gap.cloneContents();
  between.querySelectorAll(GENERATED_CHROME_SELECTOR).forEach((chrome) => chrome.remove());
  if (between.querySelector(STEP_BLOCKER_SELECTOR)) return;
  node.textContent = text.slice(0, from) + text.slice(to);
  selection.collapse(node, from);
  block.dispatchEvent(new InputEvent("input", { inputType, bubbles: true, composed: true }));
}

/** Let the browser choose visual direction (including bidi), then cross a tab's empty layout
 *  wrapper if that step did not advance in the document's content text. */
function moveAcrossTab(block: HTMLElement, ev: KeyboardEvent): void {
  if (ev.defaultPrevented || ev.shiftKey || ev.ctrlKey || ev.metaKey || ev.altKey || ev.isComposing
    || (ev.key !== "ArrowLeft" && ev.key !== "ArrowRight")) return;
  const selection = block.ownerDocument.defaultView?.getSelection();
  const at = caretOffsetIn(block);
  if (!selection?.isCollapsed || !selection.rangeCount || at === null) return;
  const text = blockContentText(block);
  if (text[at] !== "\t" && text[at - 1] !== "\t") return;
  const before = selection.getRangeAt(0).cloneRange();
  selection.modify("move", ev.key === "ArrowRight" ? "right" : "left", "character");
  ev.preventDefault();
  const after = caretOffsetIn(block);
  if (after === null) return; // Native movement may leave this editing host.
  const order = before.compareBoundaryPoints(Range.START_TO_START, selection.getRangeAt(0));
  const source = before.startContainer.nodeType === 1 ? before.startContainer as Element : before.startContainer.parentElement!;
  const rtl = block.ownerDocument.defaultView!.getComputedStyle(source).direction === "rtl";
  const step = order < 0 ? 1 : order > 0 ? -1 : (ev.key === "ArrowRight") !== rtl ? 1 : -1;
  const crossed = after === at && text[step > 0 ? at : at - 1] === "\t";
  collapseAtContentOffset(block, selection, crossed ? at + step : after);
}

/** Collapse `selection` at content offset `offset` of `block` so that typing lands there, also
 *  just after a tab, whose empty box Chromium would otherwise type before. */
function collapseAtContentOffset(block: HTMLElement, selection: Selection, offset: number): void {
  let { node, offset: at } = contentPositionIn(block, offset);
  // A DOM point after the empty tab still types before it in Chromium while inside the flex
  // wrapper. Lift an end boundary out of its trailing wrappers into the surrounding content.
  while (node !== block && node.nodeType === 1 && at === node.childNodes.length && node.parentNode) {
    at = Array.prototype.indexOf.call(node.parentNode.childNodes, node) + 1;
    node = node.parentNode;
  }
  // A terminal empty box has no native caret stop after it. The same unstamped BR Chromium
  // uses for an empty line supplies one; atomText excludes this placeholder from saved text.
  if (node === block && at === block.childNodes.length) block.append(block.ownerDocument.createElement("br"));
  selection.collapse(node, at);
}

/**
 * Backspace just after, or Delete just before, a rendered tab/break removes that atom, and so does a word
 * deletion (Ctrl+Backspace, Ctrl+Delete) that starts there: a tab is the whitespace between words.
 * The browser does not: the tab is an empty box inside the converter's tab-layout wrapper, and
 * Chromium moves the text after it into the wrapper instead, or, for Ctrl+Delete, deletes the word
 * BEFORE the caret, so the commit would move the tab or remove the wrong text. The tab is one
 * character of the session's text (issue #1387), so removing its box is exactly the edit the commit
 * then sends. Breaks need the same treatment: the converter's bidi mark after a BR can otherwise
 * consume Backspace while leaving the break in the document.
 */
function deleteAdjacentWhitespaceAtom(block: HTMLElement, ev: InputEvent): void {
  const backward = ev.inputType === "deleteContentBackward" || ev.inputType === "deleteWordBackward";
  const forward = ev.inputType === "deleteContentForward" || ev.inputType === "deleteWordForward";
  if (ev.defaultPrevented || (!backward && !forward)) return;
  if (!block.ownerDocument.defaultView?.getSelection()?.isCollapsed) return;
  const caret = caretOffsetIn(block);
  if (caret === null) return;
  const at = backward ? caret - 1 : caret;
  const text = blockContentText(block)[at];
  if (text !== "\t" && text !== "\n") return;
  const atom = Array.from(block.querySelectorAll<HTMLElement>(
    "[data-docx-tab], br, [data-page-break], [data-column-break]",
  )).find((box) => !isInMarker(box) && box.parentNode && atomText(box) === text
    && contentOffsetOf(block, box.parentNode, Array.prototype.indexOf.call(box.parentNode.childNodes, box)) === at);
  if (!atom) return;
  ev.preventDefault();
  atom.remove();
  placeCaretAtOffset(block, at);
  block.dispatchEvent(new InputEvent("input", { inputType: ev.inputType, bubbles: true, composed: true }));
}

/**
 * A rendered field — PAGE / NUMPAGES in a running story, stamped `[data-field]` by the editor
 * render profile. Fields are ATOMIC for editing (`contenteditable=false`, the caret lands on
 * either side), and their content length is the field's CACHED result — the text the session
 * holds in the result run — even after the page view substituted the per-page number into the
 * DOM. Counting the substituted text instead would shift every offset after the field and make
 * the commit diff rewrite the field runs.
 */
function fieldOf(node: Node | null): HTMLElement | null {
  const el = node && node.nodeType === 1 ? (node as Element) : node?.parentElement ?? null;
  return (el?.closest?.("[data-field]") as HTMLElement | null) ?? null;
}

function isField(node: Node): node is HTMLElement {
  return node.nodeType === 1 && (node as HTMLElement).hasAttribute("data-field");
}

/** A rendered tab: an empty box in the DOM, but one character, `"\t"`, of the session's text
 *  (issue #1387). Like a field it is atomic to the offset helpers, which place the caret on either
 *  side of it, never inside. Arrow movement also skips empty tab-layout wrappers. */
function isTab(node: Node): node is HTMLElement {
  return node.nodeType === 1 && (node as HTMLElement).hasAttribute("data-docx-tab");
}

/** Chromium's final, unstamped BR holds the caret on an empty line; it is not another break. */
function isPlaceholderBreak(node: HTMLElement): boolean {
  if (node.tagName !== "BR" || node.hasAttribute("data-docx-break")) return false;
  let current: Node = node;
  while (current.parentElement) {
    if (current.nextSibling) return false;
    if (current.parentElement.getAttribute("contenteditable") === "true") return true;
    current = current.parentElement;
  }
  return false;
}

/** Native text of an atomic node, else null. Every offset and commit helper reads this rule. */
function atomText(node: Node): string | null {
  if (isField(node)) return fieldText(node);
  if (isTab(node)) return "\t";
  if (node.nodeType === 1) {
    const el = node as HTMLElement;
    if (el.tagName === "BR") return isPlaceholderBreak(el) ? "" : "\n";
    if (el.hasAttribute("data-page-break") || el.hasAttribute("data-column-break")) return "\n";
    // A symbol-font glyph the converter mapped to Unicode (#1707) stands for the characters the run
    // stores, which the session counts and which an edit beside it must keep.
    const symbol = el.getAttribute("data-docx-symbol-text");
    if (symbol !== null) return symbol;
  }
  return null;
}

/** The text a field counts for: the cached result the session holds, else what it shows. */
function fieldText(field: HTMLElement): string {
  return field.dataset.fieldCached ?? field.textContent ?? "";
}

/** Freeze a field's rendered result as its cached text (once) and keep the caret out of it. */
function adoptFields(block: HTMLElement): void {
  for (const field of Array.from(block.querySelectorAll<HTMLElement>("[data-field]"))) {
    if (field.dataset.fieldCached === undefined) field.dataset.fieldCached = field.textContent ?? "";
    field.setAttribute("contenteditable", "false");
  }
}

/** True if `node` is, or is inside, generated chrome (not editable content). */
function isInMarker(node: Node | null): boolean {
  let el: Element | null = node && node.nodeType === 1 ? (node as Element) : node?.parentElement ?? null;
  while (el) {
    if (isGeneratedChrome(el)) return true;
    el = el.parentElement;
  }
  return false;
}

/**
 * Unicode bidi formatting marks the HTML converter injects to preserve visual order: LRM/RLM/ALM,
 * the embedding/override controls, and the isolates (see WmlToHtmlConverter — a paragraph/run gets
 * a leading U+200E or U+200F). They are presentation-only — NOT part of the paragraph's run text the
 * session holds — so the editor must exclude them from its content-offset space, the same way it
 * excludes generated list markers. Otherwise every caret offset is shifted by the leading mark and a
 * caret at end-of-line overshoots the session's text length, so SplitParagraph/ApplyFormat reject the
 * offset (symptom: Enter at the end of a Google-Docs-exported paragraph is silently dropped).
 */
// LRM, RLM, ALM; the embedding/override controls (LRE RLE PDF LRO RLO); the isolates (LRI RLI FSI PDI).
const BIDI_MARK_CLASS = "\u200E\u200F\u061C\u202A-\u202E\u2066-\u2069";
const BIDI_MARKS_RE_G = new RegExp(`[${BIDI_MARK_CLASS}]`, "g");
const BIDI_MARK_RE = new RegExp(`[${BIDI_MARK_CLASS}]`);
function stripBidi(s: string): string {
  return s.replace(BIDI_MARKS_RE_G, "");
}

/** Raw string index in `s` for content offset `n` (content = chars excluding bidi marks). */
function domOffsetForContentOffset(s: string, n: number): number {
  let content = 0;
  for (let i = 0; i < s.length; i++) {
    if (content >= n) return i;
    if (!BIDI_MARK_RE.test(s[i])) content++;
  }
  return s.length;
}

/**
 * Content-text offset of (container, offset) within `block`, EXCLUDING generated list-marker
 * text and injected bidi marks. This is the offset DocxSession ops expect (the paragraph's run
 * text, not the rendered number/bullet or bidi marks the converter injects).
 */
function contentOffsetOf(block: HTMLElement, container: Node, offset: number): number {
  let count = 0;
  let done = false;
  const walk = (node: Node): void => {
    if (done) return;
    // A field or tab inside a rendered deletion is deleted text: the walk below skips it like any other.
    const atom = atomText(node);
    if (atom !== null && !isInMarker(node)) {
      // Atomic: a point inside a field or tab resolves to its end; it counts its text either way.
      const inside = node === container || node.contains(container);
      if (!inside || offset > 0 || node !== container) count += stripBidi(atom).length;
      if (inside) done = true;
      return;
    }
    if (node.nodeType === 3 /* TEXT_NODE */) {
      if (node === container) {
        if (!isInMarker(node)) count += stripBidi((node.textContent ?? "").slice(0, offset)).length;
        done = true; return;
      }
      if (!isInMarker(node)) count += stripBidi(node.textContent ?? "").length;
    } else {
      if (node === container) {
        // Element container: `offset` is a child index — count content up to that child.
        const kids = Array.from(node.childNodes);
        for (let i = 0; i < offset && i < kids.length; i++) walk(kids[i]);
        done = true;
        return;
      }
      node.childNodes.forEach(walk);
    }
  };
  walk(block);
  return count;
}

/** Content-text offset of the collapsed caret within `block` (excludes markers), or null. */
function caretOffsetIn(block: HTMLElement): number | null {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (!block.contains(range.startContainer)) return null;
  return contentOffsetOf(block, range.startContainer, range.startOffset);
}

/**
 * Which side of a rendered tracked deletion the text the user typed sits on, for an insertion of
 * `length` characters at content offset `start` of the block's current text. Deleted text takes no
 * room in the content offsets, so the offset alone names both sides of it, and the commit runs on
 * blur, when the caret is gone; where the typed characters landed in the DOM still says (issue
 * #1290). Undefined when no deletion touches the insertion, or one touches each side.
 */
function deletedTextSideOf(block: HTMLElement, start: number, length: number): "before" | "after" | undefined {
  let before = false;
  let after = false;
  // Only whitespace may separate the deletion from the typed text: an empty paragraph's
  // placeholder space can sit between them, on either side.
  const text = blockContentText(block);
  const blank = (from: number, to: number) => from <= to && trimSpaces(text.slice(from, to)).length === 0;
  for (const deletion of Array.from(block.querySelectorAll(RENDERED_TEXT_DELETION_SELECTOR))) {
    const parent = deletion.parentNode;
    if (!parent) continue;
    const at = contentOffsetOf(block, parent, Array.prototype.indexOf.call(parent.childNodes, deletion));
    if (blank(at, start)) after = true;
    else if (blank(start + length, at)) before = true;
  }
  return before === after ? undefined : before ? "before" : "after";
}

/** True when nothing in `block` would serialize as inline markdown: no bold, italic, link or
 *  line break, so its text can go in as a plain span insertion without losing anything. */
function holdsOnlyPlainText(block: HTMLElement): boolean {
  const segs: InlineSeg[] = [];
  collectInlineSegments(block, segs);
  return segs.every((seg) => !seg.bold && !seg.italic && seg.href === null && seg.text !== "\n");
}

/** Visible content text of `block`, excluding generated list-marker text (the same content
 *  caretOffsetIn/contentOffsetOf count). */
function blockContentText(block: HTMLElement): string {
  let out = "";
  const walk = (node: Node): void => {
    const atom = atomText(node);
    if (atom !== null && !isInMarker(node)) {
      out += stripBidi(atom);
    } else if (node.nodeType === 3 /* TEXT_NODE */) {
      if (!isInMarker(node)) out += stripBidi(node.textContent ?? "");
    } else {
      node.childNodes.forEach(walk);
    }
  };
  walk(block);
  return out;
}

/**
 * `blockContentText` as Find searches it: the same characters at the same offsets, but a manual
 * line break is Word's "\v", a page break "\f" and a column break "\x0E", so ^l, ^m and ^n can
 * tell them apart (see find-query.ts).
 */
function blockStreamText(block: HTMLElement): string {
  let out = "";
  const walk = (node: Node): void => {
    const atom = atomText(node);
    if (atom !== null && !isInMarker(node)) {
      if (atom === "\n" && node.nodeType === 1) {
        const el = node as HTMLElement;
        out += el.hasAttribute("data-page-break") ? "\f" : el.hasAttribute("data-column-break") ? "\x0E" : "\v";
      } else {
        out += stripBidi(atom);
      }
    } else if (node.nodeType === 3 /* TEXT_NODE */) {
      if (!isInMarker(node)) out += stripBidi(node.textContent ?? "");
    } else {
      node.childNodes.forEach(walk);
    }
  };
  walk(block);
  return out;
}

/** Replacement text in stream characters as a run takes it: a line break is the engine's "\n";
 *  page and column breaks cannot be typed into a run, so they go. */
function streamToRunText(text: string): string {
  return text.replace(/\v/g, "\n").replace(/[\f\x0E]/g, "");
}

/** A block's text as a reader sees it: no-break spaces as spaces, an empty paragraph's
 *  placeholder as nothing. */
function plainBlockText(block: HTMLElement): string {
  const text = blockContentText(block).replace(/\u00a0/g, " ");
  return trimSpaces(text).length === 0 ? "" : text;
}

/**
 * Map a DOM caret offset (from caretOffsetIn) into the run-text offset the session holds after a
 * commit. commitBlock/syncBlock commit `serializeInlineMarkdown(el)`, which `.trim()`s leading and
 * trailing whitespace, so the session's paragraph text is shorter than the DOM text whenever the
 * block has edge whitespace — e.g. a blank document renders its empty paragraph with a placeholder
 * space, and typing lands after it. Without this adjustment the caret offset overshoots the
 * committed length, SplitParagraph returns OffsetOutOfRange, and splitAtCaret silently drops the
 * Enter (no new paragraph). Subtracting the leading whitespace before the caret and clamping to the
 * trimmed length keeps the split offset consistent with what was committed.
 */
function trimmedSplitOffset(block: HTMLElement, domOffset: number): number {
  const { leading, trimmedLen } = trimBounds(block);
  return Math.max(0, Math.min(domOffset - Math.min(domOffset, leading), trimmedLen));
}

/** How far `block`'s DOM content text is offset from, and longer than, its committed form. */
function trimBounds(block: HTMLElement): { leading: number; trimmedLen: number } {
  const content = blockContentText(block);
  return {
    leading: content.length - content.replace(LEADING_SPACES, "").length,
    // Mirrors serializeInlineMarkdown: leading and trailing whitespace dropped, plus the one
    // trailing space a commit keeps.
    trimmedLen: trimSpaces(content).length + (keepsTrailingSpace(block) ? 1 : 0),
  };
}

/**
 * Map a DOM content SPAN into the run-text space the session holds after a commit — the span
 * analogue of {@link trimmedSplitOffset}, and for the same reason.
 *
 * A block rendered with edge whitespace produces a span longer than the text the commit stores:
 * an empty header/footer story renders as a lone NBSP placeholder, so typing into it leaves a
 * trailing NBSP that `serializeInlineMarkdown(...).trim()` removes (JS `trim()` treats U+00A0 as
 * whitespace). "Select all, then Bold" then asks `ApplyFormat` for [0, len+1) — one past the
 * committed end — and the op is REJECTED with OffsetOutOfRange, so the format silently does
 * nothing. The demo's format buttons preventDefault on mousedown to keep the selection alive,
 * which is exactly the path that computes the span before `syncBlock` commits, so this is the
 * ordinary case rather than an edge one.
 */
function trimmedSpan(
  block: HTMLElement,
  span: { start: number; length: number },
): { start: number; length: number } {
  const { leading, trimmedLen } = trimBounds(block);
  const start = Math.max(0, Math.min(span.start - Math.min(span.start, leading), trimmedLen));
  const end = Math.max(start, Math.min(span.start + span.length - leading, trimmedLen));
  return { start, length: end - start };
}

/**
 * DOM (node, offset) for content offset `offset` within `el` — the same content-offset
 * space as contentOffsetOf (marker text and injected bidi marks excluded). Clamps past-end
 * offsets to the end of the last text node (or the element itself when it has none), so a
 * caller can always build a Range from the result.
 */
function contentPositionIn(el: HTMLElement, offset: number): { node: Node; offset: number } {
  let remaining = offset;
  let result: { node: Node; offset: number } | null = null;
  let lastText: Node | null = null;
  const walk = (node: Node): void => {
    if (result) return;
    const atom = atomText(node);
    if (atom !== null && !isInMarker(node)) {
      // Never land inside a field or tab: an offset at its start sits before it, anything up to
      // its length sits after it.
      const len = stripBidi(atom).length;
      const parent = node.parentNode;
      if (parent && remaining <= len) {
        const index = Array.prototype.indexOf.call(parent.childNodes, node);
        result = { node: parent, offset: index + (remaining > 0 ? 1 : 0) };
        return;
      }
      remaining -= len;
      return;
    }
    if (node.nodeType === 3 /* TEXT_NODE */) {
      if (isInMarker(node)) return;
      const raw = node.textContent ?? "";
      const len = stripBidi(raw).length;
      lastText = node;
      if (remaining <= len) {
        result = { node, offset: domOffsetForContentOffset(raw, remaining) };
        return;
      }
      remaining -= len;
    } else {
      node.childNodes.forEach(walk);
    }
  };
  walk(el);
  if (result) return result;
  if (lastText) return { node: lastText, offset: ((lastText as Node).textContent ?? "").length };
  return { node: el, offset: el.childNodes.length };
}

/**
 * Resolve a viewport coordinate to a DOM caret. Firefox exposes the standards-track
 * caretPositionFromPoint; Chromium/WebKit expose caretRangeFromPoint. Keeping the compatibility
 * shim here lets the editor bridge native mouse selection across its intentionally-independent
 * per-paragraph editing hosts without changing their commit boundaries.
 */
function caretPointFromClient(
  doc: Document,
  clientX: number,
  clientY: number,
): DomSelectionPoint | null {
  const caretDoc = doc as Document & {
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    caretRangeFromPoint?: (x: number, y: number) => Range | null;
  };
  const position = caretDoc.caretPositionFromPoint?.(clientX, clientY);
  if (position) return { node: position.offsetNode, offset: position.offset };
  const range = caretDoc.caretRangeFromPoint?.(clientX, clientY);
  return range ? { node: range.startContainer, offset: range.startOffset } : null;
}

/**
 * Set a possibly-backward pair of points as one normalized DOM Range. Chromium's
 * Selection.setBaseAndExtent() exposes cross-contenteditable endpoints but clips toString() and
 * the effective contents to the originating host; addRange() carries the complete cross-block
 * contents. Direction is immaterial to editor commands, which consume normalized start/end.
 */
function setSelectionBetween(anchor: DomSelectionPoint, focus: DomSelectionPoint): boolean {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || !anchor.node.isConnected || !focus.node.isConnected) return false;
  try {
    const probe = document.createRange();
    probe.setStart(anchor.node, anchor.offset);
    probe.collapse(true);
    const focusIsBefore = probe.comparePoint(focus.node, focus.offset) < 0;
    const range = document.createRange();
    const start = focusIsBefore ? focus : anchor;
    const end = focusIsBefore ? anchor : focus;
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    sel.removeAllRanges();
    sel.addRange(range);
    return true;
  } catch {
    return false;
  }
}

/** Word's initials for a user name: the first letter of each word ("Ada Lovelace" → "AL"). */
function nameInitials(name: string): string {
  return name.trim().split(/\s+/).filter(Boolean).map((word) => Array.from(word)[0]!.toUpperCase()).join("");
}

/** Now, in the form Word writes `w:date`: UTC to the second ("2026-11-13T20:10:45Z"). */
function commentTimestamp(): string {
  return new Date().toISOString().replace(/\.\d+Z$/, "Z");
}

/** The viewport box of a caret at `p` (an empty element's own box when it draws no caret). */
function caretRect(p: DomSelectionPoint): DOMRect | null {
  try {
    const range = document.createRange();
    range.setStart(p.node, p.offset);
    range.collapse(true);
    const rect = range.getClientRects()[0];
    if (rect) return rect;
    const el = p.node.nodeType === 1 ? (p.node as Element) : p.node.parentElement;
    return el?.getBoundingClientRect() ?? null;
  } catch {
    return null;
  }
}

/** True when carets at `a` and `b` sit on the same line (their boxes overlap vertically). */
function sameLine(a: DomSelectionPoint, b: DomSelectionPoint): boolean {
  const ra = caretRect(a), rb = caretRect(b);
  if (!ra || !rb) return false;
  return Math.min(ra.bottom, rb.bottom) - Math.max(ra.top, rb.top) > Math.min(ra.height, rb.height) / 2;
}

/** Place the caret at content offset `offset` within `el`, skipping marker text. */
function placeCaretAtOffset(el: HTMLElement, offset: number): void {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  // A remount can detach the block between a caller's lookup and this call; addRange on a
  // detached node throws "the given range isn't in document".
  if (!sel || !el.isConnected) return;
  el.focus();
  const range = document.createRange();
  // contentPositionIn clamps a past-end offset to the end of the last text node, and never
  // lands inside marker chrome or a field — the same rules every other offset helper uses.
  const pos = contentPositionIn(el, offset);
  try {
    range.setStart(pos.node, pos.offset);
  } catch {
    range.selectNodeContents(el);
    range.collapse(false);
  }
  range.collapse(true);
  sel.removeAllRanges();
  sel.addRange(range);
}

// ─── M5: formatting controls ────────────────────────────────────────────────

/** The nine levels of legal numbering (1., 1.1, 1.1.1 …): decimal at every level, each level's
 *  number the levels above it and its own, joined by periods; `periods` says where a trailing
 *  period goes (`"top"`: the top level only, `"all"`: every level, `"none"`: none). The
 *  positions are Word's own legal list's: the number room grows with the number (1. to 1.1.1.1). */
export function legalLevels(periods: "top" | "all" | "none" = "top"): ListLevelSpec[] {
  const textAt = [360, 792, 1224, 1728, 2232, 2736, 3240, 3744, 4320];
  const hanging = [360, 432, 504, 648, 792, 936, 1080, 1224, 1440];
  return Array.from({ length: 9 }, (_, level) => {
    const number = Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join(".");
    const trail = periods === "all" || (periods === "top" && level === 0) ? "." : "";
    return {
      level, format: "decimal", text: number + trail, legal: level > 0,
      numberPositionTwips: textAt[level] - hanging[level], textPositionTwips: textAt[level],
    };
  });
}

export type FormatKey = "bold" | "italic" | "underline" | "strike" | "code" | "superscript" | "subscript";

/** Paragraph alignment passed to DocxEditor.setAlignment. */
export type EditorAlignment = "left" | "center" | "right" | "justify";

/**
 * What {@link DocxEditor.setParagraphFormat} sets — Word's Paragraph dialog. Every field is
 * optional (left out = unchanged). Lengths are twips: 1440 an inch, 567 a centimetre, 20 a point.
 */
export interface EditorParagraphFormat {
  alignment?: EditorAlignment;
  /** Left indent, absolute; may be negative (into the margin). */
  leftIndent?: number;
  /** Right indent, absolute; may be negative. */
  rightIndent?: number;
  /** First-line indent (0 = none); removes a hanging indent. Not with `hangingIndent`. */
  firstLineIndent?: number;
  /** Hanging indent; removes a first-line indent. Not with `firstLineIndent`. */
  hangingIndent?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  /** Under `"auto"` (the default) 240ths of a line (240 single, 360 one and a half, 480
   *  double); under `"exact"` and `"atLeast"` twips. */
  lineSpacing?: number;
  lineSpacingRule?: "auto" | "exact" | "atLeast";
  /** Don't add space between paragraphs of the same style. */
  contextualSpacing?: boolean;
  widowControl?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
}

/** An engine paragraph snapshot in {@link EditorParagraphFormat}'s names (what the dialog and
 *  `paragraphFormat` take), so a read can be handed straight back. */
function asEditorParagraphFormat(f: ParagraphFormatting): EditorParagraphFormat {
  const out: EditorParagraphFormat = {
    alignment: f.alignment,
    leftIndent: f.leftIndentTwips,
    rightIndent: f.rightIndentTwips,
    spacingBefore: f.spacingBeforeTwips,
    spacingAfter: f.spacingAfterTwips,
    lineSpacing: f.lineSpacing,
    lineSpacingRule: f.lineSpacingRule,
    contextualSpacing: f.contextualSpacing,
    widowControl: f.widowControl,
    keepNext: f.keepNext,
    keepLines: f.keepLines,
    pageBreakBefore: f.pageBreakBefore,
  };
  // w:ind holds one of the two; a hanging indent wins when a document has both, as in Word.
  if (f.hangingIndentTwips) out.hangingIndent = f.hangingIndentTwips;
  else out.firstLineIndent = f.firstLineIndentTwips ?? 0;
  return out;
}

/** What a paste does with the source's formatting: Word's Merge Formatting (`"merge"`), Keep
 *  Source Formatting (`"keep"`) or Keep Text Only (`"text"`). */
export type PasteFormatting = "merge" | "keep" | "text";

/** Format Painter: off, on for one paint (`"once"`), or on until turned off (`"sticky"`). */
export type FormatPainterState = "off" | "once" | "sticky";

/** What Format Painter copied: the run formatting, and the paragraph's when the copy took it. */
interface CopiedFormatting {
  run: Record<string, unknown>;
  paragraph: {
    styleId: string | null;
    format: EditorParagraphFormat;
    list: { numId: number; level: number } | null;
  } | null;
}

/** Show/Hide ¶: the marks are pseudo-elements placed where they would stand but taking no room
 *  (absolutely positioned at their static position), so no line wraps differently and page view
 *  lays out the same pages. */
const FORMATTING_MARKS_CSS = `
.docx-show-marks [data-anchor][contenteditable="true"]::after {
  content: "¶"; position: absolute; color: #4a7fc1; font-weight: normal; font-style: normal;
  text-decoration: none; pointer-events: none; user-select: none; padding-left: 1px;
  /* (an absolute box indents its own first line by the paragraph's text-indent: without this a
     first-line indent pushed the mark that far past its place, out to the cell's edge) */
  text-indent: 0;
}
.docx-show-marks .page-box [data-section-break][contenteditable="true"]::after {
  content: ":::::: " var(--docx-section-break-nextPage, "Section break (Next page)") " ::::::";
  white-space: nowrap; font: 11px/1.4 system-ui, sans-serif; padding-left: 6px;
}
.docx-show-marks .page-box [data-section-break="continuous"][contenteditable="true"]::after {
  content: ":::::: " var(--docx-section-break-continuous, "Section break (Continuous)") " ::::::";
}
.docx-show-marks .page-box [data-section-break="evenPage"][contenteditable="true"]::after {
  content: ":::::: " var(--docx-section-break-evenPage, "Section break (Even page)") " ::::::";
}
.docx-show-marks .page-box [data-section-break="oddPage"][contenteditable="true"]::after {
  content: ":::::: " var(--docx-section-break-oddPage, "Section break (Odd page)") " ::::::";
}
.docx-show-marks .page-box [data-section-break="nextColumn"][contenteditable="true"]::after {
  content: ":::::: " var(--docx-section-break-nextColumn, "Section break (Next column)") " ::::::";
}
.docx-show-marks [data-docx-tab] { position: relative; }
.docx-show-marks [data-docx-tab]::after {
  content: "→"; position: absolute; left: 0; right: 0; text-align: center; text-indent: 0; color: #4a7fc1;
  font-weight: normal; font-style: normal; text-decoration: none; pointer-events: none; user-select: none;
}
`;
/** The continuous view's section-break line, as Word's Draft view draws it: a double dotted rule
 *  across the top of every section after the first, naming how that section starts. A host can
 *  translate the labels through the `--docx-section-break-*` custom properties (quoted strings).
 *  Page view lays sections out on their own pages and draws no line. */
const SECTION_BREAK_CSS = `
[data-section-index]:not([data-section-index="0"])::before {
  content: var(--docx-section-break-nextPage, "Section break (Next page)");
  display: block; margin: 8px 0 10px; padding-top: 2px; border-top: 3px double #8fa9cc;
  color: #4a7fc1; font: 11px/1.4 system-ui, sans-serif; font-style: normal; font-weight: normal;
  text-align: center; text-indent: 0; letter-spacing: normal; text-transform: none;
  pointer-events: none; user-select: none;
}
[data-section-index][data-section-type="continuous"]:not([data-section-index="0"])::before {
  content: var(--docx-section-break-continuous, "Section break (Continuous)");
}
[data-section-index][data-section-type="evenPage"]:not([data-section-index="0"])::before {
  content: var(--docx-section-break-evenPage, "Section break (Even page)");
}
[data-section-index][data-section-type="oddPage"]:not([data-section-index="0"])::before {
  content: var(--docx-section-break-oddPage, "Section break (Odd page)");
}
[data-section-index][data-section-type="nextColumn"]:not([data-section-index="0"])::before {
  content: var(--docx-section-break-nextColumn, "Section break (Next column)");
}
.page-box [data-section-index]::before, .docx-paginated [data-section-index]::before { content: none !important; }
`;
const sectionBreakStyledDocuments = new WeakSet<Document>();
function ensureSectionBreakStyles(doc: Document): void {
  if (sectionBreakStyledDocuments.has(doc)) return;
  sectionBreakStyledDocuments.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusSectionBreaks = "true";
  style.textContent = SECTION_BREAK_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}

const formattingMarksStyledDocuments = new WeakSet<Document>();
function ensureFormattingMarksStyles(doc: Document): void {
  if (formattingMarksStyledDocuments.has(doc)) return;
  formattingMarksStyledDocuments.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusFormattingMarks = "true";
  style.textContent = FORMATTING_MARKS_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}

const formatPainterStyledDocuments = new WeakSet<Document>();

function ensureFormatPainterStyles(doc: Document): void {
  if (formatPainterStyledDocuments.has(doc)) return;
  formatPainterStyledDocuments.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusFormatPainter = "true";
  style.textContent = FORMAT_PAINTER_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}

/** What there is to paste: the clipboard's HTML, its plain text, and a picture it holds. */
export interface PasteContent {
  html?: string;
  text?: string;
  image?: Blob;
}

/** Where a paste lands, measured (and the block's typing committed) when it began. */
interface PasteTarget {
  block: HTMLElement;
  unid: string;
  fullId: string;
  start: number;
  length: number;
  text: string;
}

async function blobToDataUrl(blob: Blob): Promise<string> {
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

/** The selection's content-text {start,length} within `block` (excludes markers), or null. */
function selectionSpanIn(block: HTMLElement): { start: number; length: number } | null {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (range.collapsed) return null;
  if (!block.contains(range.startContainer) || !block.contains(range.endContainer)) return null;
  const start = contentOffsetOf(block, range.startContainer, range.startOffset);
  const end = contentOffsetOf(block, range.endContainer, range.endOffset);
  // Normalized into the committed run-text space: every consumer feeds this straight to a
  // DocxSession op, which rejects a span that overshoots the committed length.
  const span = trimmedSpan(block, { start: Math.min(start, end), length: Math.abs(end - start) });
  return span.length > 0 ? span : null;
}

/**
 * A DOM Range over the content text spanning [start, start+length) within `el` (skips markers),
 * or null when the block is detached or the offsets do not resolve. Building the range is kept
 * separate from selecting it so a caller that only wants to PAINT a span (find preview) does not
 * have to touch the selection — and therefore focus — at all.
 */
function contentRangeIn(el: HTMLElement, start: number, length: number): Range | null {
  // The block may have been swapped out of the document by a re-render before this runs
  // (e.g. a focus-stealing toolbar control firing twice). addRange on a detached range
  // throws "the given range isn't in document" — skip rather than warn.
  if (!el.isConnected) return null;
  const doc = el.ownerDocument;
  const range = doc.createRange();
  const from = contentPositionIn(el, start);
  const to = contentPositionIn(el, start + length);
  try {
    range.setStart(from.node, from.offset);
    range.setEnd(to.node, to.offset);
  } catch {
    return null;
  }
  return range;
}

/** Restore a content-text selection spanning [start, start+length) within `el` (skips markers). */
function selectRange(el: HTMLElement, start: number, length: number): void {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || !el.isConnected) return;
  el.focus();
  const range = contentRangeIn(el, start, length);
  if (!range) return;
  sel.removeAllRanges();
  sel.addRange(range);
}

// ─── Find painting ──────────────────────────────────────────────────────────
//
// A find box has to show WHERE the hits are while the user is still typing into it, so it cannot
// express "the current match" as a selection: putting a caret in a contenteditable block focuses
// that block, and the next keystroke lands in the document instead of the search field. These
// highlights paint the same information with no focus and no DOM mutation — the CSS Custom
// Highlight API takes plain Ranges over the live text and styles them from the stylesheet.

/** Every match currently on screen. */
const FIND_HIGHLIGHT = "docxodus-find";
/** The one the find bar's counter is pointing at. */
const FIND_HIGHLIGHT_ACTIVE = "docxodus-find-active";
/** Ranges are cheap, but a pathological query ("e" in a book) is not worth painting in full. */
const FIND_PAINT_LIMIT = 500;
/** A find stops counting here: a pattern that matches every character of a book is a mistake. */
const FIND_HIT_LIMIT = 20000;

const findHighlightStyledDocuments = new WeakSet<Document>();
/**
 * Which editor last painted into the shared registry. Only that instance may clear it, so a
 * second surface on the page cannot wipe the first one's find painting when its own find bar
 * closes. Two live searches still share one painting — the newer one wins, which is what a
 * single visible "current match" should mean.
 */
let findPaintOwner: object | null = null;

function ensureFindHighlightStyles(doc: Document): void {
  if (findHighlightStyledDocuments.has(doc)) return;
  findHighlightStyledDocuments.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusFindHighlight = "true";
  style.textContent = `
::highlight(${FIND_HIGHLIGHT}) { background-color: rgba(15, 118, 110, .16); }
::highlight(${FIND_HIGHLIGHT_ACTIVE}) { background-color: rgba(15, 118, 110, .38); }
`;
  (doc.head ?? doc.documentElement).appendChild(style);
}

/** The highlight registry is global to the document, so two editors on one page share it. */
interface HighlightRegistry {
  set(name: string, highlight: unknown): void;
  delete(name: string): void;
}

function highlightRegistry(): HighlightRegistry | null {
  if (typeof CSS === "undefined") return null;
  const registry = (CSS as unknown as { highlights?: HighlightRegistry }).highlights;
  const ctor = (globalThis as unknown as { Highlight?: unknown }).Highlight;
  return registry && typeof ctor === "function" ? registry : null;
}

function makeHighlight(ranges: Range[]): unknown {
  const ctor = (globalThis as unknown as {
    Highlight: new (...ranges: Range[]) => unknown;
  }).Highlight;
  return new ctor(...ranges);
}

/**
 * True when `el`'s immediate parent is a paragraph-border `<div>` the full render wrapped it in
 * (CreateBorderDivs groups visibly-bordered paragraphs into a div). The body wrapper div has no
 * border. Splitting/merging such a block must re-render the whole document so the converter can
 * re-group the border boxes — an in-place node swap would leave the new (often borderless)
 * paragraph stranded inside the stale border div, drawing the rule's line under its text.
 */
function inBorderWrapper(el: HTMLElement): boolean {
  const style = el.parentElement?.getAttribute("style") ?? "";
  return /border-(top|bottom|left|right):\s*(?!none)[^;]+/i.test(style);
}

/**
 * Where a selection's format is read: the first character it covers. A range that starts at the
 * end of a text node (as the selection the editor restores after formatting part of a paragraph
 * does: "Hello from |" .. "world|"), or between elements, begins in the next text node, not the
 * one before. A collapsed caret stays where it is: it takes the format of the text before it.
 */
function selectionFormatNode(range: Range): Node {
  if (range.collapsed) return range.startContainer;
  const root = range.commonAncestorContainer;
  const scope = root.nodeType === 3 ? root.parentNode : root;
  if (!scope) return range.startContainer;
  const walker = (scope.ownerDocument ?? document).createTreeWalker(scope, NodeFilter.SHOW_TEXT);
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    if (!range.intersectsNode(t)) continue;
    const from = t === range.startContainer ? range.startOffset : 0;
    const to = t === range.endContainer ? range.endOffset : (t.textContent ?? "").length;
    if (to > from) return t;
  }
  return range.startContainer;
}

/** Whether the current selection's start already carries `key`, read from computed style. */
function selectionHasFormat(key: FormatKey, fallback: HTMLElement): boolean {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  let el: HTMLElement | null = fallback;
  if (sel && sel.rangeCount > 0) {
    const n = selectionFormatNode(sel.getRangeAt(0));
    el = n.nodeType === 3 ? n.parentElement : (n as HTMLElement);
  }
  if (!el || typeof getComputedStyle !== "function") return false;
  const cs = getComputedStyle(el);
  switch (key) {
    case "bold": return fontWeightIsBold(cs.fontWeight);
    case "italic": return cs.fontStyle === "italic" || cs.fontStyle === "oblique";
    case "underline":
    case "strike": {
      // Decorations propagate through inline descendants without being CSS-inherited. A run
      // can carry its underline on an outer span and its independent strike on an inner one.
      const line = key === "underline" ? "underline" : "line-through";
      for (let node: HTMLElement | null = el; node; node = node.parentElement) {
        if (getComputedStyle(node).textDecorationLine.includes(line)) return true;
        if (node === fallback) break;
      }
      return false;
    }
    case "code": return /mono|courier|consolas/i.test(cs.fontFamily);
    case "superscript": return cs.verticalAlign === "super" || !!el.closest("sup");
    case "subscript": return cs.verticalAlign === "sub" || !!el.closest("sub");
    default: return false;
  }
}

/** Build the full ConvertDocxToHtmlComplete arg list (revisionPresentation = last arg). */
function completeArgs(
  bytes: Uint8Array,
  cssPrefix: string,
  fabricate: boolean,
  paginated: boolean,
  scale: number,
  renderTrackedChanges: boolean,
  comments: boolean,
): any[] {
  return [
    // Comment mode 1 = Inline: the commented runs are wrapped in highlight spans and the
    // reference becomes an (editor-hidden) marker, which is what the comment gutter positions
    // its bubbles against. -1 renders no comment markup at all.
    bytes, "Document", cssPrefix, fabricate, "", comments ? 1 : -1, "comment-",
    /* paginationMode */ paginated ? 1 : 0, /* paginationScale */ scale, "page-",
    false, 0, "annot-",
    // Footnotes/endnotes ON: they are document content, and the editor makes the rendered note
    // paragraphs editable. Must stay in step with DocxSessionOps.RenderHtml (the remount path),
    // whose output has to match this first paint byte-for-byte.
    /* renderFootnotesAndEndnotes */ true, /* renderHeadersAndFooters */ paginated,
    renderTrackedChanges, true, true, false, null, /* stampAnchors */ true,
    /* revisionPresentation: Docxodus (block renders keep it) */ 0,
  ];
}

export class DocxEditor {
  private readonly exports: DocxEditorExports;
  private container: HTMLElement;
  private readonly handle: number;
  private readonly options: Required<
    Omit<DocxEditorOptions, "onEdit" | "onEditFailed" | "onMove" | "onStoryChange" | "onCommentsChange" | "showRevisions" | "resolvePasteImage" | "listDefaults" | "onFormatPainterChange" | "headlineCaseLists" | "showFormattingMarks" | "onFormattingMarksChange" | "onPictureChange">
  > &
    Pick<DocxEditorOptions, "onEdit" | "onEditFailed" | "onMove" | "onStoryChange" | "onCommentsChange" | "showRevisions" | "resolvePasteImage" | "listDefaults" | "onFormatPainterChange" | "headlineCaseLists" | "showFormattingMarks" | "onFormattingMarksChange" | "onPictureChange">;
  /** Map a block's current bare unid → its full kind:scope:unid (DocxSession anchor). */
  private readonly unidToFullId = new Map<string, string>();
  /** The element whose [data-anchor] descendants are the editable blocks (container or page container). */
  private editRoot: HTMLElement;
  /** The most recently focused editable block — the target for ribbon/format commands. */
  private activeBlock: HTMLElement | null = null;
  /** Redraws Word's line numbers in the page view after the text under them changes. */
  private lineNumberWatch: { observer: MutationObserver; timer: number } | null = null;
  private closed = false;
  /** The anchor of the paragraph the ruler last showed (see rulerBlock). */
  private rulerAnchor: string | null = null;
  /** The pause after which `typingUndo: "word"` commits the typing so far (see onTypingInput). */
  private typingPause: ReturnType<typeof setTimeout> | null = null;
  /**
   * Re-entrancy guard for node replacement. Replacing a contenteditable block that still holds
   * focus removes the focused node, which fires a SYNCHRONOUS `blur` → re-enters commitBlock; the
   * interleaved second replaceWith then throws NotFoundError ("node ... no longer a child") and the
   * structural edit (split/merge/format) is lost. While this flag is set, commitBlock no-ops.
   */
  private replacing = false;
  /**
   * A pristine copy of the paginated note number each wired block opened with, keyed by the block.
   * The number is chrome the converter derives from the whole note, so a single-block repaint
   * cannot re-create it, and a selection replaced across it removes it from the live block; the
   * swap in {@link replaceNode} restores it from here.
   */
  private readonly noteNumbers = new WeakMap<HTMLElement, HTMLElement>();

  /** Editor-owned block-move chrome. It is deliberately outside the rendered unit tree. */
  private blockDragHandle: HTMLElement | null = null;
  private blockDropIndicator: HTMLElement | null = null;
  private blockMoveMenu: HTMLElement | null = null;
  private blockMoveLive: HTMLElement | null = null;
  private blockDragSource: HTMLElement | null = null;
  /** Whether hover or focus currently wants a handle for `blockDragSource`.
   *  Tracked apart from the handle's `display`, which the viewport clip also drives: a handle
   *  withdrawn because its block scrolled out of view has to come back when it scrolls in. */
  private blockHandleWanted = false;
  private blockDragCleanup: Array<() => void> = [];
  /** Block boxes measured at drag start — see `BlockDropZone`. Empty when no drag is in flight. */
  private dropZones: BlockDropZone[] = [];
  /** Combined scroll offset when `dropZones` was measured, and the scroller measured against. */
  private dropZoneOrigin = 0;
  private dropZoneScroller: HTMLElement | null = null;
  private blockDragging = false;
  private blockDragPointerDown = false;
  /** Anchors the current drag source may legally move next to, per the engine's own rules.
   *  Null when the bridge predates ValidMoveTargets — then every block is offered, as before. */
  private blockMoveTargets: Map<string, { before: boolean; after: boolean }> | null = null;
  /** Memoized `ValidMoveTargets` answers, keyed by source anchor and dropped whenever an edit
   *  lands. The legal-target set is a property of the document, so hovering back and forth over
   *  the same blocks between edits must not re-ask the engine. */
  private blockMoveTargetCache = new Map<string, Map<string, { before: boolean; after: boolean }> | null>();
  /** Cancels the pending idle prefetch of the hovered block's targets — see `showBlockHandle`. */
  private blockMoveTargetPrefetch: (() => void) | null = null;
  /** Why the last move was refused, verbatim from the engine — diagnostics, not announcement copy. */
  lastMoveError: string | null = null;

  /**
   * The last real (non-collapsed) text selection inside an editable block. A toolbar control that
   * must take focus to be used — the font-size combobox — blurs the block and collapses the live
   * selection, so without this an operation triggered from such a control could only target the
   * whole paragraph (S-1 smoke-test finding 3). Refreshed whenever a non-empty selection sits in a
   * block, and cleared when a caret is collapsed inside a block (so it never goes stale).
   */
  private lastSelection: { unid: string; span: { start: number; length: number } } | null = null;

  /**
   * Stable bookmark for a selection spanning independent editable blocks. Native controls such as
   * a font-size combobox can take focus and collapse the live DOM selection; block ids + content
   * offsets let the command restore that same range before applying. This is the multi-block
   * counterpart to lastSelection above.
   */
  private lastCrossBlockSelection: CrossBlockSelectionBookmark | null = null;

  /** State for the mouse-selection bridge between independent contenteditable block hosts. */
  private dragSelection: CrossBlockDragSelection | null = null;

  /**
   * Which end of a selection the editor made across blocks (drag, Shift+click, Shift+arrows,
   * Ctrl+A) is fixed and which moves. Such a selection is set with addRange, which forgets its
   * direction, so Shift+arrow keys would otherwise always move its end. Valid only while the live
   * range still spans exactly these two points (see {@link selectionEnds}).
   */
  private extendedSelection: { anchor: DomSelectionPoint; focus: DomSelectionPoint } | null = null;

  /** The docked header/footer bands, when `options.headerFooter` is on. */
  private region: HeaderFooterRegion | null = null;

  /** The comment gutter, when `options.comments` is on. */
  private gutter: CommentGutter | null = null;

  /** Story host the caret is in (page view / band), published for chrome via `onStoryChange`. */
  private activeStory: BandWhich | null = null;

  /** Page geometry + fit-to-width zoom for the mounted document. Re-attached on every mount. */
  private readonly viewport: DocumentViewport;

  /** Why the last reconcile() fell back to a full remount (null = it patched). For
   *  diagnostics/specs; not part of the public API. */
  private lastReconcileFallback: string | null = null;

  private constructor(
    container: HTMLElement,
    exports: DocxEditorExports,
    handle: number,
    options: DocxEditor["options"],
  ) {
    this.container = container;
    this.exports = exports;
    this.handle = handle;
    this.options = options;
    this.editRoot = container;
    container.addEventListener("copy", this.onCopy);
    container.addEventListener("cut", this.onCut);
    container.addEventListener("paste", this.onPaste);
    this.viewport = new DocumentViewport(container, {
      columnWidth: options.columnWidth,
      fitToWidth: options.fitToWidth,
      scale: options.scale,
    });
    if (typeof document !== "undefined") {
      document.addEventListener("selectionchange", this.onSelectionChange);
      document.addEventListener("mousedown", this.onMouseDown, true);
      document.addEventListener("mousemove", this.onMouseMove, true);
      document.addEventListener("mouseup", this.onMouseUp, true);
      document.addEventListener("keydown", this.onPictureKey, true);
    }
  }

  /** Native copy treats flex tab-layout wrappers as line boundaries and omits their empty tab.
   *  Normalize only the clipboard clone; retain inline formatting in the rich representation. */
  private readonly onCopy = (ev: ClipboardEvent): void => {
    const selection = this.container.ownerDocument.defaultView?.getSelection();
    if (this.closed || ev.defaultPrevented || !ev.clipboardData || !selection?.rangeCount || selection.isCollapsed) return;
    const range = selection.getRangeAt(0);
    if (!this.container.contains(range.startContainer) || !this.container.contains(range.endContainer)) return;
    const clone = this.container.ownerDocument.createElement("div");
    clone.append(range.cloneContents());
    if (!clone.querySelector("[data-docx-tab]")) return;
    // Range.cloneContents keeps selected descendants in document order, including partial
    // ancestors. Read styles and placeholder status from their live counterparts before the
    // clone loses its editing host and the document's class-based stylesheet.
    const inlineTextStyle = (copy: HTMLElement, original: Element): CSSStyleDeclaration => {
      const style = this.container.ownerDocument.defaultView!.getComputedStyle(original);
      for (const key of ["direction", "unicode-bidi", "white-space", "font-family", "font-size", "font-weight", "font-style", "font-variant", "font-kerning", "letter-spacing", "color", "background-color", "text-decoration", "text-transform", "vertical-align"])
        copy.style.setProperty(key, style.getPropertyValue(key));
      return style;
    };
    const common = range.commonAncestorContainer;
    inlineTextStyle(clone, common.nodeType === 1 ? common as Element : common.parentElement!);
    const source = this.container.ownerDocument.createTreeWalker(common, NodeFilter.SHOW_ELEMENT,
      { acceptNode: node => range.intersectsNode(node) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT });
    clone.querySelectorAll<HTMLElement>("*").forEach(copy => {
      const original = source.nextNode() as HTMLElement;
      if (isPlaceholderBreak(original)) { copy.remove(); return; }
      const display = inlineTextStyle(copy, original).display;
      if (display.includes("flex") || display === "none") copy.style.display = display;
    });
    const context = (common.nodeType === 1 ? common as Element : common.parentElement!)
      .closest<HTMLAnchorElement>("a[href]");
    if (context) {
      const link = context.cloneNode(false) as HTMLAnchorElement;
      inlineTextStyle(link, context);
      link.append(...Array.from(clone.childNodes));
      clone.append(link);
    }
    clone.querySelectorAll<HTMLElement>("span").forEach(span => {
      if (span.style.display.includes("flex")) {
        span.style.display = "inline";
        span.style.removeProperty("width");
      }
    });
    clone.querySelectorAll<HTMLElement>("[data-docx-tab]").forEach(tab => {
      tab.textContent = "\t";
      tab.style.cssText = "display: inline; white-space: pre;";
    });
    const richHtml = clone.outerHTML;
    clone.style.position = "fixed";
    clone.style.left = "-100000px";
    clone.style.whiteSpace = "pre-wrap";
    clone.setAttribute("aria-hidden", "true");
    this.container.ownerDocument.body.append(clone);
    try {
      ev.clipboardData.setData("text/plain", stripBidi(clone.innerText));
      ev.clipboardData.setData("text/html", richHtml);
      ev.preventDefault();
    } finally { clone.remove(); }
  };

  /** Ctrl+V and Shift+Insert in a block: the editor pastes through the engine (`pasteContent`),
   *  never the browser's own insertion into the block. (Ctrl+Shift+V is Paste Formatting, and
   *  Ctrl+Alt+Shift+V the text only: see onKeydown.) */
  private readonly onPaste = (ev: ClipboardEvent): void => {
    if (this.closed || ev.defaultPrevented || !ev.clipboardData) return;
    const block = ev.target instanceof Node ? this.editableBlockOf(ev.target) : null;
    if (!block) return; // a comment box or other field pastes as the browser does
    // A selection across paragraphs that typing cannot replace (into a table or a band) keeps the
    // browser's own paste, as onPasteOverBlocks leaves it.
    const cross = this.crossBlockSelection();
    if (cross && !this.crossBlockEditable(cross)) return;
    ev.preventDefault();
    const data = ev.clipboardData;
    // A selection across paragraphs goes first (one edit), then the paste lands at the caret it leaves.
    let target: HTMLElement | null = block;
    if (this.crossBlockSelection()) {
      this.replaceCrossBlockSelection("");
      target = this.activeBlock;
      if (!target) return;
    }
    const image = Array.from(data.files).find((f) => f.type.startsWith("image/"));
    this.activeBlock = target;
    void this.pasteContent({ html: data.getData("text/html"), text: data.getData("text/plain"), image });
  };

  /** Track the last meaningful selection so focus-stealing toolbar controls can still target it. */
  private readonly onSelectionChange = (): void => {
    if (this.closed) return;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.rangeCount === 0) return; // no selection info — keep the cache as-is
    const range = sel.getRangeAt(0);
    const startBlock = this.editableBlockOf(range.startContainer);
    const endBlock = this.editableBlockOf(range.endContainer);
    const block = startBlock ?? endBlock;
    if (!block) return; // selection is outside the editor (e.g. a toolbar field) — keep the cache
    if (range.collapsed) {
      this.lastSelection = null; // an explicit caret in a block — drop any stale selection
      this.lastCrossBlockSelection = null;
      return;
    }
    if (
      startBlock && endBlock && startBlock !== endBlock &&
      this.ownerRoot(startBlock) === this.ownerRoot(endBlock)
    ) {
      const startUnid = startBlock.getAttribute("data-anchor");
      const endUnid = endBlock.getAttribute("data-anchor");
      if (startUnid && endUnid) {
        this.lastSelection = null;
        this.lastCrossBlockSelection = {
          startUnid,
          startOffset: contentOffsetOf(startBlock, range.startContainer, range.startOffset),
          endUnid,
          endOffset: contentOffsetOf(endBlock, range.endContainer, range.endOffset),
        };
      }
      return;
    }
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    this.lastCrossBlockSelection = null;
    const span = selectionSpanIn(block);
    if (span) this.lastSelection = { unid, span };
  };

  /** Start tracking a normal primary-button text drag inside one editable block. */
  private readonly onMouseDown = (event: MouseEvent): void => {
    if (this.closed || !this.options.editable) return;
    if (this.pictureMouseDown(event)) return;
    if (event.button !== 0) return;
    if (this.extendSelectionByClick(event)) { this.clearDragSelection(); return; }
    const target = event.target instanceof Node ? event.target : null;
    const origin = this.editableBlockOf(target);
    if (!origin || isInMarker(target)) return;
    const anchor = caretPointFromClient(document, event.clientX, event.clientY);
    if (!anchor || this.editableBlockOf(anchor.node) !== origin) return;
    this.clearDragSelection();
    this.dragSelection = {
      anchor,
      origin,
      crossedBlockBoundary: false,
      focus: null,
      frame: null,
    };
  };

  /**
   * Apply the latest cross-block endpoint after the browser finishes its native mousemove
   * selection update. Firefox rewrites Selection back into the originating contenteditable after
   * event dispatch even when mousemove is cancelled; writing in requestAnimationFrame wins that
   * race and happens before paint. Coalescing also avoids rebuilding a Range for every raw pointer
   * event when the mouse is moving faster than the display can refresh.
   */
  private queueDragSelection(drag: CrossBlockDragSelection, focus: DomSelectionPoint): void {
    drag.focus = focus;
    if (drag.frame !== null) return;
    const view = this.container.ownerDocument.defaultView;
    if (!view) {
      setSelectionBetween(drag.anchor, focus);
      return;
    }
    drag.frame = view.requestAnimationFrame(() => {
      drag.frame = null;
      if (this.dragSelection !== drag || !drag.focus) return;
      setSelectionBetween(drag.anchor, drag.focus);
    });
  }

  /** Cancel a queued repaint and discard the current gesture. */
  private clearDragSelection(): void {
    const drag = this.dragSelection;
    if (drag && drag.frame !== null) {
      this.container.ownerDocument.defaultView?.cancelAnimationFrame(drag.frame);
    }
    this.dragSelection = null;
  }

  /**
   * Browsers fence native mouse selection at a contenteditable host boundary. Once a drag reaches
   * another block in the same OOXML story, take over just that gesture and create the cross-block
   * Selection the editor's existing multi-block command path consumes. Intra-block selection stays
   * entirely native, and separate hosts remain intact for safe per-anchor commits.
   */
  private readonly onMouseMove = (event: MouseEvent): void => {
    const drag = this.dragSelection;
    if (!drag) return;
    if ((event.buttons & 1) === 0) {
      this.clearDragSelection();
      return;
    }
    const focus = caretPointFromClient(document, event.clientX, event.clientY);
    if (!focus || isInMarker(focus.node)) return;
    const focusBlock = this.editableBlockOf(focus.node);
    if (!focusBlock || this.ownerRoot(focusBlock) !== this.ownerRoot(drag.origin)) return;
    if (focusBlock !== drag.origin) drag.crossedBlockBoundary = true;
    if (!drag.crossedBlockBoundary) return;
    event.preventDefault();
    this.queueDragSelection(drag, focus);
  };

  /** Commit the final cross-block endpoint before releasing the gesture state. */
  private readonly onMouseUp = (event: MouseEvent): void => {
    // Format Painter paints where the click or drag ends, once the selection has settled.
    if (this.painter !== "off" && event.button === 0
      && this.editableBlockOf(event.target instanceof Node ? event.target : null)) {
      setTimeout(() => this.paintAtSelection(), 0);
    }
    const drag = this.dragSelection;
    if (!drag) return;
    if (drag.crossedBlockBoundary) {
      const focus = caretPointFromClient(document, event.clientX, event.clientY);
      const focusBlock = focus ? this.editableBlockOf(focus.node) : null;
      if (focus && focusBlock && this.ownerRoot(focusBlock) === this.ownerRoot(drag.origin)) {
        event.preventDefault();
        setSelectionBetween(drag.anchor, focus);
        this.extendedSelection = { anchor: drag.anchor, focus };
      }
    }
    this.clearDragSelection();
  };

  /** The editable block (contenteditable [data-anchor]) containing `node`, if any, within this editor.
   *  Fenced by `container`, not `editRoot`, so header/footer band blocks — which live outside the
   *  body edit root by design — also register. The fence still rejects other editors on the page. */
  private editableBlockOf(node: Node | null): HTMLElement | null {
    if (!node) return null;
    const start = node.nodeType === 1 ? (node as HTMLElement) : node.parentElement;
    const block = start?.closest<HTMLElement>('[data-anchor][contenteditable="true"]') ?? null;
    return block && this.container.contains(block) ? block : null;
  }

  /**
   * The root owning `el`'s sibling block list: its header/footer band's story container, else the
   * body edit root. Keeps a multi-block selection from spanning a band and the body, whose block
   * lists belong to different OOXML parts.
   */
  private ownerRoot(el: HTMLElement): HTMLElement {
    return this.region?.blockRootOf(el) ?? this.editRoot;
  }

  /** True when `el` is a header/footer band block rather than a body block. */
  private isBandBlock(el: HTMLElement): boolean {
    return !!this.region?.contains(el);
  }

  /**
   * Repaint after an edit to `block` that would otherwise remount the whole document: a band
   * repaints only itself (a story is one to three paragraphs), leaving the body DOM — and the
   * user's place in it — untouched; a body edit reconciles incrementally. `forceRemount` is
   * for ops whose repaint provably needs whole-document context the reconciler cannot see:
   * list membership/level changes (sibling numbering shifts without sibling XML changing)
   * and border-div regrouping (HR insert, clearBorders).
   */
  private refreshAfter(block: HTMLElement, focusIndex: number, caretAtEnd = false, forceRemount = false): void {
    const band = this.region?.bandOf(block);
    if (band) {
      this.region!.refresh(this.region!.whichOf(band));
      return;
    }
    if (forceRemount) this.remount(focusIndex, caretAtEnd);
    else this.reconcile(focusIndex, caretAtEnd);
  }

  /** Open a document, render it into `container`, and wire up editing. */
  static open(
    container: HTMLElement,
    bytes: Uint8Array,
    exports: DocxEditorExports,
    options: DocxEditorOptions = {},
  ): DocxEditor {
    const opts = DocxEditor.resolveOptions(options);
    const editor = DocxEditor.openSession(container, bytes, exports, opts);
    try {
      editor.refreshAnchorMap();
      if (opts.headerFooter) editor.createRegion();
      // First paint goes through the session-attached editor render when the bundle has it, so
      // the comment markup (and everything else in the profile) is exactly what a remount will
      // produce; older bundles take the bytes path with the same profile.
      const fullHtml = editor.renderFullHtml(bytes);
      if (opts.paginated) editor.mountPaginated(fullHtml);
      else editor.mountHtml(fullHtml);
      editor.finishMount();
      return editor;
    } catch (error) {
      editor.close();
      throw error;
    }
  }

  /**
   * Open a document without holding the main thread for the whole mount (issue #776).
   *
   * {@link open} renders and wires the entire document in one synchronous task, which on a
   * long document is the largest single block of a reading-and-editing session. This variant
   * pays only the session open up front, then mounts the document in windows of body units,
   * yielding to the event loop between them: the chrome — stylesheet, section wrappers with
   * their page geometry, footnote and endnote sections — comes from one cheap engine render
   * that sees no body content, and each window from the engine's block renderer, which lays
   * units out exactly as the full render does (the plan's border-box grouping is never split).
   * The DOM that results is the one {@link open} produces; blocks already mounted are editable
   * while later windows are still landing, and edits made meanwhile are honoured because every
   * window renders from the live session.
   *
   * Paginated mounts assemble the windows off-screen and paginate once at the end, since
   * pagination is a one-shot flow over the whole document; the main thread is still free
   * between windows, and `onProgress` lets a host show the fill.
   *
   * A bundle that predates the windowed renders mounts synchronously, exactly like {@link open}.
   */
  static async openAsync(
    container: HTMLElement,
    bytes: Uint8Array,
    exports: DocxEditorExports,
    options: DocxEditorOpenAsyncOptions = {},
  ): Promise<DocxEditor> {
    const bridge = exports.DocxSessionBridge;
    if (
      // Page view lays every page out at once: there are no windows to mount.
      options.paginated
      || typeof bridge.RenderEditorChromeHtml !== "function"
      || typeof bridge.RenderEditorRangeHtml !== "function"
      || typeof bridge.ListRenderedBlocks !== "function"
    ) {
      return DocxEditor.open(container, bytes, exports, options);
    }
    const opts = DocxEditor.resolveOptions(options);
    const editor = DocxEditor.openSession(container, bytes, exports, opts);
    try {
      editor.refreshAnchorMap();
      if (opts.headerFooter) editor.createRegion();
      await editor.mountWindowed(
        Math.max(1, Math.floor(options.windowSize ?? 24)),
        options.onProgress,
      );
      editor.finishMount();
      return editor;
    } catch (error) {
      editor.close();
      throw error;
    }
  }

  private static resolveOptions(options: DocxEditorOptions): DocxEditor["options"] {
    return {
      cssPrefix: options.cssPrefix ?? "docx-",
      fabricateClasses: options.fabricateClasses ?? false,
      editable: options.editable ?? true,
      pasteFormatting: options.pasteFormatting ?? "merge",
      resolvePasteImage: options.resolvePasteImage,
      paginated: options.paginated ?? false,
      scale: options.scale ?? 1,
      columnWidth: options.columnWidth ?? "section",
      fitToWidth: options.fitToWidth ?? true,
      headerFooter: options.headerFooter ?? false,
      blockDrag: options.blockDrag ?? false,
      trackedChanges: options.trackedChanges ?? TrackedChangeMode.Accept,
      showRevisions: options.showRevisions,
      revisionAuthor: options.revisionAuthor ?? "docxodus",
      revisionIdentities: options.revisionIdentities?.map(identity => ({...identity})) ?? null,
      comments: options.comments ?? true,
      commentAuthor: options.commentAuthor ?? options.revisionAuthor ?? "Reviewer",
      typingUndo: options.typingUndo ?? "block",
      listDefaults: options.listDefaults?.map((l) => ({ ...l })),
      headlineCaseLists: options.headlineCaseLists,
      showFormattingMarks: options.showFormattingMarks,
      commentInitials: options.commentInitials ?? "",
      onEdit: options.onEdit,
      onEditFailed: options.onEditFailed,
      onMove: options.onMove,
      onStoryChange: options.onStoryChange,
      onCommentsChange: options.onCommentsChange,
      onFormatPainterChange: options.onFormatPainterChange,
      onFormattingMarksChange: options.onFormattingMarksChange,
      onPictureChange: options.onPictureChange,
    };
  }

  private static openSession(
    container: HTMLElement,
    bytes: Uint8Array,
    exports: DocxEditorExports,
    opts: DocxEditor["options"],
  ): DocxEditor {
    // NOT persistAnchorIds: that setting applies to every Save on the session, so it put the
    // projector's Unid bookkeeping into the bytes the USER downloads — ~6x the file size for
    // attributes no renderer reads. Only the remount's re-render needs id stability across a
    // save/re-render hop, and it asks for that per call via SaveWithAnchorIds.
    // emitMarkdownPatch off: the editor re-renders from HTML, never from markdown patches,
    // so paying a whole-document re-projection per op would be dead weight.
    requireRevisionIdentitySupport(exports.DocxSessionBridge, opts.revisionIdentities);
    const handle = exports.DocxSessionBridge.OpenSession(bytes, JSON.stringify({
      emitMarkdownPatch: false,
      trackedChanges: trackedChangeWireName(opts.trackedChanges),
      revisionAuthor: opts.revisionAuthor,
      revisionIdentities: opts.revisionIdentities,
    }));
    return new DocxEditor(container, exports, handle, opts);
  }

  /** The steps every mount ends with once the body is in the DOM. */
  private finishMount(): void {
    this.applyFormattingMarks();
    this.syncRegionToBody();
    this.setupBlockDrag();
    if (this.options.comments) this.createGutter();
  }

  /**
   * The windowed mount behind {@link openAsync}: chrome first, then body units in plan order,
   * a group-aligned window per task. See {@link openAsync} for the contract.
   */
  private async mountWindowed(
    windowSize: number,
    onProgress?: (mounted: number, total: number) => void,
  ): Promise<void> {
    const bridge = this.exports.DocxSessionBridge;
    // The full render stamps source anchor ids on every unit; the windows must too, or the DOM
    // would differ from open()'s by exactly that attribute.
    const profile = JSON.stringify({ ...JSON.parse(this.editorRenderProfile()), stampAnchors: true });
    const chrome = bridge.RenderEditorChromeHtml!(this.handle, profile);
    if (chrome.trimStart().startsWith("{")) {
      throw new Error((JSON.parse(chrome) as { error?: string }).error ?? "chrome render failed");
    }
    const parsed = new DOMParser().parseFromString(chrome, "text/html");
    // The chrome renders an empty carrier paragraph where each section break and note reference
    // was; those are not units and go before anything lands. Everything else inside a section
    // wrapper is chrome that stays — page view puts the endnotes section in the last one.
    parsed.body.querySelectorAll<HTMLElement>(`[data-section-index] > [data-anchor="${CHROME_CARRIER_ANCHOR}"]`)
      .forEach((el) => el.remove());
    const styles = Array.from(parsed.querySelectorAll("style")).map((s) => s.outerHTML).join("");
    const flow = document.createElement("div");
    flow.className = "docx-body-flow";
    flow.innerHTML = parsed.body.innerHTML;
    const hosts = new Map<number, HTMLElement>();
    flow.querySelectorAll<HTMLElement>("[data-section-index]").forEach((el) =>
      hosts.set(Number(el.getAttribute("data-section-index")), el));
    const attachNow = !this.options.paginated;
    // A flow mount wires blocks as they land — the chrome's own (footnote and endnote
    // paragraphs) now, each window's as it arrives — the way open() wires everything under
    // the flow. A paginated mount hands the paginator unwired HTML and wires the page clones
    // afterwards, exactly as open() does; wiring first would leave the paginator's
    // contenteditable="false" stamp on every band paragraph.
    const wireNow = attachNow && this.options.editable;
    if (wireNow) this.wireBlocks(flow);
    if (attachNow) {
      this.region?.detachPages();
      this.container.innerHTML = styles;
      this.readoptGutter();
      this.container.appendChild(flow);
      this.editRoot = flow;
    }

    const planJson = this.renderPlanJson();
    // The session version the plan was read at: if nothing has committed by the end of the
    // mount, stampPlanState reuses this plan instead of asking the engine for it again (a
    // second whole-document plan, ~300 ms under WASM on a 9,500-paragraph document).
    const planVersion = bridge.GetVersion?.(this.handle) ?? null;
    const plan = JSON.parse(planJson) as RenderPlan & { error?: string };
    if (plan.error) throw new Error(plan.error);
    const units = plan.body;
    const sectionOf = new Map(units.map((u) => [unidOf(u.id), u.section ?? 0]));
    const mounted = new Set<string>();
    const place = (el: HTMLElement, sameHostAs?: HTMLElement): HTMLElement => {
      // Per-block converter output carries the XHTML xmlns; a full render only has it on the
      // document root, and open()'s DOM is the contract.
      el.removeAttribute("xmlns");
      const unit = el.hasAttribute("data-anchor") ? el : el.querySelector<HTMLElement>("[data-anchor]");
      const unid = unit?.getAttribute("data-anchor") ?? "";
      const host = sameHostAs ?? hosts.get(sectionOf.get(unid) ?? 0) ?? hosts.values().next().value ?? flow;
      // Units precede the chrome a section wrapper already holds (page view's endnotes section).
      host.insertBefore(el, host.querySelector(":scope > section.endnotes"));
      if (wireNow) {
        // wireBlocks wires a root's descendants; the node itself may be the unit.
        if (el.hasAttribute("data-anchor")) this.wireBlock(el);
        this.wireBlocks(el);
      }
      unit && mounted.add(unid);
      el.querySelectorAll("[data-anchor]").forEach((n) => mounted.add(n.getAttribute("data-anchor")!));
      return host;
    };
    for (let start = 0; start < units.length;) {
      let end = Math.min(start + windowSize, units.length);
      // Never cut through a border box: extend to the end of the group the window stopped in.
      while (end < units.length && units[end].group !== undefined && units[end].group === units[end - 1].group) end++;
      const window = units.slice(start, end);
      const rendered = JSON.parse(bridge.RenderEditorRangeHtml!(
        this.handle, JSON.stringify(window.map((u) => u.id)), profile,
      )) as string[] | { error?: string };
      if (!Array.isArray(rendered)) throw new Error(rendered.error ?? "window render failed");
      for (const html of rendered) {
        // One rendered node can parse into several elements: the HTML parser closes a paragraph
        // at a block-level child (page view's page-break marker div), exactly as it does for the
        // full render's string, so every element it produced lands in order in the unit's section.
        let host: HTMLElement | undefined;
        for (const el of Array.from(new DOMParser().parseFromString(html, "text/html").body.children)) {
          host = place(el as HTMLElement, host);
        }
      }
      // A unit the range render could not place (a box that would have straddled the window)
      // still lands, on its own, through the per-block renderer.
      for (const unit of window) {
        if (mounted.has(unidOf(unit.id))) continue;
        const el = this.renderInto(unit.id);
        if (el) place(el);
      }
      start = end;
      onProgress?.(start, units.length);
      if (start < units.length) await yieldToEventLoop();
    }

    if (attachNow) {
      const unchanged = planVersion !== null && bridge.GetVersion?.(this.handle) === planVersion;
      this.stampPlanState(unchanged ? planJson : undefined);
      if (this.region) this.dockBands(flow);
      this.viewport.attach(flow, true);
    } else {
      // Pagination is a one-shot flow over the whole document: hand it the assembled document
      // the way open() hands it the full render — head content (the meta tags, the title and
      // the stylesheet, which the paginator keeps in the flow) followed by the body.
      this.mountPaginated(parsed.head.innerHTML + flow.innerHTML);
    }
  }

  /**
   * Open a fresh, blank document (a "New document" — single empty paragraph, Normal style,
   * US-Letter section) and wire up editing. The seed bytes come from the WASM bridge so the
   * result opens cleanly in Word too.
   */
  static openBlank(
    container: HTMLElement,
    exports: DocxEditorExports,
    options: DocxEditorOptions = {},
  ): DocxEditor {
    return DocxEditor.open(container, exports.DocxSessionBridge.CreateBlankDocx(), exports, options);
  }

  /** Lossless DOCX bytes reflecting all edits, the typing still in the focused block included (`commitTyping`). */
  save(): Uint8Array {
    this.assertOpen();
    this.commitTyping();
    return this.exports.DocxSessionBridge.Save(this.handle);
  }

  /**
   * Commit the typing still in the focused block, as an undo step of its own, keeping the caret
   * where it is. A block commits its typing when it loses focus (with `typingUndo: "word"` also
   * as a word ends or typing pauses), so a save made while the caret stays in the block (Ctrl+S,
   * a host's save on close or an autosave) would otherwise leave out what was typed last.
   * Nothing happens when no block has focus or nothing was typed.
   */
  commitTyping(): void {
    if (this.closed) return;
    if (this.typingPause !== null) clearTimeout(this.typingPause);
    this.typingPause = null;
    const el = this.activeBlock;
    if (!el || !el.isConnected || !el.contains(el.ownerDocument.activeElement)) return;
    this.commitTypingStep(el);
  }

  /** Monotonic committed document version, when supported by the loaded engine. */
  get version(): number | null {
    this.assertOpen();
    const value = this.exports.DocxSessionBridge.GetVersion?.(this.handle);
    return value ? (JSON.parse(value) as { version: number }).version : null;
  }

  /** Release the underlying WASM session. The editor is unusable afterward. */
  close(): void {
    if (this.closed) return;
    this.setPainter("off");
    this.closed = true;
    if (this.typingPause !== null) clearTimeout(this.typingPause);
    this.unwatchLineNumbers();
    this.container.removeEventListener("copy", this.onCopy);
    this.container.removeEventListener("cut", this.onCut);
    this.container.removeEventListener("paste", this.onPaste);
    if (typeof document !== "undefined") {
      document.removeEventListener("selectionchange", this.onSelectionChange);
      document.removeEventListener("mousedown", this.onMouseDown, true);
      document.removeEventListener("mousemove", this.onMouseMove, true);
      document.removeEventListener("mouseup", this.onMouseUp, true);
      document.removeEventListener("keydown", this.onPictureKey, true);
    }
    this.pictureOverlay?.destroy();
    this.pictureOverlay = null;
    this.selectedPictureId = null;
    this.clearDragSelection();
    this.clearFindMatches();
    this.teardownBlockDrag();
    this.gutter?.dispose();
    this.gutter = null;
    this.region?.detachPages();
    this.viewport.dispose();
    this.exports.DocxSessionBridge.CloseSession(this.handle);
  }

  /**
   * Switch between continuous and paginated rendering WITHOUT losing edits. Re-renders from the
   * LIVE session, so every committed edit (and the undo/redo history) survives the toggle — unlike
   * re-opening the original bytes, which silently discards session edits. No-op if already `value`.
   */
  setPaginated(value: boolean): void {
    this.assertOpen();
    if (this.options.paginated === value) return;
    // The view renders again from the session: the typing still in the focused block goes in
    // first, or it is lost (the View tab's buttons keep the focus in the document).
    this.commitTyping();
    this.options.paginated = value;
    this.remount();
  }

  /** The editor's current DOM (for inspection/tests). */
  get root(): HTMLElement {
    return this.container;
  }

  /**
   * Move a fully rendered candidate into the host's stable public surface.
   * Internal editor chrome that binds directly to the container is recreated there;
   * document blocks and their listeners move with the DOM nodes.
   */
  adoptContainer(container: HTMLElement): void {
    this.assertOpen();
    if (container === this.container) return;
    this.teardownBlockDrag();
    this.gutter?.dispose();
    this.gutter = null;
    const previous = this.container;
    previous.removeEventListener("copy", this.onCopy);
    previous.removeEventListener("cut", this.onCut);
    previous.removeEventListener("paste", this.onPaste);
    container.replaceChildren(...Array.from(previous.childNodes));
    this.container = container;
    container.addEventListener("copy", this.onCopy);
    container.addEventListener("cut", this.onCut);
    container.addEventListener("paste", this.onPaste);
    this.viewport.adoptHost(container);
    previous.classList.remove("docx-show-marks");
    this.applyFormattingMarks();
    this.setupBlockDrag();
    if (this.options.comments) this.createGutter();
  }

  /**
   * The zoom the viewport is currently applying (1 = 100%). Below 1 the page is wider than the
   * host and has been scaled to fit rather than reflowed — the honest thing to show a user who
   * is wondering why a phone shows the whole page.
   */
  get zoom(): number {
    return this.viewport.scale;
  }

  /**
   * The live `DocxSession` handle backing this editor — the model of record.
   *
   * Surfaced because chrome around the editor (the anchor rail) reports it as engine
   * state, and reaching into the private field from a host page only worked because
   * the bundle erases TypeScript's visibility.
   */
  get sessionHandle(): number {
    return this.handle;
  }

  /**
   * Repaint from the live session after it was mutated OUTSIDE the editor's own
   * commands — a host driving {@link sessionHandle} directly (an agent pipeline,
   * `raw.replaceXml`, a batch import). Continuous mode patches incrementally from
   * the render plan (a Unid-preserving single-block mutation repaints just that
   * block); paginated mode — or anything the reconciler cannot prove — remounts.
   * The editor cannot observe external mutations, so the host owns calling this
   * once per mutation batch.
   */
  refresh(): void {
    this.assertOpen();
    this.reconcile();
  }

  /** Move one top-level body block relative to another and repaint from the live session. */
  moveBlock(
    sourceAnchorId: string,
    targetAnchorId: string,
    position: "before" | "after",
  ): boolean {
    this.assertOpen();
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.MoveBlock !== "function") return false;
    const res = this.parseEdit(
      bridge.MoveBlock(this.handle, sourceAnchorId, targetAnchorId, position),
    );
    if (!res.success) {
      // The engine's message names an OOXML construct ("a section-break paragraph"); the live
      // region is read aloud, so announce the outcome and keep the raw reason for debugging.
      this.lastMoveError = res.error?.message ?? null;
      this.announceBlockMove("This block cannot be moved there.");
      return false;
    }
    if (
      (res.created?.length ?? 0) === 0 &&
      (res.modified?.length ?? 0) === 0 &&
      (res.removed?.length ?? 0) === 0
    ) {
      this.announceBlockMove("The block is already in that position.");
      return true;
    }

    const destination = res.created?.[0] ?? res.modified?.[0];
    // Both modes reconcile. A direct move relocates the existing element, so the reconciler moves
    // the exact DOM node it already has. Review rendering additionally rewrites the SOURCE in
    // place — it becomes the move-from half — and that is visible to the diff because the plan
    // signs every unit with a content hash, so the source diffs as an in-place substitution.
    // This used to remount, which on a real charter is seconds for a one-block change.
    this.reconcile();

    const moved = destination
      ? this.bodyUnitNodes().find((el) => el.getAttribute("data-anchor") === destination.unid)
      : null;
    if (moved) {
      moved.classList.add("docx-block-move-flash");
      moved.addEventListener("animationend", () => moved.classList.remove("docx-block-move-flash"), {
        once: true,
      });
      const focusTarget = EDITABLE_TAGS.has(moved.tagName)
        ? moved
        : moved.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]');
      if (focusTarget) {
        this.activeBlock = focusTarget;
        focusTarget.focus({ preventScroll: true });
      }
    }
    this.options.onMove?.({
      sourceAnchorId,
      destinationAnchorId: destination?.id ?? sourceAnchorId,
    });
    this.announceBlockMove("Block moved.");
    return true;
  }

  /** Resolve any rendered descendant to its top-level body move unit (tables stay whole). */
  private blockUnitOf(target: EventTarget | Node | null): HTMLElement | null {
    const node = target instanceof Node ? target : null;
    const el = node?.nodeType === Node.ELEMENT_NODE
      ? node as HTMLElement
      : node?.parentElement;
    if (!el || !this.editRoot.contains(el)) return null;
    // Climb rather than scan: this runs on every pointer move over the document, and listing
    // every anchored node to find the one containing the pointer is linear in the document.
    let unit: HTMLElement | null = null;
    for (
      let candidate = el.closest<HTMLElement>("[data-anchor]");
      candidate && this.editRoot.contains(candidate);
      candidate = candidate.parentElement?.closest<HTMLElement>("[data-anchor]") ?? null
    ) {
      unit = candidate;
    }
    if (!unit || unit.closest("section.footnotes, section.endnotes")) return null;
    return unit;
  }

  private isMovableBlockUnit(unit: HTMLElement | null): unit is HTMLElement {
    if (!unit || !this.anchorIdOf(unit)) return false;
    // A source representation of an existing tracked move/deletion is evidence, not a new
    // editable block. Its accepted counterpart remains draggable.
    if (unit.matches("[class$='row-del'], [class*='row-del ']") ||
        unit.querySelector("tr[class$='row-del'], tr[class*='row-del '], del[class$='move-from'], del[class*='move-from ']")
    ) return false;
    return unit.tagName === "TABLE" || EDITABLE_TAGS.has(unit.tagName);
  }

  private currentBlockDragSource(): HTMLElement | null {
    if (this.isMovableBlockUnit(this.blockDragSource) && this.blockDragSource.isConnected)
      return this.blockDragSource;
    const fromActive = this.blockUnitOf(this.activeBlock);
    if (this.isMovableBlockUnit(fromActive)) {
      this.blockDragSource = fromActive;
      return fromActive;
    }
    return null;
  }

  /**
   * The rectangle the floating handle may occupy: the window, narrowed by every ancestor that
   * clips its overflow.
   *
   * The handle is `position: fixed`, so it is placed in viewport coordinates and no ancestor's
   * overflow clips it for us. But the editor is routinely mounted inside a bounded scroller —
   * the ribbon puts its surface in one, inside a clipped card — and a block scrolled out of that
   * scroller has to take its handle with it rather than leave it parked over the host's chrome.
   *
   * Gated on the computed `overflow` rather than on whether an element scrolls right now, so the
   * answer does not depend on how much content happens to be loaded.
   */
  private blockHandleClipRect(): { top: number; right: number; bottom: number; left: number } {
    const view = this.container.ownerDocument.defaultView;
    const clip = {
      top: 0,
      left: 0,
      right: view?.innerWidth ?? Number.MAX_SAFE_INTEGER,
      bottom: view?.innerHeight ?? Number.MAX_SAFE_INTEGER,
    };
    for (let el: HTMLElement | null = this.editRoot; el; el = el.parentElement) {
      const style = view?.getComputedStyle(el);
      if (!style || (style.overflowX === "visible" && style.overflowY === "visible")) continue;
      const box = el.getBoundingClientRect();
      if (style.overflowY !== "visible") {
        clip.top = Math.max(clip.top, box.top);
        clip.bottom = Math.min(clip.bottom, box.bottom);
      }
      if (style.overflowX !== "visible") {
        clip.left = Math.max(clip.left, box.left);
        clip.right = Math.min(clip.right, box.right);
      }
    }
    return clip;
  }

  private showBlockHandle(unit: HTMLElement): void {
    const handle = this.blockDragHandle;
    if (!handle || !this.isMovableBlockUnit(unit)) return;
    const changed = unit !== this.blockDragSource;
    this.blockDragSource = unit;
    this.blockHandleWanted = true;
    // A block the engine will not move anywhere — one owning a section break, or already carrying
    // revision markup a tracked move would have to re-wrap — gets no handle rather than a handle
    // that always fails. Asking costs an engine round trip, so hovering only ever CONSUMES a
    // cached answer and schedules the ask for idle time; the drag start and the menu ask for
    // real. Merely moving the pointer across the document must not run document-scale work.
    const known = this.blockMoveTargetsFor(unit, { cachedOnly: true });
    if (known?.size === 0) {
      handle.style.display = "none";
      return;
    }
    if (changed && known === undefined) this.prefetchBlockMoveTargets(unit);
    const rect = unit.getBoundingClientRect();
    // The handle belongs to the editor's viewport, not the window's. A block scrolled out of that
    // viewport gets no handle at all: clamping one back into view would leave a grip sitting next
    // to a block that is not there — and clamping to the WINDOW, as this used to, parked it on
    // whatever chrome the host draws above the editor.
    const clip = this.blockHandleClipRect();
    if (
      rect.bottom <= clip.top || rect.top >= clip.bottom
      || rect.right <= clip.left || rect.left >= clip.right
    ) {
      // A gesture in flight owns the handle: it is the drag's own source element, and an open
      // move menu is anchored to it and hands focus back to it on Escape. Withdrawing it
      // mid-gesture is the same mistake as leaving it behind — mirror `hideBlockHandle`'s guard
      // and leave it where the gesture put it.
      if (!this.blockDragging && this.blockMoveMenu?.style.display !== "block")
        handle.style.display = "none";
      return;
    }
    handle.style.display = "flex";
    // A block only partly in view keeps its handle, pulled to the edge it is disappearing past.
    const top = rect.top
      + (unit.tagName === "TABLE" ? 6 : Math.max(0, (rect.height - BLOCK_HANDLE_HEIGHT) / 2));
    handle.style.left = `${clampToRange(
      rect.left - (BLOCK_HANDLE_WIDTH + BLOCK_HANDLE_GAP),
      clip.left,
      clip.right - BLOCK_HANDLE_WIDTH,
    )}px`;
    handle.style.top = `${clampToRange(top, clip.top, clip.bottom - BLOCK_HANDLE_HEIGHT)}px`;
    const preview = blockPreviewText(unit);
    handle.setAttribute("aria-label", preview ? `Move block: ${preview}` : "Move block");
  }

  private hideBlockHandle(): void {
    if (this.blockDragging || this.blockMoveMenu?.style.display === "block") return;
    this.blockHandleWanted = false;
    if (this.blockDragHandle) this.blockDragHandle.style.display = "none";
  }

  private positionBlockHandle(): void {
    const source = this.currentBlockDragSource();
    if (source && this.blockHandleWanted) this.showBlockHandle(source);
  }

  /** Draw the drop line on `zone`'s requested edge, or take it away when there is no target. */
  private paintDropIndicator(data: Record<string | symbol, unknown>): void {
    const indicator = this.blockDropIndicator;
    const zone = data.zone as BlockDropZone | undefined;
    if (!indicator) return;
    if (!zone) {
      this.hideDropIndicator();
      return;
    }
    const y = Math.round(this.dropEdgeY(zone, data.position === "after" ? "after" : "before")
      + this.dropZoneShift()) - 1;
    // Position before revealing, so the entry fade never plays at a stale spot.
    indicator.style.transform = `translate3d(${Math.round(zone.left)}px, ${y}px, 0)`;
    indicator.style.width = `${Math.max(24, Math.round(zone.width))}px`;
    indicator.style.display = "block";
  }

  /**
   * Where to draw the line for an insertion on `position` of `zone` — the MIDDLE of the gap to
   * the neighbour on that side, not the zone's own border-box edge. A paragraph's `w:spacing`
   * becomes a CSS margin, which sits outside the box, so drawing on the edge underlines the
   * block's last line instead of reading as a gap between two blocks. Falls back to the raw edge
   * at the ends of the flow, and degrades to the same value when blocks are contiguous.
   */
  private dropEdgeY(zone: BlockDropZone, position: "before" | "after"): number {
    const neighbour = this.dropZones[zone.index + (position === "after" ? 1 : -1)];
    // At the ends of the flow there is nothing to bisect against, so half the block's own
    // margin stands in for half the gap — the same position, one contributor instead of two.
    if (!neighbour) {
      return position === "after"
        ? zone.bottom + zone.marginAfter / 2
        : zone.top - zone.marginBefore / 2;
    }
    return position === "after"
      ? (zone.bottom + neighbour.top) / 2
      : (neighbour.bottom + zone.top) / 2;
  }

  private hideDropIndicator(): void {
    if (this.blockDropIndicator) this.blockDropIndicator.style.display = "none";
  }

  private announceBlockMove(message: string): void {
    const live = this.blockMoveLive;
    if (!live) return;
    live.textContent = "";
    this.container.ownerDocument.defaultView?.setTimeout(() => { live.textContent = message; }, 0);
  }

  /**
   * Ask the engine which anchors the current source may move next to. One call per drag source,
   * so the UI offers only drops MoveBlock will accept — a document with section breaks is
   * partitioned into regions, and drawing an indicator across one only to fail the drop was the
   * behaviour this replaces. Null (no bridge support) keeps the previous offer-everything path.
   */
  private refreshBlockMoveTargets(source: HTMLElement | null): void {
    this.blockMoveTargets = source ? this.blockMoveTargetsFor(source) ?? null : null;
  }

  /**
   * This block's legal destinations, from the memo when it is there. Returns `undefined` — not
   * `null` — for "not asked yet", so a caller can tell an unknown answer from the engine's
   * "no bridge support, offer everything" one.
   */
  private blockMoveTargetsFor(
    source: HTMLElement,
    options: { cachedOnly?: boolean } = {},
  ): Map<string, { before: boolean; after: boolean }> | null | undefined {
    const bridge = this.exports.DocxSessionBridge;
    const sourceId = this.anchorIdOf(source);
    if (!sourceId || typeof bridge.ValidMoveTargets !== "function") return null;
    if (this.blockMoveTargetCache.has(sourceId)) return this.blockMoveTargetCache.get(sourceId);
    if (options.cachedOnly) return undefined;

    let targets: Map<string, { before: boolean; after: boolean }> | null;
    try {
      const parsed = JSON.parse(bridge.ValidMoveTargets(this.handle, sourceId)) as Array<
        { anchorId: string; before: boolean; after: boolean }
      >;
      targets = new Map(parsed.map((t) => [t.anchorId, { before: t.before, after: t.after }]));
    } catch {
      targets = null;
    }
    this.blockMoveTargetCache.set(sourceId, targets);
    return targets;
  }

  /**
   * Ask for the hovered block's destinations off the interaction path, and hide the handle if
   * the answer comes back empty and that block is still the one under the pointer. The handle
   * therefore appears immediately on hover and withdraws a beat later on the rare immovable
   * block, instead of every hover paying for the query up front.
   */
  private prefetchBlockMoveTargets(source: HTMLElement): void {
    const view = this.container.ownerDocument.defaultView;
    if (!view) return;
    this.blockMoveTargetPrefetch?.();
    const run = (): void => {
      this.blockMoveTargetPrefetch = null;
      // The pointer has moved on (or the DOM was repainted) — the answer is no longer wanted.
      if (this.closed || !source.isConnected || this.blockDragSource !== source) return;
      if (this.blockMoveTargetsFor(source)?.size === 0 && this.blockDragHandle)
        this.blockDragHandle.style.display = "none";
    };
    const idle = (view as unknown as {
      requestIdleCallback?: (cb: () => void, opts?: { timeout: number }) => number;
      cancelIdleCallback?: (id: number) => void;
    });
    if (idle.requestIdleCallback && idle.cancelIdleCallback) {
      const id = idle.requestIdleCallback(run, { timeout: 500 });
      this.blockMoveTargetPrefetch = () => idle.cancelIdleCallback!(id);
    } else {
      const id = view.setTimeout(run, 0);
      this.blockMoveTargetPrefetch = () => view.clearTimeout(id);
    }
  }

  /**
   * Whether `unit` is a legal destination for the current drag source — for a SPECIFIC side when
   * one is given. A cross-block range or a section break between the blocks can make one side
   * legal and the other not, so "this target is reachable" is not enough to pick a position.
   */
  private isValidMoveTarget(unit: HTMLElement, position?: "before" | "after"): boolean {
    if (!this.blockMoveTargets) return true; // no bridge support: engine stays the only gate
    const id = this.anchorIdOf(unit);
    const sides = id ? this.blockMoveTargets.get(id) : undefined;
    if (!sides) return false;
    return position ? sides[position] : sides.before || sides.after;
  }

  /** Measure every movable block once, at drag start. See `BlockDropZone`. */
  private captureDropZones(): void {
    const view = this.container.ownerDocument.defaultView;
    this.dropZoneScroller = this.scrollContainer();
    this.dropZoneOrigin = this.scrollOffsetSum();
    this.dropZones = [];
    const boxes: HTMLElement[] = [];
    for (const unit of this.bodyUnitNodes()) {
      const anchorId = this.isMovableBlockUnit(unit) ? this.anchorIdOf(unit) : null;
      if (!anchorId) continue;
      const box = this.unitWrapperOf(unit);
      const rect = box.getBoundingClientRect();
      boxes.push(box);
      this.dropZones.push({
        unit, anchorId, index: this.dropZones.length,
        top: rect.top, bottom: rect.bottom, left: rect.left, width: rect.width,
        marginBefore: 0, marginAfter: 0,
      });
    }
    // Only the flow's two ends ever consult a margin, so only they are worth a style read.
    const ends = new Set([0, this.dropZones.length - 1].filter((i) => i >= 0 && i < boxes.length));
    for (const i of ends) {
      const style = view?.getComputedStyle(boxes[i]);
      this.dropZones[i].marginBefore = parseFloat(style?.marginTop ?? "0") || 0;
      this.dropZones[i].marginAfter = parseFloat(style?.marginBottom ?? "0") || 0;
    }
  }

  private scrollOffsetSum(): number {
    const view = this.container.ownerDocument.defaultView;
    return (view?.scrollY ?? 0) + (this.dropZoneScroller?.scrollTop ?? 0);
  }

  /** How far the measured boxes have travelled since capture, from scrolling (drag autoscroll). */
  private dropZoneShift(): number {
    return this.dropZoneOrigin - this.scrollOffsetSum();
  }

  /**
   * Where a drop at `clientY` lands, or null when nothing there is legal.
   *
   * Resolution is by VERTICAL GEOMETRY over the measured blocks, not by which element the pointer
   * is over: the drag handle floats in the page margin, so a drag straight down the gutter — the
   * natural gesture — never crosses a paragraph box, and element hit testing gave those drags no
   * indicator and no drop at all. The nearest block by vertical distance is the target; the half
   * the pointer is in picks the side, snapped to the other side when only that one is legal
   * (a section break or a cross-block range usually makes exactly one side illegal). When neither
   * side is legal — the pointer is in a region this block cannot reach — there is no drop, and
   * nothing is drawn.
   */
  private resolveDropAt(clientY: number): { zone: BlockDropZone; position: "before" | "after" } | null {
    const y = clientY - this.dropZoneShift();
    let best: BlockDropZone | null = null;
    let bestGap = Infinity;
    for (const zone of this.dropZones) {
      const gap = y < zone.top ? zone.top - y : y > zone.bottom ? y - zone.bottom : 0;
      if (gap < bestGap) {
        best = zone;
        bestGap = gap;
      }
      if (gap === 0) break; // inside this block; zones are in document order and never overlap
    }
    if (!best || best.unit === this.blockDragSource) return null;
    const preferred = y < (best.top + best.bottom) / 2 ? "before" : "after";
    if (this.isValidMoveTarget(best.unit, preferred)) return { zone: best, position: preferred };
    const other = preferred === "before" ? "after" : "before";
    return this.isValidMoveTarget(best.unit, other) ? { zone: best, position: other } : null;
  }

  private closeBlockMoveMenu(restoreFocus = false): void {
    if (!this.blockMoveMenu || !this.blockDragHandle) return;
    this.blockMoveMenu.style.display = "none";
    this.blockDragHandle.setAttribute("aria-expanded", "false");
    this.blockDragHandle.setAttribute("aria-pressed", "false");
    if (restoreFocus) this.blockDragHandle.focus({ preventScroll: true });
  }

  private openBlockMoveMenu(): void {
    const menu = this.blockMoveMenu;
    const handle = this.blockDragHandle;
    const source = this.currentBlockDragSource();
    if (!menu || !handle || !source) return;
    this.refreshBlockMoveTargets(source);
    menu.querySelectorAll<HTMLButtonElement>("button[data-move]").forEach((button) => {
      button.disabled = !this.blockMoveDestination(source, button.dataset.move ?? "");
    });
    const rect = handle.getBoundingClientRect();
    menu.style.display = "block";
    menu.style.left = `${Math.max(4, rect.right + 6)}px`;
    menu.style.top = `${Math.max(4, rect.top)}px`;
    handle.setAttribute("aria-expanded", "true");
    handle.setAttribute("aria-pressed", "true");
    menu.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus({ preventScroll: true });
  }

  /**
   * Resolve a move-menu action to the block it should land against, considering only destinations
   * the engine accepts. "Top"/"bottom" therefore mean the ends of the source's own movable region
   * — on a document with section breaks that is the section, not the document, which is what makes
   * the commands work at all there instead of always failing.
   */
  private blockMoveDestination(
    source: HTMLElement,
    action: string,
  ): { target: HTMLElement; position: "before" | "after" } | null {
    const units = this.bodyUnitNodes().filter((el) => this.isMovableBlockUnit(el));
    const index = units.indexOf(source);
    if (index < 0) return null;

    // Only pairs the engine accepts: a target can be reachable on one side and refused on the
    // other, so the SIDE is part of what makes a candidate valid.
    const position: "before" | "after" = action === "up" || action === "top" ? "before" : "after";
    const side = position === "before" ? units.slice(0, index) : units.slice(index + 1);
    const candidates = side.filter((el) => this.isValidMoveTarget(el, position));
    if (candidates.length === 0) return null;

    if (action === "up") return { target: candidates[candidates.length - 1], position };
    if (action === "down") return { target: candidates[0], position };
    if (action === "top") return { target: candidates[0], position };
    if (action === "bottom") return { target: candidates[candidates.length - 1], position };
    return null;
  }

  private runBlockMoveMenuAction(action: string): void {
    const source = this.currentBlockDragSource();
    if (!source) return;
    const destination = this.blockMoveDestination(source, action);
    this.closeBlockMoveMenu();
    if (!destination) return;
    const sourceId = this.anchorIdOf(source);
    const targetId = this.anchorIdOf(destination.target);
    if (sourceId && targetId) this.moveBlock(sourceId, targetId, destination.position);
  }

  /** The nearest ancestor that actually scrolls the document flow, for drag autoscroll. */
  private scrollContainer(): HTMLElement | null {
    const view = this.container.ownerDocument.defaultView;
    for (let el: HTMLElement | null = this.editRoot; el; el = el.parentElement) {
      const overflowY = view?.getComputedStyle(el).overflowY ?? "visible";
      if ((overflowY === "auto" || overflowY === "scroll") && el.scrollHeight > el.clientHeight)
        return el;
    }
    return null;
  }

  /** Mount one floating handle; PDD owns pointer dragging while the menu owns keyboard moves. */
  private setupBlockDrag(): void {
    if (!this.options.blockDrag || this.options.paginated || this.closed || this.blockDragHandle)
      return;
    const doc = this.container.ownerDocument;
    ensureBlockDragStyles(doc);

    const handle = doc.createElement("div");
    handle.className = "docx-block-handle";
    handle.textContent = "⠿";
    handle.tabIndex = 0;
    handle.setAttribute("role", "button");
    handle.setAttribute("contenteditable", "false");
    handle.setAttribute("aria-haspopup", "menu");
    handle.setAttribute("aria-expanded", "false");
    handle.setAttribute("aria-pressed", "false");

    const indicator = doc.createElement("div");
    indicator.className = "docx-block-drop-indicator";
    const menu = doc.createElement("div");
    menu.className = "docx-block-move-menu";
    menu.setAttribute("role", "menu");
    menu.innerHTML = `
      <button type="button" role="menuitem" data-move="up">Move up</button>
      <button type="button" role="menuitem" data-move="down">Move down</button>
      <button type="button" role="menuitem" data-move="top">Move to top</button>
      <button type="button" role="menuitem" data-move="bottom">Move to bottom</button>`;
    const live = doc.createElement("div");
    live.className = "docx-block-live-region";
    live.setAttribute("role", "status");
    live.setAttribute("aria-live", "polite");
    this.container.append(handle, indicator, menu, live);
    this.blockDragHandle = handle;
    this.blockDropIndicator = indicator;
    this.blockMoveMenu = menu;
    this.blockMoveLive = live;

    const onPointerMove = (event: PointerEvent): void => {
      // Keep the floating draggable stationary between pointerdown and native dragstart.
      // Repositioning it under the pointer during the browser's drag threshold cancels
      // drag initiation before PDD can receive dragstart.
      if (this.blockDragging || this.blockDragPointerDown) return;
      const unit = this.blockUnitOf(event.target);
      if (this.isMovableBlockUnit(unit)) this.showBlockHandle(unit);
    };
    const onFocusIn = (event: FocusEvent): void => {
      const unit = this.blockUnitOf(event.target);
      if (this.isMovableBlockUnit(unit)) this.showBlockHandle(unit);
    };
    const onPointerLeave = (event: PointerEvent): void => {
      const next = event.relatedTarget instanceof Node ? event.relatedTarget : null;
      if (next && (handle.contains(next) || menu.contains(next))) return;
      this.hideBlockHandle();
    };
    const onViewportChange = (): void => this.positionBlockHandle();
    this.container.addEventListener("pointermove", onPointerMove);
    this.container.addEventListener("focusin", onFocusIn);
    this.container.addEventListener("pointerleave", onPointerLeave);
    doc.addEventListener("scroll", onViewportChange, true);
    doc.defaultView?.addEventListener("resize", onViewportChange);
    this.blockDragCleanup.push(
      () => this.container.removeEventListener("pointermove", onPointerMove),
      () => this.container.removeEventListener("focusin", onFocusIn),
      () => this.container.removeEventListener("pointerleave", onPointerLeave),
      () => doc.removeEventListener("scroll", onViewportChange, true),
      () => doc.defaultView?.removeEventListener("resize", onViewportChange),
    );

    const onHandleClick = (): void => {
      if (menu.style.display === "block") this.closeBlockMoveMenu();
      else this.openBlockMoveMenu();
    };
    const onHandlePointerDown = (): void => { this.blockDragPointerDown = true; };
    const onPointerUp = (): void => { this.blockDragPointerDown = false; };
    const onHandleKeydown = (event: KeyboardEvent): void => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      onHandleClick();
    };
    const onMenuClick = (event: MouseEvent): void => {
      const button = (event.target as Element | null)?.closest<HTMLButtonElement>("button[data-move]");
      if (button && !button.disabled) this.runBlockMoveMenuAction(button.dataset.move ?? "");
    };
    const onMenuKeydown = (event: KeyboardEvent): void => {
      if (event.key === "Escape") { event.preventDefault(); this.closeBlockMoveMenu(true); }
    };
    handle.addEventListener("click", onHandleClick);
    handle.addEventListener("pointerdown", onHandlePointerDown);
    handle.addEventListener("keydown", onHandleKeydown);
    menu.addEventListener("click", onMenuClick);
    menu.addEventListener("keydown", onMenuKeydown);
    doc.addEventListener("pointerup", onPointerUp, true);
    doc.addEventListener("pointercancel", onPointerUp, true);
    this.blockDragCleanup.push(
      () => handle.removeEventListener("click", onHandleClick),
      () => handle.removeEventListener("pointerdown", onHandlePointerDown),
      () => handle.removeEventListener("keydown", onHandleKeydown),
      () => menu.removeEventListener("click", onMenuClick),
      () => menu.removeEventListener("keydown", onMenuKeydown),
      () => doc.removeEventListener("pointerup", onPointerUp, true),
      () => doc.removeEventListener("pointercancel", onPointerUp, true),
    );

    this.blockDragCleanup.push(draggable({
      element: handle,
      canDrag: () => !!this.currentBlockDragSource(),
      getInitialData: () => {
        const source = this.currentBlockDragSource();
        return { type: BLOCK_DRAG_TYPE, sourceAnchorId: source ? this.anchorIdOf(source) : undefined };
      },
      // The browser would otherwise ghost the 26px grip, which says nothing about what is moving.
      onGenerateDragPreview: ({ nativeSetDragImage }) => {
        const source = this.currentBlockDragSource();
        setCustomNativeDragPreview({
          nativeSetDragImage,
          getOffset: pointerOutsideOfPreview({ x: "14px", y: "10px" }),
          render: ({ container }) => {
            const chip = doc.createElement("div");
            chip.className = "docx-block-drag-preview";
            chip.textContent = (source && blockPreviewText(source)) || "Move block";
            container.appendChild(chip);
            return () => chip.remove();
          },
        });
      },
      onDragStart: () => {
        const source = this.currentBlockDragSource();
        this.blockDragging = true;
        handle.classList.add("docx-block-dragging");
        // After the preview snapshot, so the chip is not dimmed too.
        source?.classList.add("docx-block-drag-source");
        this.closeBlockMoveMenu();
        this.refreshBlockMoveTargets(source);
        this.captureDropZones();
      },
      onDrop: () => {
        this.blockDragging = false;
        this.blockDragPointerDown = false;
        this.dropZones = [];
        handle.classList.remove("docx-block-dragging");
        // Cleared by selector rather than from a captured reference: a completed move may have
        // re-rendered the block, and the node that carries the class outlives either handle on it.
        doc.querySelectorAll(".docx-block-drag-source")
          .forEach((el) => el.classList.remove("docx-block-drag-source"));
        this.hideDropIndicator();
      },
    }));
    // ONE drop target for the whole flow. Which block a drop belongs to is decided by
    // `resolveDropAt` from the pointer's vertical position, so the gutter the handle is dragged
    // down counts as being over the document — see that method.
    this.blockDragCleanup.push(dropTargetForElements({
      element: this.editRoot,
      canDrop: ({ source }) => source.data.type === BLOCK_DRAG_TYPE,
      getData: ({ input }) => {
        const hit = this.resolveDropAt(input.clientY);
        return hit
          ? { type: BLOCK_DRAG_TYPE, targetAnchorId: hit.zone.anchorId, position: hit.position, zone: hit.zone }
          : { type: BLOCK_DRAG_TYPE };
      },
      onDragEnter: ({ self }) => this.paintDropIndicator(self.data),
      onDrag: ({ self }) => this.paintDropIndicator(self.data),
      onDragLeave: () => this.hideDropIndicator(),
    }));
    this.blockDragCleanup.push(monitorForElements({
      canMonitor: ({ source }) => source.data.type === BLOCK_DRAG_TYPE,
      onDrop: ({ source, location }) => {
        this.hideDropIndicator();
        const drop = location.current.dropTargets[0]?.data;
        const sourceId = source.data.sourceAnchorId;
        const targetId = drop?.targetAnchorId;
        const position = drop?.position === "after" ? "after" : "before";
        if (typeof sourceId === "string" && typeof targetId === "string")
          this.moveBlock(sourceId, targetId, position);
        else
          // Released outside any valid target. Say so — a silent no-op reads as a broken drag.
          this.announceBlockMove("Move cancelled — the block was not dropped on a valid position.");
      },
    }));
    // Register on the element that actually scrolls. editRoot is the document flow — in the
    // ribbon surface the scroller is an ancestor of it, and registering the flow made Pragmatic
    // log "attached to an element that appears not to be scrollable" on every document open
    // while doing nothing.
    const scroller = this.scrollContainer();
    if (scroller) {
      this.blockDragCleanup.push(autoScrollForElements({
        element: scroller,
        canScroll: ({ source }) => source.data.type === BLOCK_DRAG_TYPE,
        getAllowedAxis: () => "vertical",
      }));
    }
    this.blockDragCleanup.push(autoScrollWindowForElements({
      canScroll: ({ source }) => source.data.type === BLOCK_DRAG_TYPE,
      getAllowedAxis: () => "vertical",
    }));
  }

  private teardownBlockDrag(): void {
    this.blockMoveTargetPrefetch?.();
    this.blockMoveTargetPrefetch = null;
    this.blockMoveTargetCache.clear();
    this.dropZones = [];
    this.dropZoneScroller = null;
    for (const cleanup of this.blockDragCleanup.splice(0)) cleanup();
    this.blockDragHandle?.remove();
    this.blockDropIndicator?.remove();
    this.blockMoveMenu?.remove();
    this.blockMoveLive?.remove();
    this.blockDragHandle = null;
    this.blockDropIndicator = null;
    this.blockMoveMenu = null;
    this.blockMoveLive = null;
    this.blockDragSource = null;
    this.blockHandleWanted = false;
    this.blockDragging = false;
    this.blockDragPointerDown = false;
  }

  // ─── internals ───────────────────────────────────────────────────────

  private assertOpen(): void {
    if (this.closed) throw new Error("DocxEditor is closed");
  }

  /**
   * Rebuild unid → full-anchor-id from the live session projection.
   *
   * Unids are CONTENT-ADDRESSED, so blocks with identical content in DIFFERENT parts collide —
   * e.g. a document with empty default/first/even header stories has one unid for several
   * header parts. A collision must resolve to the BODY entry: body blocks carry only
   * `data-anchor` (the bare unid) and have nothing else to resolve through, whereas a
   * header/footer band block carries its full anchor in `data-hf-anchor` and is resolved from
   * that (see `anchorIdOf`). Letting a non-body scope win here would silently redirect a body
   * edit into a header part.
   */
  private refreshAnchorMap(): void {
    const bridge = this.exports.DocxSessionBridge;
    // ListAnchors returns the same {anchorIndex} object WITHOUT the markdown payload —
    // marshaling the full projection (a couple hundred KB on a real document) made
    // this refresh the single biggest term of every incremental repaint.
    const raw =
      typeof bridge.ListAnchors === "function"
        ? bridge.ListAnchors(this.handle)
        : bridge.Project(this.handle);
    const proj = JSON.parse(raw) as {
      anchorIndex: Record<string, AnchorTargetLite>;
    };
    this.unidToFullId.clear();
    const bodyOwned = new Set<string>();
    for (const [fullId, target] of Object.entries(proj.anchorIndex)) {
      const isBody = target.scope === "body";
      if (!isBody && bodyOwned.has(target.unid)) continue;
      this.unidToFullId.set(target.unid, fullId);
      if (isBody) bodyOwned.add(target.unid);
    }
  }

  /**
   * The full `kind:scope:unid` anchor for a rendered block. A header/footer band block carries
   * its own — the unid map cannot disambiguate one, since several parts' story paragraphs can
   * share a content-addressed unid (a real Word document with empty default/first/even stories
   * does exactly that, and a unid-keyed lookup would land the edit in the wrong header part).
   */
  private anchorIdOf(el: HTMLElement): string | undefined {
    const stamped = el.getAttribute("data-hf-anchor");
    if (stamped) return stamped;
    const unid = el.getAttribute("data-anchor");
    return unid ? this.unidToFullId.get(unid) : undefined;
  }

  /** Build the header/footer region (called once, before the first mount). */
  private createRegion(): void {
    this.region = new HeaderFooterRegion(
      this.exports.DocxSessionBridge,
      this.handle,
      {
        wireBlock: (el) => this.wireBlock(el),
        renderBlock: (anchorId) => this.renderInto(anchorId),
        refreshAnchorMap: () => this.refreshAnchorMap(),
        bodyAnchorIdOf: (unid) => this.unidToFullId.get(unid),
        remount: () => {
          if (!this.closed) this.remount();
        },
        onActiveChange: (_host, which) => {
          this.activeStory = which;
          this.options.onStoryChange?.(which);
        },
        parseEdit: (json) => this.parseEdit(json),
      },
    );
  }

  /** Build the comment gutter over the container (called once, after the first mount). */
  private createGutter(): void {
    ensureCommentGutterStyles(this.container.ownerDocument);
    const options = this.options;
    this.gutter = new CommentGutter(
      {
        container: this.container,
        // Read at each use, so setCommentAuthor reaches the gutter's drafts and replies.
        get commentAuthor() { return options.commentAuthor; },
        listComments: () => this.listComments(),
        addComment: (markdown, author, target) => this.addComment(markdown, author, target),
        addCommentReply: (parent, markdown, author) => this.addCommentReply(parent, markdown, author),
        updateComment: (anchorId, markdown) => this.updateComment(anchorId, markdown),
        removeComment: (anchorId) => this.removeComment(anchorId),
        setCommentResolved: (anchorId, resolved) => this.setCommentResolved(anchorId, resolved),
        commentTarget: () => this.commentTarget(),
      },
      { onChange: (info) => this.options.onCommentsChange?.(info) },
    );
  }

  /**
   * Point the bands at the section governing the current focus (or the first body block).
   * Called after every mount, and on focus of a body block, so a multi-section document shows
   * the stories that actually apply where the caret is.
   */
  private syncRegionToBody(fromBlock?: HTMLElement): void {
    if (!this.region) return;
    const block = fromBlock ?? this.editableList()[0];
    const unid = block?.getAttribute("data-anchor");
    const id = unid ? this.unidToFullId.get(unid) : undefined;
    this.region.syncToBody(id ?? null);
  }

  /**
   * Insert the bands around `bodyRoot` and make `bodyRoot` the edit root. Band blocks must stay
   * OUT of the edit root: `editableList()`/`blockIndex()` enumerate it to compute remount focus
   * indices, and band blocks in that list would shift every index.
   */
  private dockBands(bodyRoot: HTMLElement): void {
    if (!this.region) return;
    bodyRoot.before(this.region.headerBand);
    bodyRoot.after(this.region.footerBand);
    this.region.refreshAll();
  }

  /**
   * Continuous (non-paginated) mount: inject the converter's styles + body, wire blocks.
   *
   * The body always gets its own `.docx-body-flow` wrapper — not only when bands are docked.
   * It is the sheet: the element the viewport gives page geometry to and zooms, and the one
   * the bands dock around. Without it the container would have to be both the scrolling host
   * and the scaled page, which are different boxes.
   */
  private mountHtml(fullHtml: string): void {
    this.unwatchLineNumbers();
    this.region?.detachPages();
    const parsed = new DOMParser().parseFromString(fullHtml, "text/html");
    const styles = Array.from(parsed.querySelectorAll("style"))
      .map((s) => s.outerHTML)
      .join("");
    this.container.innerHTML = styles;
    this.readoptGutter();
    const flow = document.createElement("div");
    flow.className = "docx-body-flow";
    flow.innerHTML = parsed.body.innerHTML;
    this.container.appendChild(flow);
    this.editRoot = flow;
    if (this.options.editable) this.wireBlocks(flow);
    this.stampPlanState();
    if (this.region) this.dockBands(flow);
    this.viewport.attach(flow, true);
  }

  /** Paginated mount: flow blocks into page boxes via pagination.ts, wire the page clones. */
  private mountPaginated(fullHtml: string): void {
    this.region?.detachPages();
    // Pagination writes into its own wrapper so the container can also hold the comment gutter
    // (and so pagination's innerHTML reset can never eat it).
    this.container.innerHTML = "";
    this.readoptGutter();
    const target = document.createElement("div");
    target.className = "docx-body-flow";
    this.container.appendChild(target);
    // Fragmented paragraphs intentionally have only one addressable head and
    // are therefore unsuitable for the editor's one-block editing model.
    paginateHtml(fullHtml, target, {
      scale: this.options.scale,
      cssPrefix: "page-",
      fragmentParagraphs: false,
      // The editor owns a separate comment gutter; keep its physical page substrate unchanged.
      commentColumnPlacement: "margin",
      // Moved glyph boxes would part from text typed into them.
      alignBaselinesToWord: false,
    });
    // pagination.ts measures the hidden #pagination-staging subtree ONCE, then flows CLONES of its
    // blocks into the visible page boxes. Leaving staging in the live DOM is a trap: every
    // data-anchor exists twice (staging + page-box copy), so document.querySelector('[data-anchor]')
    // is ambiguous (hits the hidden copy first), and the staging copy goes stale because edits land
    // only on the page-box copy — a future reflow-from-staging would silently revert them. Staging
    // is a transient measurement scaffold; drop it so the page-box copies are the single source of
    // truth. A remount (setPaginated, list/undo edits) rebuilds staging fresh from the live session.
    this.container.querySelector("#pagination-staging, .page-staging")?.remove();
    const pageRoot = target.querySelector<HTMLElement>("#pagination-container") ?? target;
    this.editRoot = pageRoot;
    if (this.options.editable) this.wireBlocks(pageRoot);
    this.watchLineNumbers(pageRoot);
    this.stampPlanState();
    // The page boxes render their own header/footer margins as inert clones; the region turns
    // the clicked one into the live story in place (Word's edit-in-the-margin), so there is
    // still exactly one addressable node per story paragraph at any time.
    if (this.region) this.region.attachPages(pageRoot);
    // Page boxes already carry the section's dimensions, so the viewport contributes only the
    // fit zoom — a full page is far wider than a phone, and clipping it is not an option.
    this.viewport.attach(pageRoot, false);
  }

  /** Keep the page view's line numbers (drawn by the paginator) in step with edits made in place:
   *  any change to the pages' text redraws them, once, after the typing pauses. */
  private watchLineNumbers(pageRoot: HTMLElement): void {
    this.unwatchLineNumbers();
    this.sectionStarts.clear();
    this.markSectionBreaks();
    if (typeof MutationObserver === "undefined" || (!hasLineNumbers(pageRoot) && !this.hasSectionBreaks())) return;
    const isLayer = (node: Node): boolean =>
      node.nodeType === 1 && (node as Element).classList.contains("docx-line-numbers");
    const watch = {
      timer: 0,
      observer: new MutationObserver((records) => {
        const relevant = records.some((r) =>
          !(r.target.nodeType === 1 ? (r.target as Element) : r.target.parentElement)?.closest(".docx-line-numbers")
          && !(r.type === "childList" && [...Array.from(r.addedNodes), ...Array.from(r.removedNodes)].every(isLayer)));
        if (!relevant) return;
        window.clearTimeout(watch.timer);
        watch.timer = window.setTimeout(() => {
          if (this.closed || !pageRoot.isConnected) return;
          this.markSectionBreaks();
          if (hasLineNumbers(pageRoot)) drawLineNumbers(pageRoot);
        }, 120);
      }),
    };
    watch.observer.observe(pageRoot, { childList: true, subtree: true, characterData: true });
    this.lineNumberWatch = watch;
  }

  /** How each section starts (`w:type`), by section key, read once a mount. */
  private readonly sectionStarts = new Map<string, string>();

  private hasSectionBreaks(): boolean {
    const keys = new Set(this.editableList().filter((b) => !this.isBandBlock(b)).map((b) => this.sectionKeyOf(b)));
    return keys.size > 1;
  }

  /** In page view, mark the last paragraph of each section but the last with how the next section
   *  starts (`data-section-break`): with Show ¶ it shows Word's "Section Break (…)" there. */
  private markSectionBreaks(): void {
    if (!this.options.paginated || this.closed) return;
    const body = this.editableList().filter((b) => !this.isBandBlock(b) && !b.closest("td, th"));
    for (let i = 0; i < body.length; i++) {
      const block = body[i];
      const next = body[i + 1];
      const nextKey = next ? this.sectionKeyOf(next) : null;
      if (!next || nextKey === this.sectionKeyOf(block)) {
        if (block.hasAttribute("data-section-break")) block.removeAttribute("data-section-break");
        continue;
      }
      let start = this.sectionStarts.get(nextKey!);
      if (start === undefined) {
        start = "nextPage";
        const id = this.anchorIdOf(next);
        try {
          const info = id ? JSON.parse(this.exports.DocxSessionBridge.GetSectionInfo(this.handle, id)) as SectionInfo | null : null;
          start = info?.sectionStart ?? "nextPage";
        } catch { /* (an engine without section info) */ }
        this.sectionStarts.set(nextKey!, start);
      }
      if (block.getAttribute("data-section-break") !== start) block.setAttribute("data-section-break", start);
    }
  }

  private unwatchLineNumbers(): void {
    if (!this.lineNumberWatch) return;
    this.lineNumberWatch.observer.disconnect();
    window.clearTimeout(this.lineNumberWatch.timer);
    this.lineNumberWatch = null;
  }

  private wireBlocks(root: HTMLElement): void {
    root.querySelectorAll<HTMLElement>("[data-anchor]").forEach((el) => this.wireBlock(el));
  }

  private wireBlock(el: HTMLElement): void {
    // The paginator gives every line its font's unrounded height before it lays pages out. A block
    // rendered afresh after an edit has not had that pass; without it, it would advance at
    // Chromium's whole-pixel pitch while the rest of the page does not (issue #1151).
    if (this.options.paginated && el.isConnected) applyUnroundedNormalLineHeights(el);
    if (!EDITABLE_TAGS.has(el.tagName)) return;
    if (
      el.closest("tr[class$='row-del'], tr[class*='row-del ']") ||
      el.querySelector(":scope > del[class$='move-from'], :scope > del[class*='move-from ']")
    ) {
      el.setAttribute("contenteditable", "false");
      return;
    }
    const unid = el.getAttribute("data-anchor");
    // Only blocks the markdown projection addresses are editable via the text path. This INCLUDES
    // table-cell paragraphs (the projection indexes them), so cell text IS editable. Table-aware
    // key handling keeps structural edits safe while still providing Word-style cell navigation
    // (see onKeydown). Anything the projection does not index (unstamped content) stays read-only.
    // A band block is authoritative via its stamped `data-hf-anchor` even when the unid map
    // resolves that unid to a different part (content-addressed unids collide across parts).
    if (!unid || !this.anchorIdOf(el)) return;
    el.setAttribute("contenteditable", "true");
    ensureListMarkerSeparator(el);
    // Generated chrome (list number/bullet, footnote/endnote citation markers, note backrefs) is
    // not editable content — keep the caret out so offsets stay aligned with the run text, and so
    // a citation marker can't be deleted directly (which would orphan its note definition).
    el.querySelectorAll<HTMLElement>(GENERATED_CHROME_SELECTOR)
      .forEach((m) => m.setAttribute("contenteditable", "false"));
    const noteNumber = el.querySelector<HTMLElement>(NOTE_NUMBER_SELECTOR);
    if (noteNumber) this.noteNumbers.set(el, noteNumber.cloneNode(true) as HTMLElement);
    // Page-number fields are atomic, and remember their cached result so per-page substitution
    // in the page view never shifts the offset space (see `fieldText`).
    adoptFields(el);
    // Baseline for the commit diff: CONTENT text (list markers + injected bidi marks excluded),
    // matching the session's flat run-text offset space.
    el.dataset.committedText = blockContentText(el);
    el.addEventListener("focus", () => {
      this.activeBlock = el;
      // Follow the caret's section so a cover-page-plus-body document shows the stories that
      // actually apply. Focusing a BAND block must not re-sync: it has no governing section of
      // its own, and re-resolving would clobber the user's kind selection.
      if (this.region && !this.isBandBlock(el)) this.syncRegionToBody(el);
      this.region?.noteFocus(el);
    });
    el.addEventListener("blur", event => this.commitBlock(el, event.relatedTarget));
    el.addEventListener("input", (ev) => {
      noteTypedSpace(el, ev);
      this.onTypingInput(el, ev as InputEvent);
    });
    el.addEventListener("beforeinput", (ev) => {
      guardChromeDeletion(el, ev as InputEvent);
      deleteAdjacentWhitespaceAtom(el, ev as InputEvent);
    });
    el.addEventListener("keydown", (ev) => this.onKeydown(el, ev as KeyboardEvent));
    // A band block re-rendered by an incremental swap is a fresh DOM node; re-adopt it (with the
    // anchor its caller already stamped) so the band chrome can still address it.
    const stamped = el.getAttribute("data-hf-anchor");
    if (stamped && this.region?.contains(el)) this.region.adoptBlock(el, stamped);
  }

  /**
   * Replace `oldEl` with `newNodes`, suppressing the re-entrant blur→commit that removing a focused
   * block fires (see `replacing`), AND tolerating the case where a synchronous blur during focus
   * transfer detaches `oldEl` between the caller's checks and here. `replaceWith` then throws
   * NotFoundError ("node … no longer a child … moved in a blur event handler") — `isConnected`
   * alone doesn't catch this race. The session is already updated, so a skipped/failed visual swap
   * leaves correct content (the typed DOM); the next commit or remount reconciles it. This is why
   * the catch is silent rather than rethrowing — there's no lost data, only a deferred re-render.
   * Returns true if the swap happened.
   */
  private replaceNode(oldEl: HTMLElement, ...newNodes: Node[]): boolean {
    const prev = this.replacing;
    this.replacing = true;
    try {
      if (!oldEl.parentNode) return false;
      // A paginated note's number is chrome inside its first paragraph, drawn from the whole
      // note's context; a single-block render of that paragraph has none. A text edit never
      // changes it, so the paragraph that replaces the numbered one keeps it. The copy kept at
      // wiring covers a block whose selection was replaced across the number.
      const number = oldEl.querySelector(NOTE_NUMBER_SELECTOR) ?? this.noteNumbers.get(oldEl);
      const first = newNodes.find((node): node is HTMLElement => node.nodeType === 1);
      if (number && first && !first.querySelector(NOTE_NUMBER_SELECTOR)) {
        first.prepend(number.cloneNode(true));
      }
      const lifted = this.options.paginated ? this.liftedDrawingsOf(oldEl) : [];
      oldEl.replaceWith(...newNodes);
      if (lifted.length > 0) this.reliftDrawings(lifted, newNodes);
      return true;
    } catch {
      return false;
    } finally {
      this.replacing = prev;
    }
  }

  /** Commit a block edit on blur: diff → run-preserving session op → re-render only this block. */
  private commitBlock(el: HTMLElement, focusTarget: EventTarget | null): void {
    if (this.closed || this.replacing) return;
    const unid = el.getAttribute("data-anchor");
    if (!unid) return;
    const fullId = this.anchorIdOf(el);
    if (!fullId) return;

    const result = this.commitTextChange(el, fullId);
    if (!result) return; // no change
    if (!result.success) {
      this.restoreBlock(el, fullId);
      return;
    }

    const newAnchor = result.modified?.[0]?.id ?? fullId;
    const newUnid = result.modified?.[0]?.unid ?? unid;
    if (this.remountSplitParagraph(el, undefined, focusTarget)) {
      this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
      return;
    }

    // List items: do NOT re-render on a text commit. Re-rendering replaces the node *during* the
    // blur, cancelling the browser's in-flight focus transfer when the user clicks straight to
    // another bullet; numbering also needs whole-document context a single-block render lacks. The
    // DOM already shows what the user typed with the correct marker — sync the baseline only.
    if (el.querySelector(":scope > [data-list-marker]")) {
      el.dataset.committedText = blockContentText(el);
      this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
      return;
    }

    // Plain block: re-render ONLY this block from the live session for canonical HTML. Swapping the
    // just-blurred node here is safe (verified — focus stays on the newly-clicked block).
    const html = this.renderBlockHtml(newAnchor);
    if (html.charCodeAt(0) !== 0x7b /* not an error object */) {
      const fresh = new DOMParser().parseFromString(html, "text/html").body.firstElementChild as HTMLElement | null;
      if (fresh && this.remountSplitParagraph(el, fresh, focusTarget)) {
        this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
        return;
      }
      const inBand = this.isBandBlock(el);
      if (fresh && this.replaceNode(el, fresh)) {
        // Band blocks resolve through `data-hf-anchor`; their unid can collide with another
        // part's, so writing it into the map would corrupt that entry (see anchorIdOf).
        if (inBand) {
          this.region!.adoptBlock(fresh, newAnchor);
        } else {
          this.unidToFullId.delete(unid);
          this.unidToFullId.set(newUnid, newAnchor);
        }
        this.wireBlock(fresh);
        if (this.activeBlock === el) this.activeBlock = fresh; // keep ribbon target valid
        // The throwaway render numbers citation markers from 1 — repair in place.
        this.maybeRenumberNotes(fresh);
        if (inBand) this.region!.afterStoryEdit(fresh);
      }
    }

    this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
  }

  // ─── M2: structural editing ──────────────────────────────────────────

  private onKeydown(el: HTMLElement, ev: KeyboardEvent): void {
    if (this.closed) return;
    if (this.crossBlockKey(el, ev)) return;
    moveAcrossTab(el, ev);
    // Common formatting / history shortcuts.
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey) {
      const k = (ev.key ?? "").toLowerCase();
      const fmt: Record<string, FormatKey> = { b: "bold", i: "italic", u: "underline" };
      if (fmt[k]) { ev.preventDefault(); this.format(fmt[k]); return; }
      if (k === "z") { ev.preventDefault(); ev.shiftKey ? this.redo() : this.undo(); return; }
      if (k === "y") { ev.preventDefault(); this.redo(); return; }
      // Word's Copy Formatting / Paste Formatting (Format Painter from the keyboard).
      if (k === "c" && ev.shiftKey) { ev.preventDefault(); this.copyFormatting(); return; }
      if (k === "v" && ev.shiftKey) { ev.preventDefault(); this.pasteFormatting(); return; }
    }
    // Shift+F3: Word's Change Case cycle.
    if (ev.key === "F3" && ev.shiftKey && !ev.ctrlKey && !ev.metaKey && !ev.altKey) {
      ev.preventDefault();
      this.changeCase("cycle");
      return;
    }
    // Ctrl+Alt+Shift+V: paste the text only.
    if ((ev.ctrlKey || ev.metaKey) && ev.altKey && ev.shiftKey && (ev.key ?? "").toLowerCase() === "v") {
      ev.preventDefault();
      void this.pasteFromClipboard("text");
      return;
    }
    // Inside a table cell, Backspace at the cell's start does not merge across the cell
    // boundary (mid-text Backspace still deletes normally). Enter splits the cell paragraph
    // WITHIN the same cell — the engine keeps the new w:p in the w:tc, so the grid is unchanged.
    // Tab navigation is handled explicitly below; at the final cell it uses the existing safe
    // whole-table row insertion path to match Word's add-a-row behavior.
    const inTableCell = !!el.closest("table");

    // Plain Tab / Shift+Tab moves between table cells. Outside tables it retains list
    // nest/un-nest behavior. Modified Tab chords remain available to the browser/platform.
    if (ev.key === "Tab" && !ev.ctrlKey && !ev.metaKey && !ev.altKey && !ev.isComposing) {
      if (inTableCell) {
        ev.preventDefault();
        this.navigateTableCell(el, ev.shiftKey);
        return;
      }
      // In a list item, as in Word: Tab at the item's start (or over a selection) and Shift+Tab
      // change the level; Tab with the caret inside the text types a tab.
      const sel = el.ownerDocument.defaultView?.getSelection();
      const caretInText = !!sel?.isCollapsed && (caretOffsetIn(el) ?? 0) > 0;
      if (isListBlock(el) && (ev.shiftKey || !caretInText)) {
        ev.preventDefault();
        this.activeBlock = el;
        this.setListLevel(ev.shiftKey ? -1 : 1);
        return;
      }
      // Anywhere else Tab types a tab character (Word's), instead of moving the focus out of
      // the document. Shift+Tab has nothing to do outside a list and a table.
      ev.preventDefault();
      if (!ev.shiftKey) this.insertTabAtCaret(el);
      return;
    }
    // Shift+Enter inserts an intra-paragraph line break (a real w:br on commit),
    // not a paragraph split. Deterministic across browsers and allowed in cells
    // (a line break changes no table structure).
    if (ev.key === "Enter" && ev.shiftKey && !ev.isComposing) {
      ev.preventDefault();
      this.insertLineBreakAtCaret(el);
      return;
    }
    // Ctrl+Enter is the page break (Word's), which the host's shortcut handles.
    if (ev.key === "Enter" && (ev.ctrlKey || ev.metaKey)) return;
    if (ev.key === "Enter" && !ev.shiftKey && !ev.isComposing) {
      ev.preventDefault();
      // Splits at the caret — in a cell this stacks a second paragraph within the same
      // w:tc (grid unchanged); in the body it splits the paragraph as before.
      this.splitAtCaret(el);
    } else if (ev.key === "Backspace") {
      const sel = typeof window !== "undefined" ? window.getSelection() : null;
      // An empty paragraph shows a placeholder space the caret can sit after: it is at its start.
      // A tab or a line break is content, not placeholder padding.
      const atStart = caretOffsetIn(el) === 0 || trimSpaces(blockContentText(el)).length === 0;
      if (sel && sel.isCollapsed && atStart) {
        // A paginated note's first paragraph opens with its number, which is not run text, and
        // the note is its own story: there is nothing before the text to delete or merge into.
        if (el.querySelector(NOTE_NUMBER_SELECTOR)) {
          ev.preventDefault();
          return;
        }
        // In a table cell the paragraph joins the one before it in the same cell; at the cell's
        // first paragraph there is nothing to join (as in Word, the cell boundary stays).
        const prev = this.joinableNeighbour(el, -1);
        // At a section's first paragraph (or in an empty one that only ends its section) Backspace
        // removes the section break, as Word's Backspace over a break mark does.
        const sectionBreak = !inTableCell && !this.isBandBlock(el) ? this.sectionBreakAround(el, -1) : null;
        if (sectionBreak) {
          ev.preventDefault();
          this.removeBreak(sectionBreak, el);
          return;
        }
        if (inTableCell && !prev) {
          ev.preventDefault();
          return;
        }
        if (prev) {
          ev.preventDefault();
          // The previous paragraph's mark is already a pending deletion (a tracked join waiting
          // for review), so the two are joined already. As in Word, Backspace steps over the
          // struck pilcrow and removes the live character before it, instead of joining again.
          if (prev.querySelector(DELETED_PARAGRAPH_MARK_SELECTOR)) {
            placeCaretAtOffset(prev, blockContentText(prev).length);
            deleteLiveCharacterBeyond(prev, "deleteContentBackward");
            return;
          }
          this.mergeWithPrevious(prev, el);
        }
      }
    } else if (ev.key === "Delete" && !ev.ctrlKey && !ev.metaKey && !ev.altKey && !ev.shiftKey) {
      // Delete at a paragraph's end joins the next paragraph to it, as in Word (in a table cell,
      // only the next paragraph of the same cell). Anywhere else the browser deletes as usual.
      const sel = typeof window !== "undefined" ? window.getSelection() : null;
      const at = sel?.isCollapsed ? caretOffsetIn(el) : null;
      if (at === null || trimSpaces(blockContentText(el).slice(at)).length > 0) return;
      // At a section's last paragraph Delete removes the break after it, as Word's Delete on a
      // break mark does.
      const sectionBreak = !inTableCell && !this.isBandBlock(el) ? this.sectionBreakAround(el, 1) : null;
      if (sectionBreak) {
        ev.preventDefault();
        this.removeBreak(sectionBreak, el);
        return;
      }
      const next = this.joinableNeighbour(el, 1);
      if (!next || next.querySelector(NOTE_NUMBER_SELECTOR)) {
        if (inTableCell) ev.preventDefault();
        return;
      }
      // A pending deletion of either paragraph mark (a tracked join) is left to the browser's own
      // editing, as Backspace leaves the struck pilcrow.
      if (el.querySelector(DELETED_PARAGRAPH_MARK_SELECTOR) || next.querySelector(DELETED_PARAGRAPH_MARK_SELECTOR)) return;
      // So is a paragraph with a shown deletion: the caret may sit before deleted text, which is not
      // the paragraph's end.
      if (el.querySelector(RENDERED_TEXT_DELETION_SELECTOR)) return;
      ev.preventDefault();
      this.mergeWithPrevious(el, next);
    }
  }

  /** The section break Backspace (`step` -1, at `el`'s start) or Delete (+1, at its end) would
   *  remove: the one between `el` and the body paragraph before or after it when they are in
   *  different sections, or the one `el` ends when it is an empty paragraph that only holds it. */
  private sectionBreakAround(el: HTMLElement, step: -1 | 1): string | null {
    const body = this.editableList().filter((b) => !this.isBandBlock(b) && !b.closest("td, th"));
    const i = body.indexOf(el);
    if (i < 0) return null;
    const key = this.sectionKeyOf(el);
    const next = body[i + 1];
    const endsSection = !!next && this.sectionKeyOf(next) !== key;
    if (endsSection && (step === 1 || /^[\s\u00a0]*$/.test(blockContentText(el)))) return this.sectionBreakAfter(el);
    const prev = body[i - 1];
    if (step === -1 && prev && this.sectionKeyOf(prev) !== key) return this.sectionBreakAfter(prev);
    return null;
  }

  /** The editable paragraph before (`step` -1) or after (+1) `el` that Backspace or Delete may join
   *  it with: inside a table cell only one of the same cell, and never a table cell's paragraph
   *  from outside the table. */
  private joinableNeighbour(el: HTMLElement, step: -1 | 1): HTMLElement | null {
    const all = Array.from(this.ownerRoot(el).querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'));
    const i = all.indexOf(el);
    const other = i < 0 ? null : all[i + step] ?? null;
    if (!other) return null;
    const cell = el.closest("td, th");
    return cell ? (other.closest("td, th") === cell ? other : null) : (other.closest("td, th") ? null : other);
  }

  /** Editable cells in visual document order, excluding cells from any nested table. */
  private tableCellsFor(el: HTMLElement): { cells: HTMLElement[]; index: number } | null {
    const cell = el.closest<HTMLElement>("td, th");
    const table = cell?.closest<HTMLTableElement>("table");
    if (!cell || !table) return null;
    const cells = Array.from(table.querySelectorAll<HTMLElement>("td, th"))
      .filter((candidate) => candidate.closest("table") === table);
    const index = cells.indexOf(cell);
    return index >= 0 ? { cells, index } : null;
  }

  /** First addressable paragraph in a cell, fenced against nested-table descendants. */
  private editableInCell(cell: HTMLElement): HTMLElement | null {
    return Array.from(
      cell.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    ).find((block) => block.closest("td, th") === cell) ?? null;
  }

  /** Word-style cell navigation: Shift+Tab goes back, Tab goes forward, and Tab at the
   *  final cell appends a row before entering its first cell. The first cell is a hard
   *  boundary for Shift+Tab so focus never leaks out of the editor table. */
  private navigateTableCell(el: HTMLElement, backwards: boolean): void {
    const state = this.tableCellsFor(el);
    if (!state) return;
    const adjacent = state.cells[state.index + (backwards ? -1 : 1)];
    if (adjacent) {
      const target = this.editableInCell(adjacent);
      if (target) placeCaretAtOffset(target, 0);
      return;
    }
    if (backwards) return;

    // The current block may contain uncommitted typing. insertTableRow flushes it before the
    // structural edit, then reconcile/remount restores focus by block index. Resolve the new
    // adjacent cell from that fresh DOM rather than retaining a node from the replaced table.
    this.activeBlock = el;
    this.insertTableRow("below");
    const fresh = this.activeBlock ? this.tableCellsFor(this.activeBlock) : null;
    if (!fresh) return;
    const target = fresh.cells[fresh.index + 1];
    const block = target ? this.editableInCell(target) : null;
    if (block) placeCaretAtOffset(block, 0);
  }

  /** Shift+Enter: insert an intra-paragraph line break at the caret. Delegates to the
   *  native `insertLineBreak` command, which inserts a <br> AND positions the caret
   *  after it correctly (handling the browser's bogus trailing-<br> rule) so typing
   *  continues on the new line. Commits (on blur) as a w:br via the "  \n" hard break
   *  the serializer emits for a <br>. */
  private insertLineBreakAtCaret(block: HTMLElement): void {
    if (typeof document !== "undefined" && typeof document.execCommand === "function") {
      const selection = block.ownerDocument.defaultView?.getSelection();
      const range = selection?.rangeCount ? selection.getRangeAt(0) : null;
      const before = blockContentText(block).length;
      const start = range ? contentOffsetOf(block, range.startContainer, range.startOffset) : 0;
      const end = range ? contentOffsetOf(block, range.endContainer, range.endOffset) : 0;
      document.execCommand("insertLineBreak");
      // With pre-wrap, Chromium can append TWO literal LF text nodes instead of BRs:
      // one break and one caret placeholder. Normalize only that extra terminal character
      // to the same empty BR placeholder the shared text/offset helpers already skip.
      if (end !== before || blockContentText(block).length !== before - (end - start) + 2) return;
      const { node, offset } = contentPositionIn(block, start + 2);
      if (node.nodeType !== 3 || node.textContent?.[offset - 1] !== "\n") return;
      (node as Text).splitText(offset - 1).replaceWith(block.ownerDocument.createElement("br"));
      placeCaretAtOffset(block, start + 1);
    }
  }

  /** Enter: split the block at the caret into two paragraphs. */
  private splitAtCaret(el: HTMLElement): void {
    const rawOffset = caretOffsetIn(el);
    const unid = el.getAttribute("data-anchor");
    if (rawOffset == null || !unid) return;
    let fullId = this.anchorIdOf(el);
    if (!fullId) return;
    // The session commits trimmed text, so map the DOM caret offset into the trimmed run-text
    // offset (else an overshoot — e.g. a placeholder leading space — is rejected and Enter is lost).
    const offset = trimmedSplitOffset(el, rawOffset);

    const idx = this.blockIndex(el); // capture before the op (for list remount focus)
    // Whether this paragraph is rendered inside a border <div> — captured before the DOM mutates.
    const wrappedInBorder = inBorderWrapper(el);
    fullId = this.syncBlock(el, fullId); // flush any uncommitted typing first
    if (!fullId) return;
    const res = this.parseEdit(this.exports.DocxSessionBridge.SplitParagraph(this.handle, fullId, offset));
    if (!res.success) return;
    const first = res.modified?.[0];
    const second = res.created?.[0];
    if (!first || !second) return;

    // Splitting a list item makes a continuing list item, and splitting a bordered paragraph (e.g.
    // a horizontal rule) yields a new paragraph whose border status differs — both need a whole-
    // document re-render so numbering continues / border <div>s regroup correctly. An in-place node
    // swap would leave the new paragraph inside the old border div (the rule's line under its text).
    if (this.affectsList(res) || wrappedInBorder || el.dataset.docxSplitParagraph === "true") {
      this.refreshAfter(el, idx + 1, false);
      this.options.onEdit?.({ anchorId: second.id, unid: second.unid });
      return;
    }

    const [firstEl, secondEl] = this.renderTwo(first.id, second.id);
    if (!firstEl || !secondEl) return;

    // el is the focused block — replaceNode guards the re-entrant blur→commit and tolerates a
    // node detached mid-focus-transfer; replacing with both new blocks at once keeps them adjacent.
    const inBand = this.isBandBlock(el);
    if (!this.replaceNode(el, firstEl, secondEl)) return;
    // Band blocks resolve through `data-hf-anchor` (their unids can collide across parts).
    if (inBand) {
      this.region!.adoptBlock(firstEl, first.id);
      this.region!.adoptBlock(secondEl, second.id);
    } else {
      this.unidToFullId.delete(unid);
      this.unidToFullId.set(first.unid, first.id);
      this.unidToFullId.set(second.unid, second.id);
    }
    this.wireBlock(firstEl);
    this.wireBlock(secondEl);
    if (inBand) this.region!.afterStoryEdit(secondEl);
    placeCaretAtOffset(secondEl, 0);
    this.options.onEdit?.({ anchorId: second.id, unid: second.unid });
  }

  /** Backspace at block start: merge this block into the previous one. */
  private mergeWithPrevious(prev: HTMLElement, el: HTMLElement): void {
    const prevUnid = prev.getAttribute("data-anchor");
    const thisUnid = el.getAttribute("data-anchor");
    if (!prevUnid || !thisUnid) return;
    let prevId = this.anchorIdOf(prev);
    let thisId = this.anchorIdOf(el);
    if (!prevId || !thisId) return;

    const prevIdx = this.blockIndex(prev); // capture before the op
    // Either side rendered inside a border <div> means the merge changes border grouping — captured
    // before the DOM mutates so the post-merge branch can force a full re-render.
    const wrappedInBorder = inBorderWrapper(prev) || inBorderWrapper(el);
    prevId = this.syncBlock(prev, prevId);
    thisId = this.syncBlock(el, thisId);
    if (!prevId || !thisId) return;
    const caret = (prev.textContent ?? "").length; // merge boundary

    const res = this.parseEdit(this.exports.DocxSessionBridge.MergeParagraphs(this.handle, prevId, thisId));
    if (!res.success) return;
    const merged = res.modified?.[0];
    if (!merged) return;

    // Merging list items renumbers the list, and merging across a border <div> boundary changes the
    // border grouping — both need a whole-document re-render (caret at the merge boundary). So does a
    // tracked merge, which removes nothing: both paragraphs stay until the deleted paragraph mark
    // is accepted, so the second must not be dropped from the page.
    if (this.affectsList(res) || wrappedInBorder || !res.removed?.length
      || prev.dataset.docxSplitParagraph === "true" || el.dataset.docxSplitParagraph === "true") {
      this.refreshAfter(el, prevIdx, true);
      this.options.onEdit?.({ anchorId: merged.id, unid: merged.unid });
      return;
    }

    const mergedEl = this.renderInto(merged.id);
    if (!mergedEl) return;

    // prev may be focused — replaceNode guards re-entrancy and tolerates a detached node.
    const inBand = this.isBandBlock(prev);
    if (!this.replaceNode(prev, mergedEl)) return;
    el.remove();
    // Band blocks resolve through `data-hf-anchor` (their unids can collide across parts).
    if (inBand) {
      this.region!.adoptBlock(mergedEl, merged.id);
    } else {
      this.unidToFullId.delete(prevUnid);
      this.unidToFullId.delete(thisUnid);
      this.unidToFullId.set(merged.unid, merged.id);
    }
    this.wireBlock(mergedEl);
    if (inBand) this.region!.afterStoryEdit(mergedEl);
    placeCaretAtOffset(mergedEl, caret);
    this.options.onEdit?.({ anchorId: merged.id, unid: merged.unid });
  }

  /**
   * Apply the block's pending text change to the session with full inline-formatting fidelity.
   * Diffs the committed content text (markers + bidi excluded) against the current content text
   * and rewrites only the changed span via ReplaceTextAtSpan — every untouched run keeps its exact
   * rPr, and typed text inherits the boundary run's formatting. Returns the parsed EditResult, or
   * null when there is no change. Empty/whitespace-only baselines (e.g. the placeholder space the
   * converter renders for an empty paragraph, whose DOM text doesn't line up with the session's
   * empty run text) are rebuilt via ReplaceText — there is no inline formatting to preserve there.
   */
  private commitTextChange(el: HTMLElement, fullId: string): EditResultLite | null {
    // `old` mirrors the session's flat run-text: strip bidi marks (blockContentText strips them,
    // but wireBlock may have stored textContent before this Task 2 change, and the bidi test
    // explicitly stores textContent). Using stripBidi keeps the baseline consistent with the
    // session's offset space regardless of how committedText was stored.
    // contenteditable keeps a typed trailing space as U+00A0 so it stays visible; the document
    // wants an ordinary space there (Word's own), and the two are the same length, so the
    // offset space is unchanged. Both sides are normalized the same way, so an untouched
    // placeholder (the NBSP an empty paragraph renders as) never reads as an edit \u2014 a blur that
    // committed "" over a paragraph another op just filled would wipe that op's work.
    const old = stripBidi(el.dataset.committedText ?? "").replace(/\u00a0$/, " ");
    const next = blockContentText(el).replace(/\u00a0$/, " ");
    if (old === next) return null;

    // A whitespace-only baseline means the paragraph was (or rendered as) empty, so the typed
    // text replaces it wholesale — unless it holds a field whose cached result is empty (a PAGE
    // field in a tool-generated footer): that is content, and the span path below keeps it
    // while inserting the typed text beside it. An empty split head likewise needs a span
    // insertion so the authored break and the content on later pages survive.
    if (trimSpaces(old).length === 0 && !el.querySelector("[data-field]")
      && el.dataset.docxSplitParagraph !== "true") {
      // A paragraph whose only text is deleted renders that placeholder too, and with the markup
      // hidden nothing on the page shows the deleted text. Tracked mode refuses the rewrite when
      // the paragraph holds a bookmark; an untracked rewrite would discard a shown deletion. The
      // typed text then goes in as a plain insertion at its start, which cannot carry a line
      // break or inline formatting.
      const tracked = this.options.trackedChanges === TrackedChangeMode.RenderInline;
      const typed = trimSpaces(next) + (keepsTrailingSpace(el) ? " " : "");
      // The side of a shown deletion the text was typed on. A whole-paragraph rewrite would put
      // it after the deletion whichever side that was, so plain text goes in as a span insertion
      // on its own side; a line break or inline formatting still needs the rewrite.
      const typedText = typed.replace(TRAILING_SPACES, "");
      const side = deletedTextSideOf(el, next.indexOf(typedText), typedText.length);
      const besideShownDeletion = side !== undefined && holdsOnlyPlainText(el);
      if (!besideShownDeletion && (tracked || !el.querySelector(RENDERED_TEXT_DELETION_SELECTOR))) {
        const whole = this.parseEdit(
          this.exports.DocxSessionBridge.ReplaceText(this.handle, fullId, serializeInlineMarkdown(el)),
          { expectRejection: true },
        );
        if (whole.success || whole.error?.code !== "tracked_operation_unsupported") {
          if (!whole.success) this.reportEditFailure(whole.error);
          return whole;
        }
      }
      return this.parseEdit(
        this.exports.DocxSessionBridge.ReplaceTextAtSpan(this.handle, fullId, 0, 0, typed, side ?? ""),
      );
    }

    const minLen = Math.min(old.length, next.length);
    let p = 0;
    while (p < minLen && old[p] === next[p]) p++;
    let s = 0;
    while (s < minLen - p && old[old.length - 1 - s] === next[next.length - 1 - s]) s++;

    let start = p;
    let len = old.length - p - s;
    let middle = next.slice(p, next.length - s);

    if (len === 0) {
      // A pure insertion. The engine inserts a NEW run at a run boundary (and steps outside a
      // field's chrome, so typing after "Page X of Y" never lands inside the NUMPAGES result) or
      // inside an ordinary text run, and inside a tracked insertion it splits another author's
      // insertion around the new text or extends the typist's own (issue #1318). It refuses an
      // offset inside a hyperlink, field or other inline container or a mixed-content run, as
      // does an older engine build. Then anchor a neighbour char so the span is non-empty and the
      // inserted text inherits an adjacent run's rPr (the LEFT run when there is one, matching
      // contenteditable; the first run at the very start). Under tracking that records the
      // neighbour as deleted and retyped: the accepted and rejected text are right, but the
      // markup has one extra pair of revisions.
      // Beside a shown deletion, the side the text was typed on decides where it goes (#1290).
      const side = deletedTextSideOf(el, start, middle.length) ?? "";
      const inserted = this.parseEdit(
        this.exports.DocxSessionBridge.ReplaceTextAtSpan(this.handle, fullId, start, 0, middle, side),
        { expectRejection: true }, // a refusal here only selects the fallback below
      );
      if (inserted.success) return inserted;
      if (start > 0) { start -= 1; len = 1; middle = old[start] + middle; }
      else { len = 1; middle = middle + old[0]; }
    }

    return this.parseEdit(
      this.exports.DocxSessionBridge.ReplaceTextAtSpan(this.handle, fullId, start, len, middle),
    );
  }

  /**
   * Flush a block's current (uncommitted) text to the session; returns the live full id, or
   * undefined when the engine rejected the text. The block is then restored from the session, and the
   * caller must stop: the operation it was about to run would act on text the document does not
   * hold, at offsets measured in the rejected DOM.
   */
  private syncBlock(el: HTMLElement, fullId: string): string | undefined {
    const result = this.commitTextChange(el, fullId);
    if (!result) return fullId;
    if (!result.success) {
      this.restoreBlock(el, fullId);
      return undefined;
    }
    el.dataset.committedText = blockContentText(el);
    return result.modified?.[0]?.id ?? fullId;
  }

  /**
   * The session refused an edit to `el`, so re-render the block from the session to discard the
   * DOM the user typed but the document never accepted.
   */
  private restoreBlock(el: HTMLElement, fullId: string): void {
    const fresh = this.renderInto(fullId);
    if (fresh && this.remountSplitParagraph(el, fresh)) return;
    if (fresh && this.replaceNode(el, fresh)) {
      this.wireBlock(fresh);
      if (this.activeBlock === el) this.activeBlock = fresh;
    }
  }

  /** A split source owns pieces on several pages; replacing only its editable head leaves stale tails. */
  private remountSplitParagraph(old: HTMLElement, fresh?: HTMLElement, focusTarget: EventTarget | null = old): boolean {
    if (!this.options.paginated || (old.dataset.docxSplitParagraph !== "true"
      && !fresh?.hasAttribute("data-docx-break-paragraph"))) return false;
    const replacing = this.replacing;
    this.replacing = true;
    const target = focusTarget instanceof HTMLElement ? focusTarget.closest<HTMLElement>('[data-anchor]') : null;
    try { this.refreshAfter(old, target ? this.blockIndex(target) : -1); }
    finally { this.replacing = replacing; }
    return true;
  }

  /** Render a block by anchor and parse it into a detached element (null on error). */
  private renderInto(anchorId: string): HTMLElement | null {
    const html = this.renderBlockHtml(anchorId);
    if (html.charCodeAt(0) === 0x7b /* error object */) return null;
    return new DOMParser().parseFromString(html, "text/html").body.firstElementChild as HTMLElement | null;
  }

  private get renderTrackedChanges(): boolean {
    return this.options.showRevisions ?? this.options.trackedChanges === TrackedChangeMode.RenderInline;
  }

  private renderBlockHtml(anchorId: string): string {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.RenderEditorBlockHtml === "function") {
      return bridge.RenderEditorBlockHtml(this.handle, anchorId, this.editorRenderProfile());
    }
    if (typeof bridge.RenderBlockHtmlForReview === "function") {
      return bridge.RenderBlockHtmlForReview(
        this.handle, anchorId, this.options.cssPrefix, this.options.fabricateClasses,
        this.renderTrackedChanges,
      );
    }
    return bridge.RenderBlockHtml(
      this.handle, anchorId, this.options.cssPrefix, this.options.fabricateClasses,
    );
  }

  private renderPlanJson(): string {
    const bridge = this.exports.DocxSessionBridge;
    return typeof bridge.ListRenderedBlocks === "function"
      ? bridge.ListRenderedBlocks(this.handle, this.renderTrackedChanges)
      : bridge.ListBlocks!(this.handle);
  }

  /** Batch-render `idsJson` anchors through the richest endpoint the bundle carries; the JSON
   *  maps each anchor to its HTML (null when it failed to resolve), or carries `error`. */
  private renderBlocksJson(idsJson: string): string {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.RenderEditorBlocksHtml === "function") {
      return bridge.RenderEditorBlocksHtml(this.handle, idsJson, this.editorRenderProfile());
    }
    if (typeof bridge.RenderBlocksHtmlForReview === "function") {
      return bridge.RenderBlocksHtmlForReview(
        this.handle, idsJson, this.options.cssPrefix, this.options.fabricateClasses,
        this.renderTrackedChanges,
      );
    }
    return bridge.RenderBlocksHtml!(
      this.handle, idsJson, this.options.cssPrefix, this.options.fabricateClasses,
    );
  }

  /** Render two blocks in ONE batched bridge call when the bundle carries RenderBlocksHtml —
   *  the per-render shell/converter setup is paid once instead of twice, which matters on the
   *  Enter path (split renders both halves synchronously under the keystroke). Falls back to
   *  two per-block renders on older bundles. Output is renderInto-identical per block (both
   *  routes share the same extraction). */
  private renderTwo(a: string, b: string): [HTMLElement | null, HTMLElement | null] {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.RenderBlocksHtml === "function") {
      try {
        const idsJson = JSON.stringify([a, b]);
        const json = this.renderBlocksJson(idsJson);
        const map = JSON.parse(json) as Record<string, string | null> & { error?: string };
        if (!map.error) {
          const parse = (h: string | null | undefined): HTMLElement | null =>
            h
              ? (new DOMParser().parseFromString(h, "text/html").body
                  .firstElementChild as HTMLElement | null)
              : null;
          const fa = parse(map[a]);
          const fb = parse(map[b]);
          if (fa && fb) return [fa, fb];
        }
      } catch {
        /* fall through to per-block renders */
      }
    }
    return [this.renderInto(a), this.renderInto(b)];
  }

  /** The editable block immediately before `el` within its own root, or null. */
  private previousEditable(el: HTMLElement): HTMLElement | null {
    const all = Array.from(
      this.ownerRoot(el).querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    );
    const i = all.indexOf(el);
    return i > 0 ? all[i - 1] : null;
  }

  /**
   * Tab: a tab character (a real `w:tab`) at the caret, or in place of the selected text within
   * this block. Typing not yet committed goes first; the block is then re-rendered from the session
   * with the caret after the tab. One edit, one undo step.
   */
  private insertTabAtCaret(el: HTMLElement): void {
    if (this.closed || !el.isConnected) return;
    const unid = el.getAttribute("data-anchor");
    let fullId = this.anchorIdOf(el);
    if (!unid || !fullId) return;
    const sel = el.ownerDocument.defaultView?.getSelection();
    const range = sel?.rangeCount ? sel.getRangeAt(0) : null;
    if (!range || !el.contains(range.startContainer) || !el.contains(range.endContainer)) return;
    const from = contentOffsetOf(el, range.startContainer, range.startOffset);
    const to = contentOffsetOf(el, range.endContainer, range.endOffset);
    const span = trimmedSpan(el, { start: from, length: to - from });
    fullId = this.syncBlock(el, fullId);
    if (!fullId) return;
    const bridge = this.exports.DocxSessionBridge;
    let res = this.parseEdit(
      bridge.ReplaceTextAtSpan(this.handle, fullId, span.start, span.length, "\t"),
      { expectRejection: true },
    );
    // An empty paragraph has no run to insert into: give it the tab as its whole text.
    if (!res.success && trimSpaces(blockContentText(el)).length === 0) {
      res = this.parseEdit(bridge.ReplaceText(this.handle, fullId, "\t"), { expectRejection: true });
    }
    if (!res.success) {
      this.reportEditFailure(res.error);
      return;
    }
    const fresh = this.swapBlock(el, unid, res.modified?.[0]);
    if (!fresh?.isConnected) return;
    fresh.focus();
    const after = fresh.ownerDocument.defaultView?.getSelection();
    if (after) collapseAtContentOffset(fresh, after, span.start + 1);
  }

  /**
   * Read the result of a session mutation. Every editor mutation's result passes through here, so
   * this is where a rejection is reported (see {@link DocxEditorOptions.onEditFailed}), once, for
   * all of them. A caller probing an op it expects may be refused passes `expectRejection`.
   */
  private parseEdit(json: string, opts?: { expectRejection?: boolean }): EditResultLite {
    let result: EditResultLite;
    try {
      result = JSON.parse(json) as EditResultLite;
    } catch {
      result = { success: false };
    }
    if (result.success) {
      this.invalidateBlockMoveTargets();
      this.gutter?.schedule();
    } else if (!opts?.expectRejection) {
      this.reportEditFailure(result.error);
    }
    return result;
  }

  private reportEditFailure(error: EditResultLite["error"]): void {
    const active = this.activeBlock;
    const failure: EditFailure = {
      code: error?.code,
      message: error?.message ?? "The engine rejected the edit, or returned an unreadable result",
      anchorId: error?.anchorId ?? (active ? this.anchorIdOf(active) ?? null : null),
    };
    if (this.options.onEditFailed) this.options.onEditFailed(failure);
    else console.warn(`[DocxEditor] the engine rejected an edit${failure.code ? ` (${failure.code})` : ""}: ${failure.message}`);
  }

  /**
   * Drop the memoized `ValidMoveTargets` answers. Which blocks a block may move next to is a
   * fact about the DOCUMENT, so it survives hovering but not editing — and the two places a
   * document changes are `parseEdit` (every mutation that returns an `EditResult`) and
   * undo/redo, which return a bare boolean and so cannot go through it.
   */
  private invalidateBlockMoveTargets(): void {
    this.blockMoveTargetCache.clear();
  }

  // ─── M5: formatting commands (ribbon) ────────────────────────────────

  // ─── Multi-block selection helpers (format a whole stack of paragraphs at once) ──────

  /**
   * Restore a cross-block selection after a native toolbar control took focus. The bookmark uses
   * stable anchor ids and content offsets, so it also survives incremental block swaps.
   */
  private restoreCrossBlockSelection(): boolean {
    const bookmark = this.lastCrossBlockSelection;
    const active = this.activeBlock;
    if (!bookmark || !active) return false;
    const all = Array.from(
      this.ownerRoot(active).querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    );
    const first = all.find((block) => block.getAttribute("data-anchor") === bookmark.startUnid);
    const last = all.find((block) => block.getAttribute("data-anchor") === bookmark.endUnid);
    if (!first || !last || first === last || all.indexOf(first) > all.indexOf(last)) return false;
    const start = contentPositionIn(first, bookmark.startOffset);
    const end = contentPositionIn(last, bookmark.endOffset);
    return setSelectionBetween(start, end);
  }

  // ─── Selections across blocks: made by keyboard and Shift+click, and typed over ────────────
  //
  // Every block is its own editing host, and a browser fences a selection at a host's edge: past
  // the mouse drag (onMouseMove), it cannot extend one into the next paragraph by Shift+click,
  // Shift+arrow keys or Ctrl+A, and typing over one edits only the host that has the focus. Here
  // the editor does those as Word does, and turns typing, Backspace, Delete, Enter, cut and paste
  // over such a selection into one engine edit (one undo step).

  /** The editable blocks a live selection spanning two or more blocks covers, or null. */
  private crossBlockSelection(): HTMLElement[] | null {
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null;
    const range = sel.getRangeAt(0);
    const a = this.editableBlockOf(range.startContainer);
    const b = this.editableBlockOf(range.endContainer);
    if (!a || !b || a === b || this.ownerRoot(a) !== this.ownerRoot(b)) return null;
    const blocks = this.selectedBlocks();
    return blocks.length > 1 ? blocks : null;
  }

  /** The editable blocks of `el`'s story (the body, or one header/footer band), in order. */
  private storyBlocks(el: HTMLElement): HTMLElement[] {
    const root = this.ownerRoot(el);
    return Array.from(root.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))
      .filter((b) => this.ownerRoot(b) === root);
  }

  /** The fixed and the moving end of the live selection. */
  private selectionEnds(): { anchor: DomSelectionPoint; focus: DomSelectionPoint } | null {
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.rangeCount === 0 || !sel.anchorNode || !sel.focusNode) return null;
    const kept = this.extendedSelection;
    if (kept && kept.anchor.node.isConnected && kept.focus.node.isConnected) {
      const range = sel.getRangeAt(0);
      const same = (p: DomSelectionPoint, node: Node, offset: number) => p.node === node && p.offset === offset;
      if ((same(kept.anchor, range.startContainer, range.startOffset) && same(kept.focus, range.endContainer, range.endOffset))
        || (same(kept.focus, range.startContainer, range.startOffset) && same(kept.anchor, range.endContainer, range.endOffset))) {
        return kept;
      }
    }
    return { anchor: { node: sel.anchorNode, offset: sel.anchorOffset }, focus: { node: sel.focusNode, offset: sel.focusOffset } };
  }

  /**
   * Select from `anchor` to `focus` (in either order, across blocks) and remember which is which.
   * Chromium trims a range that starts outside the focused block back to that block's edge, so
   * the block holding the range's start takes the focus first. Leaving the old block commits its
   * typing and may render it afresh, so both points are held as (block, text offset) across it.
   */
  private selectBetween(anchor: DomSelectionPoint, focus: DomSelectionPoint): void {
    const aBlock = this.editableBlockOf(anchor.node);
    const fBlock = this.editableBlockOf(focus.node);
    if (aBlock && fBlock && aBlock !== fBlock && typeof document !== "undefined") {
      const story = this.storyBlocks(aBlock);
      const mark = (p: DomSelectionPoint, b: HTMLElement) => ({ index: story.indexOf(b), offset: contentOffsetOf(b, p.node, p.offset) });
      const am = mark(anchor, aBlock), fm = mark(focus, fBlock);
      const start = story[Math.min(am.index, fm.index)];
      if (am.index >= 0 && fm.index >= 0 && start && document.activeElement !== start) {
        start.focus({ preventScroll: true });
        const fresh = this.storyBlocks(start);
        if (fresh[am.index] && fresh[fm.index]) {
          anchor = contentPositionIn(fresh[am.index], am.offset);
          focus = contentPositionIn(fresh[fm.index], fm.offset);
        }
      }
    }
    if (setSelectionBetween(anchor, focus)) this.extendedSelection = { anchor, focus };
  }

  /** The point where `block`'s text starts or ends. */
  private static blockEdge(block: HTMLElement, end: boolean): DomSelectionPoint {
    return contentPositionIn(block, end ? blockContentText(block).length : 0);
  }

  /**
   * Shift+arrow keys, Shift+Home/End and their Ctrl forms, as Word has them. Inside one block the
   * browser's own movement stands (words, lines, bidi); the editor steps in only where the browser
   * cannot go on: at a block's edge, moving on into the next or previous block, and once the
   * selection spans blocks, where it moves the selection's moving end with the same native step
   * taken from a caret at that end. Returns true when it handled the key.
   */
  private extendSelectionByKey(el: HTMLElement, ev: KeyboardEvent): boolean {
    if (!ev.shiftKey || ev.altKey || ev.metaKey) return false;
    const forward = ev.key === "ArrowRight" || ev.key === "ArrowDown" || ev.key === "End";
    const ctrl = ev.ctrlKey;
    let granularity: "character" | "word" | "line" | "paragraph" | "lineboundary" | null = null;
    if (ev.key === "ArrowLeft" || ev.key === "ArrowRight") granularity = ctrl ? "word" : "character";
    else if (ev.key === "ArrowUp" || ev.key === "ArrowDown") granularity = ctrl ? "paragraph" : "line";
    else if (ev.key === "Home" || ev.key === "End") granularity = "lineboundary";
    if (!granularity) return false;
    const sel = window.getSelection();
    const ends = this.selectionEnds();
    if (!sel || !ends) return false;
    const focusBlock = this.editableBlockOf(ends.focus.node);
    const anchorBlock = this.editableBlockOf(ends.anchor.node);
    if (!focusBlock || !anchorBlock || this.ownerRoot(focusBlock) !== this.ownerRoot(anchorBlock)) return false;
    const story = this.storyBlocks(focusBlock);

    // Ctrl+Shift+Home/End: to the start or end of the story.
    if (ctrl && granularity === "lineboundary") {
      const target = forward ? story[story.length - 1] : story[0];
      if (!target) return false;
      ev.preventDefault();
      this.selectBetween(ends.anchor, DocxEditor.blockEdge(target, forward));
      return true;
    }

    // Where the browser's own step takes a caret at the moving end (within that block).
    const across = anchorBlock !== focusBlock;
    const extend = sel as Selection & { modify?: (alter: string, direction: string, granularity: string) => void };
    if (typeof extend.modify !== "function") return false;
    let moved: DomSelectionPoint | null = null;
    try {
      sel.collapse(ends.focus.node, ends.focus.offset);
      extend.modify("move", forward ? "forward" : "backward", granularity);
      if (sel.focusNode) moved = { node: sel.focusNode, offset: sel.focusOffset };
    } catch {
      moved = null;
    }
    // (a line step on a block's last or first line goes to that line's end or start instead of
    // failing: still the edge, in Word the next paragraph's line)
    const stayed = !moved || (moved.node === ends.focus.node && moved.offset === ends.focus.offset)
      || this.editableBlockOf(moved.node) !== focusBlock
      || (granularity === "line" && sameLine(ends.focus, moved));
    // Put the selection back as it was before deciding.
    if (across) setSelectionBetween(ends.anchor, ends.focus);
    else sel.setBaseAndExtent(ends.anchor.node, ends.anchor.offset, ends.focus.node, ends.focus.offset);
    // One block, and the browser can move within it: its own extend does the rest.
    if (!across && !stayed) return false;

    let next: DomSelectionPoint | null = stayed ? null : moved;
    // Word's Ctrl+Shift+Down/Up at a paragraph's edge goes on to the next one's end/start.
    if (stayed && granularity !== "lineboundary") {
      const i = story.indexOf(focusBlock);
      const target = story[i + (forward ? 1 : -1)];
      if (!target) { ev.preventDefault(); return true; }    // at the story's edge: nothing more to take
      if (granularity === "line") next = this.pointOnLineOf(target, ends.focus, forward);
      next ??= DocxEditor.blockEdge(target, granularity === "paragraph" ? forward : !forward);
    }
    ev.preventDefault();
    if (next) this.selectBetween(ends.anchor, next);
    return true;
  }

  /** The point in `target`'s first (going down) or last (going up) line nearest the x of `from`. */
  private pointOnLineOf(target: HTMLElement, from: DomSelectionPoint, down: boolean): DomSelectionPoint | null {
    const x = caretRect(from)?.left ?? null;    // (no layout: the caller falls back to the block's edge)
    const box = target.getBoundingClientRect();
    if (x === null || box.height === 0) return null;
    const y = down ? box.top + Math.min(4, box.height / 2) : box.bottom - Math.min(4, box.height / 2);
    const point = caretPointFromClient(document, Math.min(Math.max(x, box.left + 1), box.right - 1), y);
    return point && this.editableBlockOf(point.node) === target ? point : null;
  }

  /** Ctrl+A: the whole story the caret is in (the body, or the header or footer being edited). */
  private selectStory(el: HTMLElement): void {
    const story = this.storyBlocks(el);
    if (story.length < 2) return;    // one block: the browser's own select-all is the same thing
    this.selectBetween(DocxEditor.blockEdge(story[0], false), DocxEditor.blockEdge(story[story.length - 1], true));
  }

  /**
   * Shift+click in another block of the same story extends the selection to the click, from its
   * fixed end. Returns true when it took the click.
   */
  private extendSelectionByClick(event: MouseEvent): boolean {
    if (!event.shiftKey || event.button !== 0) return false;
    const ends = this.selectionEnds();
    const anchorBlock = ends ? this.editableBlockOf(ends.anchor.node) : null;
    if (!ends || !anchorBlock) return false;
    const point = caretPointFromClient(document, event.clientX, event.clientY);
    const target = point ? this.editableBlockOf(point.node) : null;
    if (!point || !target || this.ownerRoot(target) !== this.ownerRoot(anchorBlock)) return false;
    // Within the anchor's own block, and the selection all in it, the browser extends natively.
    if (target === anchorBlock && this.editableBlockOf(ends.focus.node) === anchorBlock) return false;
    event.preventDefault();
    this.selectBetween(ends.anchor, point);
    return true;
  }

  /**
   * Replace a selection that spans blocks with `text` (empty: just delete it), as Word does: the
   * first block keeps its start and takes the text, the blocks between go, the last block's
   * selected start goes and what is left of it joins the first. Lines of `text` after the first
   * become paragraphs of their own. One engine transaction, so one undo step. Returns false, with
   * nothing changed, when the selection does not span blocks or is one {@link crossBlockEditable}
   * leaves to the browser; true when it took the edit, whether or not the engine could make it.
   */
  private replaceCrossBlockSelection(text: string): boolean {
    const blocks = this.crossBlockSelection();
    if (!blocks || !this.crossBlockEditable(blocks)) return false;
    const bridge = this.exports.DocxSessionBridge;
    const targets = this.multiBlockTargets(blocks);
    const first = targets[0], last = targets[targets.length - 1];
    const firstSpan = first.span ?? { start: 0, length: blockContentText(first.block).length };
    const lastSpan = last.span ?? { start: 0, length: blockContentText(last.block).length };
    const firstIndex = this.editableList().indexOf(first.block);
    const lines = text.split(/\r\n|\r|\n/);
    const tx = typeof bridge.BeginTransaction === "function" ? bridge.BeginTransaction(this.handle) : 0;
    let ok = true;
    let caretIndex = firstIndex, caretOffset = firstSpan.start + lines[0].length;
    try {
      // (typing not yet committed in any of them goes in first, so the spans match the session)
      const ids = targets.map((t) => (t.fullId ? this.syncBlock(t.block, t.fullId) : undefined));
      if (ids.some((id) => !id)) throw new Error("a block could not be committed");
      const step = (json: string) => {
        const res = this.parseEdit(json);
        if (!res.success) throw new Error("the engine refused the edit");
        return res;
      };
      let firstId = ids[0]!, lastId = ids[ids.length - 1]!;
      firstId = step(bridge.ReplaceTextAtSpan(this.handle, firstId, firstSpan.start, firstSpan.length, lines[0], ""))
        .modified?.[0]?.id ?? firstId;
      if (lastSpan.length > 0) {
        lastId = step(bridge.ReplaceTextAtSpan(this.handle, lastId, 0, lastSpan.length, "", "")).modified?.[0]?.id ?? lastId;
      }
      if (ids.length > 2) {
        if (typeof bridge.DeleteRange === "function") step(bridge.DeleteRange(this.handle, ids[1]!, lastId));
        else for (const id of ids.slice(1, -1)) step(bridge.DeleteBlock(this.handle, id!));
      }
      let current = step(bridge.MergeParagraphs(this.handle, firstId, lastId)).modified?.[0]?.id ?? firstId;
      // MergeParagraphs puts a space between the two when neither side has one (right for
      // Backspace joining two paragraphs, which Word does the same way); deleting a selection
      // joins its two ends as they are. (Tracked, the merge is only recorded: nothing to take out.)
      // The space is taken out only where the merged paragraph has one: an engine that joins
      // without it leaves the text as it is.
      const tracked = this.options.trackedChanges === TrackedChangeMode.RenderInline;
      const kept = blockContentText(first.block).slice(0, firstSpan.start) + lines[0];
      const rest = blockContentText(last.block).slice(lastSpan.start + lastSpan.length);
      if (!tracked && kept && rest && !/\s$/.test(kept) && !/^\s/.test(rest)
        && this.sessionTextOf(current)?.[kept.length] === " ") {
        current = step(bridge.ReplaceTextAtSpan(this.handle, current, kept.length, 1, "", "")).modified?.[0]?.id ?? current;
      }
      for (const line of lines.slice(1)) {
        const split = step(bridge.SplitParagraph(this.handle, current, caretOffset));
        const second = split.created?.[0]?.id;
        if (!second) throw new Error("the engine made no new paragraph");
        current = line ? step(bridge.ReplaceTextAtSpan(this.handle, second, 0, 0, line, "")).modified?.[0]?.id ?? second : second;
        caretIndex++;
        caretOffset = line.length;
      }
    } catch {
      ok = false;
    } finally {
      if (tx) {
        if (ok) bridge.CommitTransaction!(tx);
        else bridge.RollbackTransaction?.(tx);
      }
    }
    this.extendedSelection = null;
    this.lastCrossBlockSelection = null;
    // A merge or a split can change list numbering and border grouping: the whole flow again.
    this.remount(ok ? caretIndex : firstIndex);
    const target = this.editableList()[ok ? caretIndex : firstIndex];
    if (target) placeCaretAtOffset(target, ok ? caretOffset : firstSpan.start);
    if (ok && target) this.options.onEdit?.({ anchorId: this.anchorIdOf(target) ?? "", unid: target.getAttribute("data-anchor") ?? "" });
    return true;
  }

  /**
   * Whether {@link replaceCrossBlockSelection} edits this selection. Not one reaching into a table
   * (cells are not merged into paragraphs) nor one in a header or footer (a band repaints on its
   * own, not through the body's remount): over those the browser's own editing runs, so a key is
   * never dropped.
   */
  private crossBlockEditable(blocks: HTMLElement[]): boolean {
    return !blocks.some((b) => b.closest("table") || this.isBandBlock(b));
  }

  /** The session's text of block `fullId` (`GetAnchorInfo`'s visibleText without the list number
   *  and the space that visibleText puts before a numbered paragraph's text), or null. */
  private sessionTextOf(fullId: string): string | null {
    try {
      const raw = this.exports.DocxSessionBridge.GetAnchorInfo?.(this.handle, fullId);
      if (raw === undefined) return null;
      const info = JSON.parse(raw) as { visibleText?: unknown; autoNumberPrefix?: unknown };
      const text = info.visibleText;
      if (typeof text !== "string") return null;
      const prefix = typeof info.autoNumberPrefix === "string" ? info.autoNumberPrefix : "";
      if (!prefix) return text;
      if (text === prefix) return "";
      return text.startsWith(prefix + " ") ? text.slice(prefix.length + 1) : null;
    } catch {
      return null;
    }
  }

  /** The selected text of a selection spanning blocks, a line for each block. */
  private crossBlockText(blocks: HTMLElement[]): string {
    return this.multiBlockTargets(blocks).map((t) => {
      const text = blockContentText(t.block);
      return t.span ? text.slice(t.span.start, t.span.start + t.span.length) : text;
    }).join("\n");
  }

  /** Cut of a selection spanning blocks: the browser's cut would remove the focused block's part
   *  only. The clipboard gets the selection (as copy gives it), then the selection goes. */
  private readonly onCut = (ev: ClipboardEvent): void => {
    if (this.closed || ev.defaultPrevented || !ev.clipboardData) return;
    const blocks = this.crossBlockSelection();
    if (!blocks || !this.crossBlockEditable(blocks)) return;
    ev.preventDefault();
    const range = window.getSelection()!.getRangeAt(0);
    const holder = document.createElement("div");
    holder.append(range.cloneContents());
    ev.clipboardData.setData("text/html", holder.innerHTML);
    ev.clipboardData.setData("text/plain", this.crossBlockText(blocks));
    this.replaceCrossBlockSelection("");
  };

  /** Paste over a selection spanning blocks: the selection goes, and the clipboard's text takes
   *  its place, a paragraph for each line. */
  private readonly onPasteOverBlocks = (ev: ClipboardEvent): void => {
    if (this.closed || ev.defaultPrevented || !ev.clipboardData) return;
    const blocks = this.crossBlockSelection();
    if (!blocks || !this.crossBlockEditable(blocks)) return;
    ev.preventDefault();
    this.replaceCrossBlockSelection(ev.clipboardData.getData("text/plain"));
  };

  /**
   * Keys over a selection spanning blocks, and the keys that make one (see
   * {@link extendSelectionByKey}, {@link selectStory}). Returns true when it handled the key.
   */
  private crossBlockKey(el: HTMLElement, ev: KeyboardEvent): boolean {
    if (ev.isComposing) return false;
    if (this.extendSelectionByKey(el, ev)) return true;
    const ctrl = ev.ctrlKey || ev.metaKey;
    if (ctrl && !ev.altKey && !ev.shiftKey && (ev.key ?? "").toLowerCase() === "a") {
      // (a story whose selection could not be typed over keeps the browser's own select-all)
      const story = this.storyBlocks(el);
      if (story.length < 2 || !this.crossBlockEditable(story)) return false;
      ev.preventDefault();
      this.selectStory(el);
      return true;
    }
    const blocks = this.crossBlockSelection();
    if (!blocks || !this.crossBlockEditable(blocks)) return false;
    const printable = !ctrl && !ev.altKey && typeof ev.key === "string" && [...ev.key].length === 1;
    let text: string | null = null;
    if (printable) text = ev.key;
    else if ((ev.key === "Backspace" || ev.key === "Delete") && !ev.altKey) text = "";
    else if (ev.key === "Enter" && !ctrl && !ev.altKey) text = ev.shiftKey ? "" : "\n";
    if (text === null) return false;
    ev.preventDefault();
    const shiftEnter = ev.key === "Enter" && ev.shiftKey;
    this.replaceCrossBlockSelection(text);
    if (shiftEnter && this.activeBlock) this.insertLineBreakAtCaret(this.activeBlock);
    return true;
  }

  /** Editable blocks the current selection covers, in document order. Uses Range.comparePoint
   *  (robust to a selection boundary that normalized onto a wrapper element rather than a block
   *  or text node — Range.intersectsNode misses the end block at a `(block, childCount)` boundary).
   *  A collapsed or single-block selection yields just the active block. */
  private selectedBlocks(): HTMLElement[] {
    let sel = typeof window !== "undefined" ? window.getSelection() : null;
    if ((!sel || sel.rangeCount === 0 || sel.isCollapsed) && this.lastCrossBlockSelection) {
      this.restoreCrossBlockSelection();
      sel = typeof window !== "undefined" ? window.getSelection() : null;
    }
    // Enumerate the ACTIVE block's own root — a band's story container, else the body edit root —
    // so a selection can never span a band and the body (different OOXML parts).
    const root = this.activeBlock ? this.ownerRoot(this.activeBlock) : this.editRoot;
    const all = Array.from(
      root.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    );
    if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
      const range = sel.getRangeAt(0);
      const hit = all.filter((b) => {
        try {
          const startsAfterEnd = range.comparePoint(b, 0) > 0; // block begins after the selection ends
          const endsBeforeStart = range.comparePoint(b, b.childNodes.length) < 0; // block ends before it starts
          return !startsAfterEnd && !endsBeforeStart;
        } catch {
          return false;
        }
      });
      if (hit.length > 1) return hit;
    }
    return this.activeBlock ? [this.activeBlock] : [];
  }

  /** The selection's span within `block`, clipped to the block (for inline ops across blocks):
   *  the first block runs selection-start→end-of-block, middle blocks are whole, the last block
   *  runs start-of-block→selection-end. Returns null for a whole-block apply. */
  private blockSpanForSelection(block: HTMLElement): { start: number; length: number } | null {
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null;
    const range = sel.getRangeAt(0);
    const hasStart = block.contains(range.startContainer);
    const hasEnd = block.contains(range.endContainer);
    if (hasStart && hasEnd) return selectionSpanIn(block);
    // Partial/whole slices are normalized into the committed run-text space for the same reason
    // selectionSpanIn is (see trimmedSpan): a session op rejects a span past the committed end.
    const contentLen = blockContentText(block).length;
    if (hasStart) {
      const start = contentOffsetOf(block, range.startContainer, range.startOffset);
      return trimmedSpan(block, { start, length: Math.max(0, contentLen - start) });
    }
    if (hasEnd) {
      const end = contentOffsetOf(block, range.endContainer, range.endOffset);
      return trimmedSpan(block, { start: 0, length: end });
    }
    return trimmedSpan(block, { start: 0, length: contentLen }); // fully-spanned middle block
  }

  /** Apply an inline ApplyFormat op to each block's slice of the selection, then reconcile
   *  the DOM incrementally (see {@link finishMultiBlockOp} — a full remount costs a whole-
   *  document convert, seconds on a large doc, where per-block swaps are ~10 ms each).
   *  Returns false (caller falls back to the single-block path) for a 1-block selection. */
  private applyInlineOpAcrossBlocks(blocks: HTMLElement[], op: object): boolean {
    if (blocks.length <= 1) return false;
    const targets = this.multiBlockTargets(blocks);
    for (const t of targets) {
      if (!t.fullId || (t.span && t.span.length === 0)) continue;
      const synced = this.syncBlock(t.block, t.fullId);
      if (!synced) continue;
      const res = this.parseEdit(this.exports.DocxSessionBridge.ApplyFormat(
        this.handle, synced, t.span ? JSON.stringify(t.span) : "", JSON.stringify(op),
      ));
      if (res.success) t.res = res;
    }
    this.finishMultiBlockOp(targets, false);
    return true;
  }

  /** Apply a whole-block (paragraph-level) op to each selected block, then reconcile the DOM
   *  incrementally ({@link finishMultiBlockOp}). `forceRemount` is for ops whose rendering
   *  needs whole-document context (e.g. border-div regrouping after clearBorders). Returns
   *  false for a 1-block selection (caller uses the single-block path). */
  private applyParagraphOpAcrossBlocks(
    blocks: HTMLElement[],
    run: (fullId: string) => string,
    forceRemount = false,
  ): boolean {
    if (blocks.length <= 1) return false;
    const targets = this.multiBlockTargets(blocks);
    for (const t of targets) {
      if (!t.fullId) continue;
      const synced = this.syncBlock(t.block, t.fullId);
      if (!synced) continue;
      const res = this.parseEdit(run(synced));
      if (res.success) t.res = res;
    }
    this.finishMultiBlockOp(targets, forceRemount);
    return true;
  }

  /** Snapshot each selected block's identity + selection slice BEFORE any session op runs
   *  (ops never mutate the DOM, so spans captured here stay valid until the swap phase). */
  private multiBlockTargets(blocks: HTMLElement[]): Array<{
    block: HTMLElement;
    unid: string | null;
    fullId: string | undefined;
    span: { start: number; length: number } | null;
    res: EditResultLite | null;
  }> {
    return blocks.map((b) => {
      const unid = b.getAttribute("data-anchor");
      return {
        block: b,
        unid,
        fullId: this.anchorIdOf(b),
        span: this.blockSpanForSelection(b),
        res: null,
      };
    });
  }

  /**
   * Reconcile the DOM after a multi-block op. Fidelity-identical to the single-block path by
   * construction: each edited block is swapped for its own session-attached single-block
   * render — exactly what format()/setFontSize()/applyParagraphFormat() do for one block —
   * so a multi-block apply is N single-block applies, not one whole-document re-render
   * (which froze the UI for the full-document convert time on every ribbon action). Falls
   * back to ONE full remount when the op touched a list item (numbering continuation needs
   * whole-document context) or the caller forced it. Restores the cross-block selection so
   * consecutive ribbon actions (center, then bold) keep targeting the same range.
   */
  private finishMultiBlockOp(
    targets: Array<{
      block: HTMLElement;
      unid: string | null;
      span: { start: number; length: number } | null;
      res: EditResultLite | null;
    }>,
    forceRemount: boolean,
    /** In page view too, swap each block in place (an op that keeps every block's length, as
     *  Change Case does) unless one is split across two pages. */
    swapInPageView = false,
  ): void {
    const edited = targets.filter((t) => t.res && t.unid);
    if (edited.length === 0) return;
    const splitAcrossPages = (unid: string) => this.editRoot.querySelectorAll(`[data-anchor="${CSS.escape(unid)}"]`).length > 1;
    const pageViewSwap = swapInPageView && !!this.options.paginated && !edited.some((t) => splitAcrossPages(t.unid!));
    // Paginated mode: a multi-block format change can alter block heights, and page boxes
    // need a reflow — keep the (pre-existing) full remount there until M4 lands a scoped
    // re-paginate. Continuous mode reconciles incrementally.
    if (forceRemount || (this.options.paginated && !pageViewSwap) || edited.some((t) => this.affectsList(t.res!))) {
      this.refreshAfter(edited[0].block, -1, false);
      return;
    }
    const swapped: Array<{ el: HTMLElement; span: { start: number; length: number } | null }> = [];
    for (const t of edited) {
      const fresh = this.swapBlock(t.block, t.unid!, t.res!.modified?.[0]);
      swapped.push({ el: fresh ?? t.block, span: t.span });
    }
    const first = swapped[0];
    const last = swapped[swapped.length - 1];
    if (!first.el.isConnected || !last.el.isConnected) return;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel) return;
    const startOff = first.span?.start ?? 0;
    const endOff = last.span ? last.span.start + last.span.length : blockContentText(last.el).length;
    const startPos = contentPositionIn(first.el, startOff);
    const endPos = contentPositionIn(last.el, endOff);
    try {
      const range = document.createRange();
      range.setStart(startPos.node, startPos.offset);
      range.setEnd(endPos.node, endPos.offset);
      sel.removeAllRanges();
      sel.addRange(range);
    } catch {
      /* selection restore is best-effort — the edits themselves are already committed */
    }
  }

  /**
   * Toggle (or set) an inline format on the current selection in the active block.
   * A selection spanning multiple blocks applies to each. With no selection, applies to
   * the whole paragraph. Routes through DocxSession (`ApplyFormat`) so it is lossless and
   * supports underline/strike, not just markdown.
   */
  format(key: FormatKey, value?: boolean): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (blocks.length > 1) {
      const on0 = value ?? !selectionHasFormat(key, blocks[0]);
      const multiOp =
        key === "superscript" || key === "subscript"
          ? { vertAlign: on0 ? key : "" }
          : { [key]: on0 };
      if (this.applyInlineOpAcrossBlocks(blocks, multiOp)) return;
    }
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;

    const span = selectionSpanIn(block);
    const on = value ?? !selectionHasFormat(key, block);
    // Super/subscript map to the single-valued w:vertAlign; the rest are boolean toggles.
    const op =
      key === "superscript" || key === "subscript"
        ? { vertAlign: on ? key : "" }
        : { [key]: on };
    fullId = this.syncBlock(block, fullId); // don't clobber uncommitted typing
    if (!fullId) return;
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyFormat(
        this.handle,
        fullId,
        span ? JSON.stringify(span) : "",
        JSON.stringify(op),
      ),
    );
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, this.blockIndex(block), false); return; }
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh && span) selectRange(fresh, span.start, span.length);
    else fresh?.focus();
  }

  /**
   * Set the font size (in points) of the current selection in the active block; with no
   * selection, applies to the whole paragraph. `pts <= 0` clears the explicit size. Routes
   * through DocxSession `ApplyFormat` (`w:sz`), so it is lossless and survives save.
   */
  setFontSize(pts: number): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (blocks.length > 1 && this.applyInlineOpAcrossBlocks(blocks, { fontSizePts: pts })) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    // Use the live selection; if the font-size combobox stole focus and collapsed it, fall back to
    // the last real selection cached for this block so a sub-range still sizes (finding 3).
    let span = selectionSpanIn(block);
    if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyFormat(
        this.handle,
        fullId,
        span ? JSON.stringify(span) : "",
        JSON.stringify({ fontSizePts: pts }),
      ),
    );
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, this.blockIndex(block), false); return; }
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh && span) selectRange(fresh, span.start, span.length);
    else fresh?.focus();
  }

  /**
   * Set the font family of the current selection in the active block; with no selection,
   * applies to the whole paragraph. `""` clears the explicit font (inherits the style/default).
   * Routes through DocxSession `ApplyFormat` (`w:rFonts`), so it is lossless and survives save.
   * Multi-block + last-selection plumbing matches {@link setFontSize} (a focus-stealing font
   * dropdown still applies to the real sub-range).
   */
  setFontFamily(name: string): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (blocks.length > 1 && this.applyInlineOpAcrossBlocks(blocks, { fontFamily: name })) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    let span = selectionSpanIn(block);
    if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyFormat(
        this.handle,
        fullId,
        span ? JSON.stringify(span) : "",
        JSON.stringify({ fontFamily: name }),
      ),
    );
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, this.blockIndex(block), false); return; }
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh && span) selectRange(fresh, span.start, span.length);
    else fresh?.focus();
  }

  /** Set paragraph alignment (left/center/right/justify) on the active block. */
  setAlignment(alignment: EditorAlignment): void {
    this.applyParagraphFormat({ alignment });
  }

  /**
   * Insert an S-1-style horizontal rule (an empty paragraph with a bottom border) after the
   * active block. `weight` is the rule thickness in eighths of a point (default 12 ≈ 1.5pt).
   * Re-renders fully (a new block needs whole-document context to lay out).
   */
  insertHorizontalRule(weight = 12, style = "single", position: "above" | "below" = "below"): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.InsertHorizontalRule(
        this.handle,
        fullId,
        position === "above" ? "before" : "after",
        JSON.stringify({ style, size: weight, color: "auto" }),
      ),
    );
    if (!res.success) return;
    // Remount from the active block's index re-renders the new rule whether it landed just
    // above (at idx) or just below (at idx+1) the active block. Forced: a rule is a bordered
    // paragraph, and border-div grouping is whole-document render context.
    this.refreshAfter(block, idx, false, /* forceRemount */ true);
  }

  /**
   * Insert a `rows`×`cols` table after the active block. `options.cellContents` (row-major
   * markdown) seeds the cells, `options.borderless` makes an invisible layout table, and
   * `options.cellAlignment` aligns every cell. Re-renders fully (tables need document context).
   */
  insertTable(
    rows: number,
    cols: number,
    options?: {
      borderless?: boolean;
      cellContents?: string[];
      cellAlignment?: EditorAlignment;
      columnWidths?: number[];
    },
  ): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    // If the caret is on an empty paragraph (not a table cell), insert the table BEFORE it so the
    // empty paragraph becomes the editable line BELOW the table — no stray blank line above it, and
    // a reachable paragraph below (S-1 smoke-test findings 2 + 4). Otherwise insert after.
    const emptyHere =
      !block.closest("table") && blockContentText(block).replace(/[\s ]+/g, "").length === 0;
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.InsertTable(
        this.handle,
        fullId,
        emptyHere ? "before" : "after",
        rows,
        cols,
        options ? JSON.stringify(options) : "",
      ),
    );
    if (!res.success) return;
    this.refreshAfter(block, idx, false);
  }

  // ─── Table row / column editing (active block must be inside a table cell) ──────────

  /** Run a table-structure op on the active cell's canonical tc anchor and re-render. */
  private tableEdit(run: (cellAnchor: string) => string): void {
    const block = this.activeBlock;
    const cell = block?.closest<HTMLElement>("td, th");
    if (this.closed || !block || !cell || !block.closest("table")) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    const paragraphId = this.anchorIdOf(block);
    const cellId = this.anchorIdOf(cell);
    if (!paragraphId || !cellId) return;
    const idx = this.blockIndex(block);
    if (!this.syncBlock(block, paragraphId)) return; // flush uncommitted cell text first
    const res = this.parseEdit(run(cellId));
    if (!res.success) return;
    this.refreshAfter(block, idx, false);
  }

  /** Insert a row above/below the active cell's row. No-op outside a table. */
  insertTableRow(where: "above" | "below"): void {
    this.tableEdit((a) =>
      this.exports.DocxSessionBridge.InsertTableRow(this.handle, a, where === "above" ? "before" : "after"),
    );
  }

  /** Insert a column left/right of the active cell's column. No-op outside a table. */
  insertTableColumn(where: "left" | "right"): void {
    this.tableEdit((a) =>
      this.exports.DocxSessionBridge.InsertTableColumn(this.handle, a, where === "left" ? "before" : "after"),
    );
  }

  /** Delete the active cell's row (deleting the last row removes the table). No-op outside a table. */
  deleteTableRow(): void {
    this.tableEdit((a) => this.exports.DocxSessionBridge.DeleteTableRow(this.handle, a));
  }

  /** Delete the active cell's column (deleting the last column removes the table). No-op outside a table. */
  deleteTableColumn(): void {
    this.tableEdit((a) => this.exports.DocxSessionBridge.DeleteTableColumn(this.handle, a));
  }

  /**
   * Indent/outdent the active block. On a LIST item this changes the list NESTING LEVEL
   * (`SetListLevel`) so numbering nests (e.g. 1, 2 → a sub-level) rather than the item just
   * shifting sideways with flat numbering. On a plain paragraph it adjusts the left indent by
   * `deltaTwips` (default ±720 = 0.5"), clamped at 0.
   */
  indent(deltaTwips = 720): void {
    if (this.activeBlock && isListBlock(this.activeBlock)) {
      this.setListLevel(deltaTwips >= 0 ? 1 : -1);
      return;
    }
    this.applyParagraphFormat({ indentDelta: deltaTwips });
  }

  /** Change the active list item's nesting level by `delta` (+1 deeper, −1 shallower). */
  private setListLevel(delta: number): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const res = this.parseEdit(this.exports.DocxSessionBridge.SetListLevel(this.handle, fullId, delta));
    if (!res.success) return;
    // A level change ripples through the whole list's numbering — re-render with full document
    // context (sibling numbers shift without sibling XML changing), keeping the caret in place.
    this.refreshAfter(block, idx, false, /* forceRemount */ true);
  }

  /** Toggle (or set) page-break-before on the active block. */
  pageBreakBefore(value = true): void {
    this.applyParagraphFormat({ pageBreakBefore: value });
  }

  /**
   * Toggle the active block between a bullet/numbered list item and a plain paragraph.
   * Clicking the same kind it already is removes the list; any other state applies the kind.
   */
  toggleList(kind: "bullet" | "decimal"): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;

    let membership: { format?: string } | null = null;
    try {
      membership = JSON.parse(this.exports.DocxSessionBridge.GetListMembership(this.handle, fullId));
    } catch { /* treat as not-a-list */ }
    const isThisKind =
      !!membership && typeof membership.format === "string" &&
      membership.format.toLowerCase().startsWith(kind === "bullet" ? "bullet" : "decimal");

    const idx = this.blockIndex(block); // capture before the op
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const tx = this.beginListTransaction();
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyListFormat(this.handle, fullId, isThisKind ? "none" : kind),
    );
    if (res.success && !membership && !isThisKind) this.applyListDefaults(res.modified?.[0]?.id ?? fullId, kind);
    this.endListTransaction(tx);
    if (!res.success) return;
    // Numbering continuation across the list needs whole-document context — re-render fully
    // (sibling numbers shift without sibling XML changing).
    this.refreshAfter(block, idx, false, /* forceRemount */ true);
  }

  /** Clear all paragraph borders (e.g. remove an inserted horizontal rule) on the active block —
   *  or every block in a multi-block selection. The engine/wire already accept `clearBorders`;
   *  this surfaces it on the editor so an HR border is removable (S-1 smoke-test finding 1b). */
  clearParagraphBorders(): void {
    this.applyParagraphFormat({ clearBorders: true });
  }

  /**
   * Delete the active block (e.g. a stray empty paragraph left above/below a table). Routes
   * through DocxSession `DeleteBlock` + re-render, focusing the previous block. No-op when the
   * caret is inside a table (remove cells via the table toolbar's delete row/column instead) and
   * no-op when it is the only editable block (don't empty the document). Closes the S-1
   * smoke-test "no block-delete affordance" gap.
   */
  deleteBlock(): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    if (block.closest("table")) return; // cells are removed via the table toolbar, not here
    if (this.editableList().length <= 1) return; // never delete the last editable block
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    const fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    const res = this.parseEdit(this.exports.DocxSessionBridge.DeleteBlock(this.handle, fullId));
    if (!res.success) return;
    this.refreshAfter(block, Math.max(0, idx - 1), true);
  }

  /**
   * Cite a new footnote from the caret position in the active body block. The note definition is
   * created (writing the whole Word scaffold — part, reserved separator notes, settings
   * declaration, styles — on a document that has none yet) and its body renders as ordinary
   * editable `data-anchor` blocks in the notes section, so editing it afterwards needs no new op.
   *
   * Body blocks only: Word disallows a note reference inside a header/footer story or inside
   * another note, and the session rejects those with `AnchorWrongKind`. Reconciliation creates
   * the first notes section when needed and renumbers every affected citation in place.
   */
  insertFootnote(markdown = "New footnote."): void {
    this.insertNote("footnote", markdown);
  }

  /** Cite a new endnote from the caret — see {@link insertFootnote}; writes the endnotes part. */
  insertEndnote(markdown = "New endnote."): void {
    this.insertNote("endnote", markdown);
  }

  private insertNote(kind: "footnote" | "endnote", markdown: string): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    // A note reference is legal only in the main story: not in a header/footer band, and not
    // inside an existing note's body (both render as editable blocks here).
    if (this.isBandBlock(block) || block.closest(".footnotes, .endnotes")) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    // Offset first: syncBlock re-renders the block and would drop the live selection.
    const raw = caretOffsetIn(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const offset = trimmedSplitOffset(block, raw ?? (block.textContent ?? "").length);
    const bridge = this.exports.DocxSessionBridge;
    const call = kind === "footnote" ? bridge.InsertFootnote : bridge.InsertEndnote;
    if (!call) return; // bridge predates note authoring
    const res = this.parseEdit(call.call(bridge, this.handle, fullId, offset, markdown));
    if (!res.success) return;
    this.reconcile(idx, false);
  }

  /**
   * Add a native Word comment on the current selection in the active block — or on the whole
   * block when the selection is collapsed (a null span comments the paragraph). The definition,
   * threading parts, and body-side range markers are written by the session's `AddComment`
   * (issue #300); the re-render shows the marker chrome the converter already draws. The
   * annotation type legal review runs on, finally authorable from the shipped surface
   * (issue #580).
   */
  addComment(
    markdown = "New comment.",
    author = this.options.commentAuthor,
    target?: CommentTarget,
  ): CommentListEntry | null {
    let block = target?.block.isConnected ? target.block : this.activeBlock;
    let span = target ? target.span : null;
    if (target && !target.block.isConnected) {
      // The draft's block was swapped since it was captured — the click that posts a draft
      // blurs and commits the paragraph. Find the same paragraph by anchor rather than taking
      // whatever block is active now, and keep the span only while the text it was measured
      // on is unchanged; otherwise comment the whole paragraph rather than the wrong characters.
      const live = target.anchor === undefined ? null : this.blockByAnchor(target.anchor);
      if (live) block = live;
      if (!live || target.text === undefined || blockContentText(live) !== target.text) span = null;
    }
    if (this.closed || !block) return null;
    const bridge = this.exports.DocxSessionBridge;
    if (!bridge.AddComment) return null; // bridge predates comment authoring
    let fullId = this.anchorIdOf(block);
    if (!fullId) return null;
    const idx = this.blockIndex(block);
    // Span first: syncBlock re-renders the block and would drop the live selection. A target
    // captured earlier (the gutter's draft bubble) wins over whatever the selection is now.
    if (!target) span = selectionSpanIn(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return null;
    const res = this.parseEdit(
      bridge.AddComment(
        this.handle, fullId, span ? JSON.stringify(span) : "", author, this.commentInitialsFor(author),
        commentTimestamp(), markdown,
      ),
    );
    if (!res.success) return null;
    const created = res.created?.find((a) => a.kind === "cmt")?.id ?? null;
    // Comment markup lands inside the host paragraph, so it re-renders like any text edit; the
    // gutter then picks the new highlight up on its next layout.
    if (this.isBandBlock(block)) this.refreshAfter(block, idx, false);
    else this.reconcile(idx, false);
    const entry = created ? this.listComments().find((c) => c.anchorId === created) ?? null : null;
    return entry;
  }

  /** Reply to a thread root (or any comment) as a native Word reply (`commentsExtended`). */
  addCommentReply(parentAnchorId: string, markdown: string, author = this.options.commentAuthor): boolean {
    if (this.closed) return false;
    const bridge = this.exports.DocxSessionBridge;
    if (!bridge.AddCommentReply) return false;
    const res = this.parseEdit(
      bridge.AddCommentReply(this.handle, parentAnchorId, author, this.commentInitialsFor(author),
        commentTimestamp(), markdown),
    );
    if (!res.success) return false;
    // A reply adds a reference run beside the parent's, so the host paragraph re-renders.
    this.reconcile();
    return true;
  }

  /** Replace a comment's body text; author/date are preserved. */
  updateComment(commentAnchorId: string, markdown: string): boolean {
    if (this.closed) return false;
    const bridge = this.exports.DocxSessionBridge;
    if (!bridge.UpdateComment) return false;
    return this.parseEdit(bridge.UpdateComment(this.handle, commentAnchorId, markdown)).success;
  }

  /** Delete a comment: the definition and its range markers everywhere. */
  removeComment(commentAnchorId: string): boolean {
    if (this.closed) return false;
    const bridge = this.exports.DocxSessionBridge;
    if (!bridge.RemoveComment) return false;
    const res = this.parseEdit(bridge.RemoveComment(this.handle, commentAnchorId));
    if (!res.success) return false;
    this.reconcile();
    return true;
  }

  /** What "New comment" would comment on right now: the active block and the selection in it. */
  commentTarget(): CommentTarget | null {
    const block = this.activeBlock;
    if (this.closed || !block || !block.isConnected) return null;
    let span = selectionSpanIn(block);
    const unid = block.getAttribute("data-anchor");
    if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
    return { block, span, anchor: this.anchorIdOf(block), text: blockContentText(block) };
  }

  /** The live block for a full anchor id, or null when no mounted block resolves to it. */
  private blockByAnchor(anchor: string): HTMLElement | null {
    for (const el of Array.from(this.container.querySelectorAll<HTMLElement>("[data-anchor]"))) {
      if (this.anchorIdOf(el) === anchor) return el;
    }
    return null;
  }

  /** Open a draft comment bubble beside the selection (Word's "New Comment"). */
  beginComment(): boolean {
    return this.gutter?.beginDraft() ?? false;
  }

  cancelComment(): void {
    this.gutter?.cancelDraft();
  }

  /** Activate a thread by `cmt` anchor id or numeric comment id (null clears). */
  activateComment(id: string | null): void {
    this.gutter?.setActive(id, { scrollBubble: true, scrollAnchor: true });
  }

  /** Step to the next (+1) / previous (−1) thread in document order. */
  stepComment(direction: 1 | -1): CommentListEntry | null {
    return this.gutter?.step(direction) ?? null;
  }

  /** The active thread root's anchor id, or null. */
  get activeComment(): string | null {
    return this.gutter?.active ?? null;
  }

  /** Show or hide the comment gutter (the markup stays in the document). */
  showComments(visible: boolean): void {
    this.gutter?.setVisible(visible);
  }

  get commentsVisible(): boolean {
    return this.gutter?.isVisible ?? false;
  }

  /** Force the gutter to lay out now (tests). */
  layoutComments(): void {
    this.gutter?.layout();
  }

  /** The document's native comment threads (session truth), for review UIs — flat entries
   *  with `parentAnchorId` linking replies and `resolved` carrying thread state. */
  listComments(): CommentListEntry[] {
    if (this.closed) return [];
    const bridge = this.exports.DocxSessionBridge;
    if (!bridge.ListComments) return [];
    try {
      return JSON.parse(bridge.ListComments(this.handle)) as CommentListEntry[];
    } catch {
      return [];
    }
  }

  /** Resolve or reopen a comment thread by its `cmt` anchor id (from {@link listComments}).
   *  Resolution is thread metadata (`commentsExtended`), not body markup, so no re-render is
   *  needed — a review UI re-reads {@link listComments} for the new state. */
  setCommentResolved(commentAnchorId: string, resolved: boolean): boolean {
    if (this.closed) return false;
    const bridge = this.exports.DocxSessionBridge;
    if (!bridge.SetCommentResolved) return false;
    return this.parseEdit(bridge.SetCommentResolved(this.handle, commentAnchorId, resolved)).success;
  }

  /**
   * Word's Paragraph dialog in one call: every field given is set on the active paragraph, or on
   * every paragraph of a multi-paragraph selection, as ONE undo step; fields left out stay as
   * they are. Indents and spacing are twips (1440 an inch, 20 a point); see
   * {@link EditorParagraphFormat}. Returns false when there is no paragraph to format.
   */
  setParagraphFormat(op: EditorParagraphFormat): boolean {
    if (this.closed || !this.activeBlock) return false;
    const clean: EditorParagraphFormat = {};
    for (const [key, value] of Object.entries(op) as Array<[keyof EditorParagraphFormat, unknown]>) {
      if (value !== undefined && value !== null) (clean as Record<string, unknown>)[key] = value;
    }
    if (Object.keys(clean).length === 0) return true;
    const bridge = this.exports.DocxSessionBridge;
    // Several paragraphs are several engine ops: one transaction makes them one undo step.
    const tx = this.selectedBlocks().length > 1 && typeof bridge.BeginTransaction === "function"
      ? bridge.BeginTransaction(this.handle) : 0;
    try {
      this.applyParagraphFormat(clean);
    } finally {
      if (tx) bridge.CommitTransaction!(tx);
    }
    return true;
  }

  /** The effective paragraph formatting of each paragraph the selection covers (the caret's
   *  paragraph when nothing is selected) — what a Paragraph dialog shows; it leaves a field
   *  blank where they differ. Empty when there is no paragraph. */
  paragraphFormats(): EditorParagraphFormat[] {
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !this.activeBlock || typeof bridge.GetFormatting !== "function") return [];
    const out: EditorParagraphFormat[] = [];
    for (const block of this.selectedBlocks()) {
      const fullId = this.anchorIdOf(block);
      if (!fullId) continue;
      try {
        const parsed = JSON.parse(bridge.GetFormatting(this.handle, fullId)) as FormattingInspection | null;
        if (parsed && typeof parsed === "object" && parsed.effectiveParagraph) out.push(asEditorParagraphFormat(parsed.effectiveParagraph));
      } catch {
        // A block the engine cannot inspect (a table, a field result) has no paragraph format.
      }
    }
    return out;
  }

  private applyParagraphFormat(op: EditorParagraphFormat & {
    indentDelta?: number;
    clearBorders?: boolean;
    suppressLineNumbers?: boolean;
  }): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    // A border change regroups the wrapping border <div>s, and a bordered paragraph's wrapper
    // carries its indentation and outer spacing — both are whole-document context. (A bordered
    // table cell is no such wrapper.)
    // In page view the page-keeping rules decide where pages break, which only a new layout of the
    // whole flow shows (a swapped block keeps its old page).
    const repaginates = !!this.options.paginated && (op.pageBreakBefore !== undefined || op.keepNext !== undefined
      || op.keepLines !== undefined || op.widowControl !== undefined);
    // Line numbers are counted across the page view: suppressing one paragraph's renumbers the rest.
    const rewraps = !!op.clearBorders || repaginates || op.suppressLineNumbers !== undefined
      || [block, ...blocks].some((b) => b.parentElement?.tagName === "DIV" && inBorderWrapper(b));
    if (
      blocks.length > 1 &&
      this.applyParagraphOpAcrossBlocks(
        blocks,
        (id) => this.exports.DocxSessionBridge.SetParagraphFormat(this.handle, id, JSON.stringify(op)),
        rewraps,
      )
    )
      return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.SetParagraphFormat(this.handle, fullId, JSON.stringify(op)),
    );
    if (!res.success) return;
    // A single-block swap can't restructure or restyle the wrapping border <div> — re-render fully
    // (like list edits) so the wrapper appears, disappears or takes its new spacing cleanly.
    if (this.affectsList(res) || rewraps) { this.refreshAfter(block, idx, false, true); return; }
    this.swapBlock(block, unid, res.modified?.[0])?.focus();
  }

  /**
   * Where the caret's paragraph sits on screen and its indents and tab stops — what a ruler
   * draws. Null without an editable paragraph (or with a right-to-left one, which the ruler does
   * not draw yet).
   */
  rulerState(): EditorRulerState | null {
    const block = this.rulerBlock();
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !block || typeof bridge.GetFormatting !== "function") return null;
    this.rulerAnchor = block.getAttribute("data-anchor");
    const fullId = this.anchorIdOf(block);
    if (!fullId) return null;
    let format: EditorParagraphFormat;
    let tabs: TabStopsInfo | null = null;
    try {
      const parsed = JSON.parse(bridge.GetFormatting(this.handle, fullId)) as FormattingInspection | null;
      if (!parsed?.effectiveParagraph) return null;
      format = asEditorParagraphFormat(parsed.effectiveParagraph);
      if (typeof bridge.GetTabStops === "function")
        tabs = JSON.parse(bridge.GetTabStops(this.handle, fullId)) as TabStopsInfo | null;
    } catch {
      return null;
    }
    if (typeof getComputedStyle === "function" && getComputedStyle(block).direction === "rtl") return null;
    const rect = block.getBoundingClientRect();
    // Screen pixels per CSS pixel, whatever scales the view (a transform, CSS zoom): a twip is a
    // fifteenth of a CSS pixel (1/20 pt, 96 px to 72 pt).
    const screenScale = block.offsetWidth > 0 ? rect.width / block.offsetWidth : 1;
    const pxPerTwip = screenScale / 15;
    const leftIndent = format.leftIndent ?? 0;
    const rightIndent = format.rightIndent ?? 0;
    const originLeft = rect.left - leftIndent * pxPerTwip;
    const originRight = rect.right + rightIndent * pxPerTwip;

    // The page: the page box in page view; in continuous view the section wrapper, which is
    // either the sheet itself or only its text column.
    const section = block.closest<HTMLElement>("[data-section-index]")
      ?? (this.editRoot.querySelector<HTMLElement>("[data-section-index]") ?? this.editRoot);
    const dims = parseSectionDimensions(section);
    const pxPerPt = pxPerTwip * 20;
    const page = block.closest<HTMLElement>(".page-box");
    let pageLeft: number;
    if (page) {
      pageLeft = page.getBoundingClientRect().left;
    } else {
      const box = section.getBoundingClientRect();
      const sheet = Math.abs(box.width - dims.pageWidth * pxPerPt) < Math.abs(box.width - dims.contentWidth * pxPerPt);
      pageLeft = sheet ? box.left : box.left - dims.marginLeft * pxPerPt;
    }
    const pageRight = pageLeft + dims.pageWidth * pxPerPt;
    const textLeft = pageLeft + dims.marginLeft * pxPerPt;
    const textRight = pageRight - dims.marginRight * pxPerPt;
    const cell = block.closest<HTMLElement>("td, th");
    const inSectionColumn = !cell && !this.isBandBlock(block)
      && Math.abs(originLeft - textLeft) < 1.5 && Math.abs(originRight - textRight) < 1.5;
    const cellBox = cell?.getBoundingClientRect();
    let table: EditorRulerState["table"];
    const tableInfo = cell ? this.tablePropertiesOf(cell) : null;
    const row = cell?.closest("tr");
    if (tableInfo && row) {
      const cells = Array.from(row.children).filter((c): c is HTMLElement => c.tagName === "TD" || c.tagName === "TH");
      const boundaries: { x: number; grid: number }[] = [];
      let grid = 0;
      cells.forEach((c, i) => {
        if (i > 0) boundaries.push({ x: c.getBoundingClientRect().left, grid });
        grid += Math.max(1, (c as HTMLTableCellElement).colSpan || 1);
      });
      const last = cells[cells.length - 1];
      if (last) boundaries.push({ x: last.getBoundingClientRect().right, grid: tableInfo.columnWidthsTwips.length });
      table = { boundaries, widthsTwips: tableInfo.columnWidthsTwips };
    }
    return {
      pageLeft, pageRight, textLeft, textRight, originLeft, originRight, pxPerTwip, inSectionColumn,
      boundsLeft: cellBox ? cellBox.left : pageLeft,
      boundsRight: cellBox ? cellBox.right : pageRight,
      ...(table ? { table } : {}),
      marginLeftTwips: Math.round(dims.marginLeft * 20),
      marginRightTwips: Math.round(dims.marginRight * 20),
      leftIndent,
      rightIndent,
      firstLine: format.hangingIndent ? -format.hangingIndent : (format.firstLineIndent ?? 0),
      tabs: tabs?.stops ?? [],
      defaultTabStopTwips: tabs?.defaultTabStopTwips ?? 720,
    };
  }

  /** The paragraph whose layout the ruler shows, as Word shows it whether or not the document has
   *  the focus: the caret's; else, after a re-render (a view switch) or with the focus in the
   *  ribbon, the one it showed last, found again by its anchor; else the document's first. */
  private rulerBlock(): HTMLElement | null {
    const shown = (el: HTMLElement | null | undefined): el is HTMLElement =>
      !!el && el.isConnected && !el.closest("#pagination-staging, .page-staging");
    if (shown(this.activeBlock)) return this.activeBlock;
    if (this.rulerAnchor) {
      const again = Array.from(this.container.querySelectorAll<HTMLElement>(
        `[data-anchor="${CSS.escape(this.rulerAnchor)}"][contenteditable="true"]`)).find(shown);
      if (again) return again;
    }
    return Array.from(this.editRoot.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))
      .find((el) => shown(el) && !this.isBandBlock(el)) ?? null;
  }

  /** Before a ruler edit: when the caret is not in the document (the focus is in the ribbon, or a
   *  view switch rendered it afresh), put it in the paragraph the ruler shows, so the edit lands
   *  there. A selection in the document is kept. */
  focusRulerParagraph(): void {
    const block = this.rulerBlock();
    if (!block || block === this.activeBlock) return;
    block.focus({ preventScroll: true });
    if (typeof window !== "undefined") {
      const sel = window.getSelection();
      if (sel && !block.contains(sel.anchorNode)) sel.collapse(block, 0);
    }
  }

  /**
   * Make `stops` the tab stops of the active paragraph, or of every selected one, as ONE undo
   * step (Word's Tabs dialog and ruler). Returns false when there is no paragraph or the engine
   * has no tab-stop op.
   */
  setTabStops(stops: TabStop[]): boolean {
    const block = this.activeBlock;
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !block || typeof bridge.SetTabStops !== "function") return false;
    const json = JSON.stringify(stops);
    const blocks = this.selectedBlocks();
    const tx = blocks.length > 1 && typeof bridge.BeginTransaction === "function" ? bridge.BeginTransaction(this.handle) : 0;
    try {
      if (blocks.length > 1 && this.applyParagraphOpAcrossBlocks(blocks, (id) => bridge.SetTabStops!(this.handle, id, json)))
        return true;
      const unid = block.getAttribute("data-anchor");
      let fullId = unid ? this.anchorIdOf(block) : undefined;
      if (!unid || !fullId) return false;
      fullId = this.syncBlock(block, fullId);
      if (!fullId) return false;
      const res = this.parseEdit(bridge.SetTabStops(this.handle, fullId, json));
      if (!res.success) return false;
      this.swapBlock(block, unid, res.modified?.[0])?.focus();
      return true;
    } finally {
      if (tx) bridge.CommitTransaction!(tx);
    }
  }

  /** The document's default tab interval (twips): the automatic stops after the last set one.
   *  One undo step; the whole document is laid out again. */
  setDefaultTabStop(twips: number): boolean {
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || typeof bridge.SetDefaultTabStop !== "function") return false;
    const res = this.parseEdit(bridge.SetDefaultTabStop(this.handle, twips));
    if (!res.success) return false;
    this.remount(this.activeBlock ? this.blockIndex(this.activeBlock) : -1, false);
    return true;
  }

  /** Word's "Suppress line numbers" for the active paragraph (or every selected one). */
  setSuppressLineNumbers(on = true): void {
    this.applyParagraphFormat({ suppressLineNumbers: on });
  }

  /** Set the paragraph style of the active block — or of every block in a multi-block selection
   *  (e.g. "Heading1", "Heading2", "Normal"). */
  setParagraphStyle(styleId: string): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (
      blocks.length > 1 &&
      this.applyParagraphOpAcrossBlocks(blocks, (id) =>
        this.exports.DocxSessionBridge.SetParagraphStyle(this.handle, id, styleId),
      )
    )
      return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const res = this.parseEdit(this.exports.DocxSessionBridge.SetParagraphStyle(this.handle, fullId, styleId));
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, idx, false, true); return; }
    this.swapBlock(block, unid, res.modified?.[0])?.focus();
  }

  /** Undo the last edit (incremental repaint; falls back to a full re-render). Typing still in the
   *  focused block is committed first, as a step of its own, so it is what Undo takes back. */
  undo(): void {
    this.deselectPicture();
    this.stepHistory((handle) => this.exports.DocxSessionBridge.Undo(handle));
  }

  /** Redo the last undone edit (incremental repaint; falls back to a full re-render). */
  redo(): void {
    this.deselectPicture();
    this.stepHistory((handle) => this.exports.DocxSessionBridge.Redo(handle));
  }

  /**
   * Undo or Redo as Word does them from the keyboard. A block commits its typing on blur, so
   * without this a Ctrl+Z pressed while typing undid an OLDER step, and the typing then committed
   * on blur as a new edit, which emptied Redo. (A ribbon button worked: the click blurs the block
   * first.) The caret goes back where it was, so the next Ctrl+Z, Ctrl+Shift+Z or Ctrl+Y still
   * reaches the editor after the repaint has replaced the focused block.
   */
  private stepHistory(step: (handle: number) => boolean): void {
    if (this.closed) return;
    const el = this.activeBlock;
    const focused = !!el && el.isConnected && el.contains(el.ownerDocument.activeElement);
    let caret: { index: number; offset: number } | null = null;
    let story: { which: BandWhich; offset: number } | null = null;
    if (this.typingPause !== null) clearTimeout(this.typingPause); // (committed below, if anything)
    this.typingPause = null;
    if (focused && el) {
      if (this.commitTypingStep(el) === false) return; // refused: the block shows the session again
      const offset = caretOffsetIn(el) ?? 0;
      const band = this.region?.bandOf(el);
      if (band) story = { which: this.region!.whichOf(band), offset };
      else {
        const index = this.editableList().indexOf(el);
        if (index >= 0) caret = { index, offset };
      }
    }
    const sectionsBefore = this.sectionSignature();
    if (!step(this.handle)) return;
    this.invalidateBlockMoveTargets();
    this.reconcile();
    // A section's columns, start and page live in its w:sectPr, which the paragraph diff cannot
    // see either: when an undone or redone step changed them, the whole flow renders again.
    if (this.sectionSignature() !== sectionsBefore) this.remount();
    // A list's levels live in the numbering part, not in any paragraph, so the paragraph diff
    // cannot see an undone or redone level change: when the numbers shown no longer match the
    // document's, the whole flow renders again.
    if (this.listNumbersStale()) this.remount();
    // The header and footer stories sit outside the body's diff, so repaint them from the session
    // too (two small renders); otherwise an undone header edit stays on screen.
    this.region?.refreshAll();
    let target: HTMLElement | undefined;
    if (story) {
      if (!this.region?.focusStory(story.which, null)) return;
      const active = el?.ownerDocument.activeElement;
      target = active instanceof HTMLElement && active.isContentEditable ? active : undefined;
    } else if (caret) {
      const blocks = this.editableList();
      target = blocks[Math.min(caret.index, blocks.length - 1)];
    }
    if (!target) return;
    this.activeBlock = target;
    const offset = (story ?? caret)!.offset;
    placeCaretAtOffset(target, Math.min(offset, blockContentText(target).length));
  }

  /**
   * Commit the typing in `el` as one step, keeping the DOM (it already shows the typing). Null
   * when nothing was typed (no step, so Redo keeps its steps); false when the engine refused it.
   */
  /**
   * `typingUndo: "word"`: commit the typing as a word ends (a space typed after a non-space), and
   * after a pause of a second, so each becomes its own undo step. Nothing happens while an IME is
   * composing, nor for anything but plain typing and deleting (paste, drops and line breaks have
   * their own paths, and Enter splits the block, which commits it).
   */
  private onTypingInput(el: HTMLElement, ev: InputEvent): void {
    if (this.options.typingUndo !== "word" || this.closed) return;
    if (this.typingPause !== null) clearTimeout(this.typingPause);
    this.typingPause = null;
    if (ev.isComposing || !/^(insertText|deleteContent)/.test(ev.inputType)) return;
    const commitIfStillHere = () => {
      this.typingPause = null;
      if (this.closed || this.replacing || !el.isConnected) return;
      if (el.ownerDocument.activeElement !== el) return; // the blur already committed it
      this.commitTypingStep(el);
    };
    const text = blockContentText(el);
    const caret = caretOffsetIn(el);
    const endsWord = ev.inputType === "insertText" && typeof ev.data === "string" && /^[  ]$/.test(ev.data)
      && caret !== null && caret >= 2 && !/\s/.test(text[caret - 2] ?? " ");
    if (endsWord) commitIfStillHere();
    else this.typingPause = setTimeout(commitIfStillHere, 1000);
  }

  /**
   * Show the session's text in `el` (its contents re-rendered in place: the element, its wiring
   * and the focus stay), the caret at `caret` when the block has the focus. See commitTypingStep.
   */
  private showCommittedText(el: HTMLElement, anchorId: string, caret: number | null): void {
    const fresh = this.renderInto(anchorId);
    if (!fresh || fresh.tagName !== el.tagName) return;
    el.replaceChildren(...Array.from(fresh.childNodes));
    if (caret !== null && el.ownerDocument.activeElement === el)
      placeCaretAtOffset(el, Math.min(caret, blockContentText(el).length));
  }

  private commitTypingStep(el: HTMLElement): boolean | null {
    const fullId = this.anchorIdOf(el);
    if (!fullId) return null;
    // The first typing in a blank paragraph goes in trimmed (commitTextChange), while the DOM
    // still holds the placeholder space the paragraph rendered as, before or after the typing.
    // The DOM then no longer matches the session, offset by offset, and the next commit's span
    // lands past the session's text: the engine refuses it, and the block shows the session
    // again, taking back what was typed since (and the focus). So such a commit shows the
    // session's text in the block, the caret mapped into it as the commit trimmed it.
    const wasBlank = trimSpaces(stripBidi(el.dataset.committedText ?? "")).length === 0;
    const domCaret = wasBlank ? caretOffsetIn(el) : null;
    const caretAfter = domCaret === null ? null : trimmedSplitOffset(el, domCaret);
    const result = this.commitTextChange(el, fullId);
    if (!result) return null;
    if (!result.success) {
      this.restoreBlock(el, fullId);
      return false;
    }
    if (wasBlank) this.showCommittedText(el, result.modified?.[0]?.id ?? fullId, caretAfter);
    el.dataset.committedText = blockContentText(el);
    // The block keeps its unid, and its render signature still describes the text BEFORE the
    // typing. Undo gives the session that text back, so the repaint's diff would match the
    // signature and leave the typing on screen. Without one the block reads as a single-block swap
    // (which never carries a signature) and is repainted. Should the commit move the block to
    // another unid, follow it, as the blur commit does. (A band block resolves through
    // data-hf-anchor and is repainted by its own region.)
    el.removeAttribute("data-render-sig");
    const modified = result.modified?.[0];
    const oldUnid = el.getAttribute("data-anchor");
    if (modified?.unid && modified.id && oldUnid && modified.unid !== oldUnid && !this.isBandBlock(el)) {
      el.setAttribute("data-anchor", modified.unid);
      this.unidToFullId.delete(oldUnid);
      this.unidToFullId.set(modified.unid, modified.id);
    }
    this.options.onEdit?.({ anchorId: modified?.id ?? fullId, unid: modified?.unid ?? oldUnid ?? "" });
    return true;
  }

  // ─── Header/footer region commands (no-ops unless `headerFooter` is on) ───────────────

  /**
   * Select which story kind a band edits (`"default"` / `"first"` / `"even"`). A kind with no
   * existing part is created empty, so the band always presents something editable. `"first"`
   * sets the section's `w:titlePg`; `"even"` sets the document-global `w:evenAndOddHeaders`
   * (which also governs footers — the band surfaces that caveat inline). Under tracked changes
   * the engine refuses `"even"` (that setting has no revision form): the refusal reaches
   * `onEditFailed` and the band keeps the kind it showed.
   */
  setHeaderFooterKind(which: BandWhich, kind: HeaderFooterKind): void {
    this.assertOpen();
    this.region?.setKind(which, kind);
  }

  /** The story kind a band is currently editing, or null when the region is off. */
  headerFooterKind(which: BandWhich): HeaderFooterKind | null {
    return this.region?.kindOf(which) ?? null;
  }

  /**
   * Append a page-number field to the focused header/footer story paragraph (falling back to the
   * band's last paragraph — Word's convention). No-op outside a band.
   */
  insertPageNumber(field: "currentPage" | "totalPages" | "pageOfTotal" = "currentPage"): void {
    this.assertOpen();
    if (!this.region) return;
    const block = this.activeBlock;
    const band = block ? this.region.bandOf(block) : null;
    const anchorId = block?.getAttribute("data-hf-anchor");
    if (block && band && anchorId) {
      // Flush uncommitted typing first: the field appends to the SESSION paragraph, and the
      // story repaint that follows would otherwise blur-commit stale DOM text over it.
      const synced = this.syncBlock(block, anchorId);
      if (!synced) return;
      this.region.insertPageNumber(this.region.whichOf(band), synced, field);
      return;
    }
    // No band block focused — target the footer, where page numbers overwhelmingly live.
    this.region.insertPageNumberInBand("footer", field);
  }

  /**
   * Set the page numbering of the section the bands describe (`w:pgNumType`) — Word's *Format Page
   * Numbers…*: `start` restarts numbering at that number, `format` chooses `1, 2, 3` vs
   * `i, ii, iii` etc. Omitted fields are left unchanged. Requires the header/footer region
   * (`{ headerFooter: true }`); a no-op otherwise.
   *
   * Inserted page-number fields are plain, so they render through this. The editor's own view still
   * shows each field's cached result — Word recomputes on open — but `{ paginated: true }`
   * substitutes the real per-page number and so reflects the change immediately.
   */
  setPageNumbering(op: { start?: number; format?: NumberFormat }): void {
    this.assertOpen();
    this.region?.setPageNumbering(op);
  }

  /** Remove the section's page-numbering start/format: it reverts to continuing the previous
   *  section's numbering in Word's default `1, 2, 3`. */
  clearPageNumbering(): void {
    this.assertOpen();
    this.region?.clearPageNumbering();
  }

  /** This section's page numbering as the document currently states it — `{}` when the section
   *  sets neither (it continues the previous section in the default format). */
  pageNumbering(): { start?: number; format?: NumberFormat } {
    return this.region?.pageNumbering() ?? {};
  }

  // ─── Font group ──────────────────────────────────────────────────────

  /**
   * Apply one inline `FormatOp` to the selection: a sub-range of the active block, the whole
   * block when the caret is collapsed, or every block of a multi-block selection. The last real
   * selection is used when a focus-stealing control (a colour picker) collapsed the live one.
   */
  private applyInlineFormat(op: Record<string, unknown>): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (blocks.length > 1 && this.applyInlineOpAcrossBlocks(blocks, op)) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    let span = selectionSpanIn(block);
    if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return;
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyFormat(
        this.handle, fullId, span ? JSON.stringify(span) : "", JSON.stringify(op),
      ),
    );
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, this.blockIndex(block), false); return; }
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh && span) selectRange(fresh, span.start, span.length);
    else fresh?.focus();
  }

  /** Font colour as a hex triplet (with or without '#'); `""` clears the explicit colour. */
  setFontColor(hex: string): void {
    this.applyInlineFormat({ color: hex.replace(/^#/, "").toUpperCase() });
  }

  /** Word highlight colour name (`"yellow"`, `"green"`, …); `""` removes the highlight. */
  setHighlight(name: string): void {
    this.applyInlineFormat({ highlight: name });
  }

  setAllCaps(on: boolean): void {
    this.applyInlineFormat({ caps: on });
  }

  setSmallCaps(on: boolean): void {
    this.applyInlineFormat({ smallCaps: on });
  }

  /** Word's "Clear All Formatting": drop every direct run property on the selection. */
  clearFormatting(): void {
    this.applyInlineFormat({
      bold: false, italic: false, underline: false, strike: false, code: false, color: "",
      highlight: "", vertAlign: "", fontSizePts: 0, fontFamily: "", runStyle: "",
      caps: false, smallCaps: false,
    });
  }

  // ─── Format Painter ──────────────────────────────────────────────────

  /** What Copy Formatting took, for Paste Formatting and the painter. */
  private copiedFormat: CopiedFormatting | null = null;
  /** Format Painter's state: off, one paint, or until turned off. */
  private painter: FormatPainterState = "off";

  /**
   * Word's Copy Formatting (Ctrl+Shift+C): remember the formatting at the caret or of the
   * selection. With a caret it takes the paragraph's formatting (style, alignment, indents,
   * spacing, list) and the character formatting there; with a selection, the character
   * formatting most of it has, and the paragraph's only when the selection takes in the
   * paragraph's end. False when there is nothing to copy from.
   */
  copyFormatting(): boolean {
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !this.activeBlock || typeof bridge.GetFormatting !== "function") return false;
    const blocks = this.selectedBlocks();
    const block = blocks[0] ?? this.activeBlock;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    const unid = block.getAttribute("data-anchor");
    const cross = blocks.length > 1;
    let span = cross ? this.blockSpanForSelection(block) : selectionSpanIn(block);
    if (!span && !cross && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
    const caret = span ? null : caretOffsetIn(block) ?? 0;
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return false;
    let inspection: FormattingInspection | null;
    try {
      inspection = JSON.parse(bridge.GetFormatting(this.handle, fullId)) as FormattingInspection | null;
    } catch {
      inspection = null;
    }
    if (!inspection) return false;
    const run = runOpFromDirect(pickRunFormat(inspection.runs ?? [], span, caret));
    const length = trimBounds(block).trimmedLen;
    const withParagraph = !span || cross || span.start + span.length >= length;
    let paragraph: CopiedFormatting["paragraph"] = null;
    if (withParagraph && inspection.effectiveParagraph) {
      let list: { numId: number; level: number } | null = null;
      try {
        const m = JSON.parse(bridge.GetListMembership(this.handle, fullId)) as { numId?: number; level?: number } | null;
        if (m && typeof m.numId === "number" && m.numId > 0) list = { numId: m.numId, level: m.level ?? 0 };
      } catch {
        list = null;
      }
      paragraph = {
        styleId: inspection.effectiveParagraph.styleId ?? inspection.directParagraph?.styleId ?? null,
        format: asEditorParagraphFormat(inspection.effectiveParagraph),
        list,
      };
    }
    this.copiedFormat = { run, paragraph };
    return true;
  }

  /**
   * Word's Paste Formatting (Ctrl+Shift+V): give the selection what {@link copyFormatting} took —
   * with a caret, the word around it — and, when the copy took the paragraph's formatting, every
   * paragraph the selection touches. One undo step. False when nothing was copied or there is
   * nowhere to paint.
   */
  pasteFormatting(): boolean {
    const copied = this.copiedFormat;
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !copied || !this.activeBlock) return false;
    const blocks = this.selectedBlocks();
    const first = blocks[0] ?? this.activeBlock;
    // What to paint in each block, read before any edit: a span, the whole block (null), or no
    // characters at all (a caret between words: the paragraph only).
    let caret: number | null = null;
    let restore: { start: number; length: number } | null = null;
    let targets: Array<{ block: HTMLElement; fullId: string | undefined; span: { start: number; length: number } | null; chars: boolean }>;
    if (blocks.length > 1) {
      targets = this.multiBlockTargets(blocks).map((t) => ({ block: t.block, fullId: t.fullId, span: t.span, chars: true }));
    } else {
      const block = this.activeBlock;
      const unid = block.getAttribute("data-anchor");
      let span = selectionSpanIn(block);
      if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
      restore = span;
      let chars = true;
      if (!span) {
        caret = caretOffsetIn(block) ?? 0;
        const word = wordSpanAt(blockContentText(block), caret);
        span = word ? trimmedSpan(block, word) : null;
        chars = !!span && span.length > 0;
      }
      targets = [{ block, fullId: this.anchorIdOf(block), span, chars }];
    }
    if (!copied.paragraph && !targets.some((t) => t.chars)) return false;
    const index = this.blockIndex(first);
    const tx = this.beginListTransaction();
    let edited: { anchorId: string; unid: string } | null = null;
    // Like every other format command, a paint repaints only the blocks it changed (a swap
    // each); the whole document renders again only when the paint changed what a swap cannot
    // show: list numbering, a border wrapper, or — in page view — where pages break.
    let wholeFlow = false;
    const done: Array<{ block: HTMLElement; unid: string | null; span: { start: number; length: number } | null; res: EditResultLite | null }> = [];
    try {
      for (const t of targets) {
        if (!t.fullId) continue;
        let id = this.syncBlock(t.block, t.fullId);
        if (!id) continue;
        const unid = t.block.getAttribute("data-anchor");
        let last: EditResultLite | null = null;
        const step = (json: string): void => {
          const res = this.parseEdit(json);
          if (!res.success) return;
          id = res.modified?.[0]?.id ?? id;
          last = res;
          if (this.affectsList(res)) wholeFlow = true;
          edited = { anchorId: id!, unid: unid ?? "" };
        };
        const para = copied.paragraph;
        if (para) {
          if (t.block.parentElement?.tagName === "DIV" && inBorderWrapper(t.block)) wholeFlow = true;
          if (this.options.paginated && this.pageKeepingDiffers(id, para.format)) wholeFlow = true;
        }
        if (t.chars) {
          step(bridge.ApplyFormat(this.handle, id, t.span ? JSON.stringify(t.span) : "", JSON.stringify(copied.run)));
        }
        if (para) {
          if (para.styleId) step(bridge.SetParagraphStyle(this.handle, id!, para.styleId));
          if (para.list && typeof bridge.SetListMembership === "function") {
            step(bridge.SetListMembership(this.handle, id!, para.list.numId, para.list.level));
            wholeFlow = true;
          } else if (!para.list && isListBlock(t.block)) {
            step(bridge.ApplyListFormat(this.handle, id!, "none"));
            wholeFlow = true;
          }
          step(bridge.SetParagraphFormat(this.handle, id!, JSON.stringify(para.format)));
        }
        done.push({ block: t.block, unid, span: t.span, res: last });
      }
    } finally {
      this.endListTransaction(tx);
    }
    if (!edited) return false;
    if (blocks.length > 1) {
      // The multi-block finisher swaps each block and keeps the selection (or renders the flow
      // again where it must: lists, page view).
      this.finishMultiBlockOp(done, wholeFlow);
      this.options.onEdit?.(edited);
      return true;
    }
    const only = done[0];
    let block: HTMLElement | null | undefined = null;
    if (!wholeFlow && only?.res && only.unid) {
      block = this.swapBlock(only.block, only.unid, (only.res as EditResultLite).modified?.[0]);
    }
    if (!block) {
      this.refreshAfter(first, index, false, true);
      block = this.editableList()[index];
      this.options.onEdit?.(edited);
    }
    if (block) {
      this.activeBlock = block;
      block.focus();
      if (restore) selectRange(block, restore.start, restore.length);
      else placeCaretAtOffset(block, caret ?? 0);
    }
    return true;
  }

  /** Whether painting `format` on the paragraph `fullId` changes where pages break (Page break
   *  before, Keep with next, Keep lines together, Widow/Orphan control). */
  private pageKeepingDiffers(fullId: string, format: EditorParagraphFormat): boolean {
    let now: ParagraphFormatting | undefined;
    try {
      now = (JSON.parse(this.exports.DocxSessionBridge.GetFormatting!(this.handle, fullId)) as FormattingInspection | null)
        ?.effectiveParagraph;
    } catch {
      return true;
    }
    if (!now) return true;
    return !!format.pageBreakBefore !== !!now.pageBreakBefore || !!format.keepNext !== !!now.keepNext
      || !!format.keepLines !== !!now.keepLines || (format.widowControl ?? true) !== (now.widowControl ?? true);
  }

  /**
   * Turn Format Painter on: copy the formatting here ({@link copyFormatting}), then paint it on
   * the next click or selection in the document — once, or with `sticky` (Word's double-click on
   * the brush) every time until {@link stopFormatPainter} or Esc. False when there is nothing to
   * copy from.
   */
  formatPainter(sticky = false): boolean {
    if (this.closed || !this.copyFormatting()) return false;
    this.setPainter(sticky ? "sticky" : "once");
    return true;
  }

  /** Turn Format Painter off (Esc, or a second click on its button). */
  stopFormatPainter(): void {
    this.setPainter("off");
  }

  /** Whether Format Painter is on, and for how long. */
  get formatPainterState(): FormatPainterState {
    return this.painter;
  }

  private setPainter(state: FormatPainterState): void {
    if (state === this.painter) return;
    const doc = this.container.ownerDocument;
    if (state !== "off") {
      ensureFormatPainterStyles(doc);
      doc.addEventListener("keydown", this.onPainterKey, true);
    } else {
      doc.removeEventListener("keydown", this.onPainterKey, true);
    }
    this.container.classList.toggle("docx-format-painting", state !== "off");
    this.painter = state;
    this.options.onFormatPainterChange?.(state);
  }

  private readonly onPainterKey = (ev: KeyboardEvent): void => {
    if (ev.key !== "Escape" || this.painter === "off") return;
    ev.preventDefault();
    ev.stopPropagation();
    this.stopFormatPainter();
  };

  /** A click or drag ended in the document while Format Painter is on. */
  private paintAtSelection(): void {
    if (this.closed || this.painter === "off") return;
    this.pasteFormatting();
    if (this.painter === "once") this.stopFormatPainter();
  }

  // ─── Change Case ─────────────────────────────────────────────────────

  /**
   * Word's Change Case on the selection — with a caret, the word around it: Sentence case,
   * lowercase, UPPERCASE, Capitalize Each Word, tOGGLE cASE, or Headline Case (small words in
   * lower case, abbreviations as written: the `headlineCaseLists`). `"cycle"` is Word's Shift+F3:
   * capitals go to lower case, lower case to Capitalize Each Word, anything else to capitals. Only
   * the letters that change are written, each keeping its formatting. One undo step; the
   * selection stays. False when there is nothing to change.
   */
  changeCase(mode: CaseMode | "cycle"): boolean {
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !this.activeBlock || typeof bridge.GetFormatting !== "function") return false;
    const blocks = this.selectedBlocks();
    const first = blocks[0] ?? this.activeBlock;
    let caret: number | null = null;
    let restore: { start: number; length: number } | null = null;
    let targets: Array<{ block: HTMLElement; fullId: string | undefined; span: { start: number; length: number } | null }>;
    if (blocks.length > 1) {
      targets = this.multiBlockTargets(blocks).map((t) => ({ block: t.block, fullId: t.fullId, span: t.span }));
    } else {
      const block = this.activeBlock;
      const unid = block.getAttribute("data-anchor");
      let span = selectionSpanIn(block);
      if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
      restore = span;
      if (!span) {
        caret = caretOffsetIn(block) ?? 0;
        const word = wordSpanAt(blockContentText(block), caret);
        span = word ? trimmedSpan(block, word) : null;
        if (!span || span.length === 0) return false;
      }
      targets = [{ block, fullId: this.anchorIdOf(block), span }];
    }
    // Each block's text as the engine has it, run by run (characters no run covers are passed
    // over), and the selected stretch of it.
    type Piece = { block: HTMLElement; span: { start: number; length: number } | null; id: string; runs: InlineSpan[]; flat: string[]; hidden: boolean[]; from: number; to: number };
    const pieces: Piece[] = [];
    for (const t of targets) {
      if (!t.fullId) continue;
      const id = this.syncBlock(t.block, t.fullId);
      if (!id) continue;
      let runs: InlineSpan[];
      try {
        runs = (JSON.parse(bridge.GetFormatting(this.handle, id)) as FormattingInspection | null)?.runs ?? [];
      } catch {
        continue;
      }
      runs = runs.filter((r) => r.span.length > 0 && r.text.length === r.span.length);
      const end = runs.reduce((m, r) => Math.max(m, r.span.start + r.span.length), 0);
      const flat: string[] = new Array<string>(end).fill(" ");
      const hidden: boolean[] = new Array<boolean>(end).fill(true);
      for (const r of runs) {
        for (let k = 0; k < r.span.length; k++) { flat[r.span.start + k] = r.text[k]!; hidden[r.span.start + k] = false; }
      }
      const from = Math.min(t.span?.start ?? 0, end);
      const to = Math.min(t.span ? t.span.start + t.span.length : end, end);
      if (to > from) pieces.push({ block: t.block, span: t.span, id, runs, flat, hidden, from, to });
    }
    if (pieces.length === 0) return false;
    // The selection as one text (a paragraph a line), so Sentence and Headline Case see where
    // lines start and end.
    const text = pieces.map((p) => p.flat.slice(p.from, p.to).join("")).join("\n");
    const hidden = pieces.flatMap((p, i) => [...p.hidden.slice(p.from, p.to), ...(i < pieces.length - 1 ? [false] : [])]);
    const resolved: CaseMode = mode === "cycle" ? nextCycleCase(text.split("").filter((_, i) => !hidden[i]).join("")) : mode;
    const cased = changeCase(text, resolved, this.options.headlineCaseLists ?? DEFAULT_HEADLINE_CASE_LISTS, hidden);
    if (cased === text) return false;
    const index = this.blockIndex(first);
    const tx = this.beginListTransaction();
    let edited: { anchorId: string; unid: string } | null = null;
    const done: Array<{ block: HTMLElement; unid: string | null; span: { start: number; length: number } | null; res: EditResultLite | null }> = [];
    try {
      let at = 0;
      for (const p of pieces) {
        const want = cased.slice(at, at + (p.to - p.from));
        at += p.to - p.from + 1;
        const unid = p.block.getAttribute("data-anchor");
        let id = p.id;
        let last: EditResultLite | null = null;
        // The stretches to write: changed characters inside one run, joined across unchanged
        // letters, digits and spaces (never across a quote, which the engine would make curly).
        for (const r of p.runs) {
          const lo = Math.max(p.from, r.span.start), hi = Math.min(p.to, r.span.start + r.span.length);
          let k = lo;
          while (k < hi) {
            if (want[k - p.from] === p.flat[k]) { k++; continue; }
            let e = k + 1, lastChanged = k;
            while (e < hi) {
              const same = want[e - p.from] === p.flat[e];
              if (!same) lastChanged = e;
              else if (!/[\p{L}\p{N} ]/u.test(p.flat[e]!)) break;
              e++;
            }
            const len = lastChanged - k + 1;
            const res = this.parseEdit(bridge.ReplaceTextAtSpan(this.handle, id, k, len, want.slice(k - p.from, k - p.from + len)));
            if (res.success) {
              id = res.modified?.[0]?.id ?? id;
              last = res;
              edited = { anchorId: id, unid: unid ?? "" };
            }
            k = lastChanged + 1;
          }
        }
        done.push({ block: p.block, unid, span: p.span, res: last });
      }
    } finally {
      this.endListTransaction(tx);
    }
    if (!edited) return false;
    if (blocks.length > 1) {
      this.finishMultiBlockOp(done, false, /* swapInPageView */ true);
      this.options.onEdit?.(edited);
      return true;
    }
    const only = done[0];
    let block: HTMLElement | null | undefined = null;
    if (only?.res && only.unid) block = this.swapBlock(only.block, only.unid, (only.res as EditResultLite).modified?.[0]);
    if (!block) {
      this.refreshAfter(first, index, false, true);
      block = this.editableList()[index];
      this.options.onEdit?.(edited);
    }
    if (block) {
      this.activeBlock = block;
      block.focus();
      if (restore) selectRange(block, restore.start, restore.length);
      else placeCaretAtOffset(block, caret ?? 0);
    }
    return true;
  }

  /** The word lists Headline Case uses (see the `headlineCaseLists` option). */
  setHeadlineCaseLists(lists: HeadlineCaseLists | null): void {
    this.options.headlineCaseLists = lists ?? undefined;
  }

  // ─── Show/Hide ¶ and joining paragraphs ──────────────────────────────

  /** Word's Show/Hide ¶ (see the `showFormattingMarks` option). */
  setShowFormattingMarks(on: boolean): void {
    if (!!this.options.showFormattingMarks === on) return;
    this.options.showFormattingMarks = on;
    this.applyFormattingMarks();
    this.options.onFormattingMarksChange?.(on);
  }

  /** Whether Show/Hide ¶ is on. */
  get showFormattingMarks(): boolean {
    return !!this.options.showFormattingMarks;
  }

  private applyFormattingMarks(): void {
    const on = !!this.options.showFormattingMarks;
    ensureSectionBreakStyles(this.container.ownerDocument);
    this.container.classList.toggle("docx-paginated", !!this.options.paginated);
    if (on) ensureFormattingMarksStyles(this.container.ownerDocument);
    this.container.classList.toggle("docx-show-marks", on);
  }

  /**
   * Join the paragraphs the selection touches into one, keeping the last one's paragraph mark —
   * Word's "delete every ¶ but the last", the way to mend text copied from a PDF with a paragraph
   * for every printed line. Where two lines meet without a space, a space goes in, except after a
   * hyphen. The joined paragraph keeps the first one's formatting. One undo step. False when the
   * selection is in one paragraph, or takes in a table.
   */
  joinParagraphs(): boolean {
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !this.activeBlock) return false;
    const blocks = this.selectedBlocks();
    if (blocks.length < 2) return false;
    // Paragraphs side by side in the body (a table, or a cell's paragraphs, are left alone).
    if (blocks.some((b) => b.closest("table") || this.isBandBlock(b))) return false;
    const first = blocks[0]!;
    const index = this.blockIndex(first);
    const ids: string[] = [];
    for (const b of blocks) {
      const full = this.anchorIdOf(b);
      const id = full ? this.syncBlock(b, full) : undefined;
      if (!id) return false;
      ids.push(id);
    }
    const text = (id: string): string => {
      try {
        const runs = (JSON.parse(bridge.GetFormatting!(this.handle, id)) as FormattingInspection | null)?.runs ?? [];
        return runs.map((r) => r.text).join("");
      } catch {
        return "";
      }
    };
    const tracked = this.options.trackedChanges === TrackedChangeMode.RenderInline;
    const tx = this.beginListTransaction();
    let current = ids[0]!;
    let joined = 0;
    try {
      for (const next of ids.slice(1)) {
        const before = text(current);
        const after = text(next);
        const res = this.parseEdit(bridge.MergeParagraphs(this.handle, current, next));
        if (!res.success) break;
        current = res.modified?.[0]?.id ?? current;
        joined++;
        // The engine puts a space between two lines that have none; after a hyphen there is none.
        if (!tracked && /-$/.test(before) && before.length > 0 && after.length > 0 && !/^\s/.test(after)) {
          const fix = this.parseEdit(bridge.ReplaceTextAtSpan(this.handle, current, before.length, 1, ""));
          if (fix.success) current = fix.modified?.[0]?.id ?? current;
        }
      }
    } finally {
      this.endListTransaction(tx);
    }
    if (joined === 0) return false;
    this.invalidateBlockMoveTargets();
    this.reconcile();
    const block = this.editableList()[index];
    if (block) {
      this.activeBlock = block;
      block.focus();
      placeCaretAtOffset(block, blockContentText(block).length);
    }
    this.options.onEdit?.({ anchorId: current, unid: block?.getAttribute("data-anchor") ?? "" });
    return true;
  }

  /** The caret's rendered font size in points (from computed style), or null. */
  fontSizeAtCaret(): number | null {
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    let el: HTMLElement | null = this.activeBlock;
    if (sel && sel.rangeCount > 0) {
      const n = sel.getRangeAt(0).startContainer;
      const candidate = n.nodeType === 3 ? n.parentElement : (n as HTMLElement);
      if (candidate && this.container.contains(candidate)) el = candidate;
    }
    if (!el || typeof getComputedStyle !== "function") return null;
    const px = parseFloat(getComputedStyle(el).fontSize);
    return px > 0 ? Math.round(px * 0.75 * 2) / 2 : null;
  }

  /** Grow (+) or shrink (−) the selection's font size by `delta` points (Word's A↑ / A↓). */
  adjustFontSize(delta: number): void {
    const size = this.fontSizeAtCaret();
    if (size == null) return;
    this.setFontSize(Math.max(1, Math.round((size + delta) * 2) / 2));
  }

  // ─── Paragraph group ─────────────────────────────────────────────────

  /** Line spacing as a multiple of single (1, 1.15, 1.5, 2 …) — `w:spacing/@w:line` under `auto`. */
  setLineSpacing(multiple: number): void {
    this.applyParagraphFormat({ lineSpacing: Math.round(multiple * 240), lineSpacingRule: "auto" });
  }

  /** Space before/after the paragraph, in POINTS (Word's Paragraph dialog units). */
  setParagraphSpacing(op: { beforePt?: number; afterPt?: number }): void {
    const payload: { spacingBefore?: number; spacingAfter?: number } = {};
    if (op.beforePt != null) payload.spacingBefore = Math.max(0, Math.round(op.beforePt * 20));
    if (op.afterPt != null) payload.spacingAfter = Math.max(0, Math.round(op.afterPt * 20));
    if (Object.keys(payload).length === 0) return;
    this.applyParagraphFormat(payload);
  }

  /** First-line indent in twips (0 = none); removes any hanging indent. */
  setFirstLineIndent(twips: number): void {
    this.applyParagraphFormat({ firstLineIndent: Math.max(0, twips) });
  }

  /** Hanging indent in twips (0 = none); removes any first-line indent. */
  setHangingIndent(twips: number): void {
    this.applyParagraphFormat({ hangingIndent: Math.max(0, twips) });
  }

  /**
   * Make the active block (or every selected block) a list item of `kind` — the full Word
   * numbering gallery, not just bullets/decimal — or a plain paragraph with `"none"`.
   */
  setListFormat(kind: ListFormat): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    const wasList = isListBlock(block);
    const tx = this.beginListTransaction();
    try {
      let firstId: string | undefined;
      if (
        blocks.length > 1 &&
        this.applyParagraphOpAcrossBlocks(
          blocks,
          (id) => {
            const json = this.exports.DocxSessionBridge.ApplyListFormat(this.handle, id, kind);
            try { firstId ??= (JSON.parse(json) as EditResultLite).modified?.[0]?.id ?? id; } catch { /* the op's own error */ }
            return json;
          },
          true,
        )
      ) {
        if (!wasList && firstId && kind !== "none") this.applyListDefaults(firstId, kind, true);
        return;
      }
      const unid = block.getAttribute("data-anchor");
      if (!unid) return;
      let fullId = this.anchorIdOf(block);
      if (!fullId) return;
      const idx = this.blockIndex(block);
      fullId = this.syncBlock(block, fullId);
      if (!fullId) return;
      const res = this.parseEdit(this.exports.DocxSessionBridge.ApplyListFormat(this.handle, fullId, kind));
      if (!res.success) return;
      if (!wasList && kind !== "none") this.applyListDefaults(res.modified?.[0]?.id ?? fullId, kind);
      this.refreshAfter(block, idx, false, /* forceRemount */ true);
    } finally {
      this.endListTransaction(tx);
    }
  }

  // ─── List levels (Word's Define New Multilevel List), restart and continue ──────────────

  /** True when a list item's number on the page differs from the number the document gives it
   *  (bullets, whose glyph the engine does not report, are not compared). */
  /** The body's sections as the engine has them now (columns, start, page, line numbers), one entry per
   *  section the view shows, for telling whether a step changed any. */
  private sectionSignature(): string {
    if (this.closed) return "";
    const bridge = this.exports.DocxSessionBridge;
    const seen = new Set<string>();
    const out: string[] = [];
    for (const block of this.editableList()) {
      if (this.isBandBlock(block)) continue;
      const key = this.sectionKeyOf(block);
      if (seen.has(key)) continue;
      seen.add(key);
      const fullId = this.anchorIdOf(block);
      if (!fullId) continue;
      try {
        const i = JSON.parse(bridge.GetSectionInfo(this.handle, fullId)) as SectionInfo | null;
        if (i) out.push(JSON.stringify([i.columns, i.columnSpaceTwips, i.columnSeparator, i.sectionStart,
          i.pageWidthTwips, i.pageHeightTwips, i.marginTopTwips, i.marginBottomTwips, i.marginLeftTwips, i.marginRightTwips,
          i.lineNumbering ?? null]));
      } catch {
        // (an engine without section info: nothing to compare)
      }
    }
    return out.join(";");
  }

  private listNumbersStale(): boolean {
    const bridge = this.exports.DocxSessionBridge;
    for (const block of this.editableList()) {
      if (!isListBlock(block)) continue;
      const fullId = this.anchorIdOf(block);
      if (!fullId) continue;
      let label: string | null | undefined;
      try {
        label = (JSON.parse(bridge.GetListMembership(this.handle, fullId)) as { generatedLabel?: string } | null)?.generatedLabel;
      } catch {
        continue;
      }
      if (!label) continue;
      const shown = Array.from(block.querySelectorAll(":scope > [data-list-marker]"))
        .map((m) => m.textContent ?? "").join("").replace(/\s+/g, " ").trim();
      if (shown !== label.trim()) return true;
    }
    return false;
  }

  /** Several list ops as one undo step (0 when the engine has no transactions). */
  private beginListTransaction(): number {
    const bridge = this.exports.DocxSessionBridge;
    return typeof bridge.BeginTransaction === "function" ? bridge.BeginTransaction(this.handle) : 0;
  }

  private endListTransaction(tx: number): void {
    if (tx) this.exports.DocxSessionBridge.CommitTransaction!(tx);
  }

  /** Write the `listDefaults` into the list `anchorId` was just made an item of: whole for a
   *  numbered list, positions only for any other format. `remount` re-renders afterwards (the
   *  multi-block path has rendered already). When the nearest list before it already has those
   *  levels (and the same format), the new list continues that one instead, as a new list of the
   *  engine's own levels continues the one before it. */
  private applyListDefaults(anchorId: string, kind: string, remount = false): void {
    const defaults = this.options.listDefaults;
    const bridge = this.exports.DocxSessionBridge;
    if (!defaults?.length || typeof bridge.SetListLevels !== "function" || typeof bridge.GetListLevels !== "function") return;
    const levels: ListLevelSpec[] = kind === "decimal" ? defaults
      : defaults.map((l) => ({ level: l.level, numberPositionTwips: l.numberPositionTwips, textPositionTwips: l.textPositionTwips }));
    const read = (id: string) => {
      try { return JSON.parse(bridge.GetListLevels!(this.handle, id)) as ListLevelInfo[] | null; } catch { return null; }
    };
    const mine = read(anchorId);
    const all = this.editableList();
    const here = this.activeBlock ? all.indexOf(this.activeBlock) : -1;
    const before = here > 0 ? all.slice(0, here).reverse().find((b) => isListBlock(b) && !b.closest("table")) : undefined;
    const beforeId = before ? this.anchorIdOf(before) : undefined;
    const theirs = beforeId ? read(beforeId) : null;
    const same = !!mine && !!theirs && theirs[0]?.format === mine[0]?.format
      && levels.every((spec) => {
        const l = theirs.find((x) => x.level === spec.level);
        return !!l && (Object.keys(spec) as Array<keyof ListLevelSpec>).every((k) => spec[k] === undefined || spec[k] === l[k as keyof ListLevelInfo]);
      });
    // The levels first: ContinueNumbering joins only a list with the same number text (Word's rule).
    const set = this.parseEdit(bridge.SetListLevels(this.handle, anchorId, JSON.stringify(levels)));
    if (same && typeof bridge.ContinueNumbering === "function") {
      this.parseEdit(bridge.ContinueNumbering(this.handle, set.modified?.[0]?.id ?? anchorId));
    }
    if (remount) this.remount(here);
  }

  /** The levels new lists start from (see the `listDefaults` option); [] for the engine's own. */
  setListDefaults(levels: ListLevelSpec[]): void {
    this.options.listDefaults = levels.map((l) => ({ ...l }));
  }

  /** The levels of the caret's list — format, number text (`%1.%2`), start, positions — or null
   *  when the caret is not in a list item. */
  listLevels(): ListLevelInfo[] | null {
    const block = this.activeBlock;
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !block || typeof bridge.GetListLevels !== "function") return null;
    const fullId = this.anchorIdOf(block);
    if (!fullId) return null;
    try {
      return JSON.parse(bridge.GetListLevels(this.handle, fullId)) as ListLevelInfo[] | null;
    } catch {
      return null;
    }
  }

  /** Run a list op on the caret's paragraph (committing its typing first) and re-render. */
  private listOp(run: (fullId: string) => string): boolean {
    const block = this.activeBlock;
    if (this.closed || !block) return false;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return false;
    const res = this.parseEdit(run(fullId));
    if (!res.success) return false;
    this.refreshAfter(block, idx, false, /* forceRemount */ true);
    return true;
  }

  /**
   * Change the caret's list's levels (Word's Define New Multilevel List: format, number text,
   * start, positions, legal flag), as one undo step. Paragraphs not yet in a list (the selection's,
   * or the caret's) are made a numbered list first.
   */
  setListLevels(levels: ListLevelSpec[]): boolean {
    const bridge = this.exports.DocxSessionBridge;
    const block = this.activeBlock;
    if (this.closed || !block || typeof bridge.SetListLevels !== "function") return false;
    if (isListBlock(block)) return this.listOp((id) => bridge.SetListLevels!(this.handle, id, JSON.stringify(levels)));
    // Not a list yet: a numbered list first, then its levels, as one step.
    const tx = this.beginListTransaction();
    let ok = false;
    try {
      // (the paragraph is found again by its place: making the list renders the document afresh)
      const index = this.editableList().indexOf(this.selectedBlocks()[0] ?? block);
      const saved = this.options.listDefaults;
      this.options.listDefaults = undefined;    // (the levels given are the ones wanted)
      try { this.setListFormat("decimal"); } finally { this.options.listDefaults = saved; }
      const now = this.editableList()[index];
      if (now && isListBlock(now)) {
        this.activeBlock = now;
        ok = this.listOp((fullId) => bridge.SetListLevels!(this.handle, fullId, JSON.stringify(levels)));
      }
    } finally {
      this.endListTransaction(tx);
    }
    return ok;
  }

  /**
   * Legal numbering — 1., 1.1, 1.1.1 by level — on the caret's list (the selection, or the
   * caret's paragraph, made a numbered list first when it is not one). `periods`: `"top"` ends
   * the top level's number with a period (1. then 1.1), `"all"` every level's (1.1.), `"none"`
   * none (1, 1.1).
   */
  setLegalNumbering(periods: "top" | "all" | "none" = "top"): boolean {
    return this.setListLevels(legalLevels(periods));
  }

  /** Restart the caret's list at `start` (Word's Restart at 1 / Set Numbering Value). */
  restartNumbering(start = 1): boolean {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetListStartOverride !== "function" || !this.activeBlock || !isListBlock(this.activeBlock)) return false;
    return this.listOp((id) => bridge.SetListStartOverride!(this.handle, id, Math.max(0, Math.round(start))));
  }

  /** Word's Continue Numbering: the caret's item, and the rest of its list, join the list before. */
  continueNumbering(): boolean {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.ContinueNumbering !== "function" || !this.activeBlock || !isListBlock(this.activeBlock)) return false;
    return this.listOp((id) => bridge.ContinueNumbering!(this.handle, id));
  }

  /** The active block's list format (`"bullet"`, `"decimal"`, …), or null when it is not a list item. */
  listFormatAtCaret(): string | null {
    const block = this.activeBlock;
    if (this.closed || !block || !isListBlock(block)) return null;
    const fullId = this.anchorIdOf(block);
    if (!fullId) return null;
    try {
      const membership = JSON.parse(this.exports.DocxSessionBridge.GetListMembership(this.handle, fullId)) as { format?: string } | null;
      return membership?.format ?? null;
    } catch {
      return null;
    }
  }

  /** Paragraph formatting (direct + effective) of the active block, for ribbon state. */
  paragraphFormatting(): FormattingInspection | null {
    const block = this.activeBlock;
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !block || typeof bridge.GetFormatting !== "function") return null;
    const fullId = this.anchorIdOf(block);
    if (!fullId) return null;
    try {
      const parsed = JSON.parse(bridge.GetFormatting(this.handle, fullId)) as FormattingInspection | null;
      return parsed && typeof parsed === "object" && "effectiveParagraph" in parsed ? parsed : null;
    } catch {
      return null;
    }
  }

  // ─── Links, images, references ───────────────────────────────────────

  /** Wrap the selection in a hyperlink (external URL, or an internal bookmark name). */
  insertHyperlink(target: string, kind: "external" | "internal" = "external"): boolean {
    const block = this.activeBlock;
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !block || typeof bridge.AddHyperlink !== "function") return false;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return false;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    let span = selectionSpanIn(block);
    if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
    if (!span) return false; // Word needs a range to link
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return false;
    const res = this.parseEdit(
      bridge.AddHyperlink(this.handle, fullId, span.start, span.length, kind, target),
    );
    if (!res.success) return false;
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh) selectRange(fresh, span.start, span.length);
    return true;
  }

  /** The hyperlink the caret (or selection start) sits in, or null. */
  hyperlinkAtCaret(): HyperlinkInfo | null {
    const block = this.activeBlock;
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !block || typeof bridge.ListHyperlinks !== "function") return null;
    const fullId = this.anchorIdOf(block);
    if (!fullId) return null;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    let offset = 0;
    if (sel && sel.rangeCount > 0 && block.contains(sel.getRangeAt(0).startContainer)) {
      const r = sel.getRangeAt(0);
      offset = contentOffsetOf(block, r.startContainer, r.startOffset);
    }
    try {
      const links = JSON.parse(bridge.ListHyperlinks(this.handle, 63)) as HyperlinkInfo[];
      return (
        links.find(
          (l) => l.anchorId === fullId && offset >= l.span.start && offset <= l.span.start + l.span.length,
        ) ?? null
      );
    } catch {
      return null;
    }
  }

  /** Remove the hyperlink at the caret, keeping its text. */
  removeHyperlink(): boolean {
    const link = this.hyperlinkAtCaret();
    const bridge = this.exports.DocxSessionBridge;
    const block = this.activeBlock;
    if (!link || !block || typeof bridge.RemoveHyperlink !== "function") return false;
    const unid = block.getAttribute("data-anchor");
    const res = this.parseEdit(bridge.RemoveHyperlink(this.handle, link.id));
    if (!res.success) return false;
    if (unid) this.swapBlock(block, unid, res.modified?.[0]);
    return true;
  }

  /** Insert an inline image (base64 bytes) at the caret in the active body block. */
  insertImage(imageBase64: string, options: ImageInsertOptions = {}): boolean {
    const block = this.activeBlock;
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !block || typeof bridge.InsertImage !== "function") return false;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    const idx = this.blockIndex(block);
    const raw = caretOffsetIn(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return false;
    const offset = trimmedSplitOffset(block, raw ?? (block.textContent ?? "").length);
    const res = this.parseEdit(
      bridge.InsertImage(this.handle, fullId, offset, imageBase64, JSON.stringify(options)),
    );
    if (!res.success) return false;
    // Image parts are not in the block-render shell, so a single-block swap would drop the
    // picture; the full render is the only path that shows it.
    this.refreshAfter(block, idx, true, /* forceRemount */ true);
    return true;
  }

  /**
   * Word's Ctrl+Enter / Ctrl+Shift+Enter: a page or column break at the caret. The break ends
   * the text before the caret; the rest goes on as a paragraph of its own, where the caret lands.
   * False (nothing changed) in a table, a text box, a header or footer, or while changes are
   * tracked.
   */
  insertBreak(kind: "page" | "column" = "page"): boolean {
    return this.splitWithBreak((bridge, id, offset) => bridge.InsertBreak?.(this.handle, id, offset, kind), false);
  }

  /**
   * Word's Layout > Breaks > Section Breaks: end the section at the caret. The new section keeps
   * the page setup, headers and footers of the one it came from and starts as `start` says.
   * False (nothing changed) where {@link insertBreak} is refused.
   */
  insertSectionBreak(start: SectionStart = "nextPage"): boolean {
    return this.splitWithBreak((bridge, id, offset) => bridge.InsertSectionBreak?.(this.handle, id, offset, start), true);
  }

  private splitWithBreak(
    call: (bridge: DocxEditorExports["DocxSessionBridge"], fullId: string, offset: number) => string | undefined,
    sections: boolean,
  ): boolean {
    const block = this.activeBlock;
    if (this.closed || !block || this.isBandBlock(block)) return false;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    const idx = this.blockIndex(block);
    const raw = caretOffsetIn(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return false;
    const offset = trimmedSplitOffset(block, raw ?? (block.textContent ?? "").length);
    const json = call(this.exports.DocxSessionBridge, fullId, offset);
    if (json === undefined) return false;
    const res = this.parseEdit(json);
    if (!res.success) return false;
    // A section break changes the section wrappers (and page view's page sizes), which only
    // the full render lays out; a page or column break is two paragraphs, which reconcile does.
    this.refreshAfter(block, idx + 1, false, sections);
    const created = res.created?.[0];
    if (created) this.options.onEdit?.({ anchorId: created.id, unid: created.unid });
    return true;
  }

  // ─── Pictures: selection, size, wrapping, position, alt text ─────────

  private pictureOverlay: PictureOverlay | null = null;
  private selectedPictureId: string | null = null;

  private overlay(): PictureOverlay {
    if (!this.pictureOverlay) {
      this.pictureOverlay = new PictureOverlay(this.container.ownerDocument, {
        resize: (wr, hr) => this.resizePictureBy(wr, hr),
        move: (dx, dy) => this.movePictureBy(dx, dy),
      });
    }
    return this.pictureOverlay;
  }

  /** Every image occurrence the engine knows, in document order. */
  private listImages(): ImageOccurrence[] {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.ListImages !== "function") return [];
    try {
      return JSON.parse(bridge.ListImages(this.handle, 63 /* ProjectionScopes.All */)) as ImageOccurrence[];
    } catch {
      return [];
    }
  }

  /** The rendered pictures of a block, outermost only (a vector frame's own parts are not). */
  private picturesIn(block: HTMLElement): HTMLElement[] {
    const own = Array.from(block.querySelectorAll<HTMLElement>(PICTURE_SELECTOR))
      .filter((p) => !p.parentElement?.closest(PICTURE_SELECTOR));
    // Page view lifts floating drawings out to the page box; it remembers where each stood.
    const lifted = Array.from(this.container.querySelectorAll<HTMLElement>("[data-docx-drawing-anchor]"))
      .filter((anchor) => { const home = liftedDrawingHomes.get(anchor); return !!home && block.contains(home.parent); })
      .flatMap((anchor) => Array.from(anchor.querySelectorAll<HTMLElement>(PICTURE_SELECTOR))
        .filter((p) => !p.parentElement?.closest(PICTURE_SELECTOR)));
    if (lifted.length === 0) return own;
    const doc = block.ownerDocument;
    const point = (p: HTMLElement): Range => {
      const r = doc.createRange();
      const home = this.pictureHome(p);
      if (home) r.setStart(home.parent, home.index); else r.setStartBefore(p);
      return r;
    };
    return [...own, ...lifted].sort((a, b) => point(a).compareBoundaryPoints(Range.START_TO_START, point(b)));
  }

  /** Where a picture page view lifted out of its paragraph stood in it, or null. */
  private pictureHome(picture: HTMLElement): { parent: Node; index: number } | null {
    for (let el: HTMLElement | null = picture; el && el !== this.container; el = el.parentElement) {
      const home = liftedDrawingHomes.get(el);
      if (home) return home.parent.isConnected && this.container.contains(home.parent) ? home : null;
    }
    return null;
  }

  /** A collapsed range just before (or after) a picture, at its place in its paragraph. */
  private rangeAtPicture(picture: HTMLElement, before: boolean): Range {
    const range = picture.ownerDocument.createRange();
    const home = this.pictureHome(picture);
    if (home) range.setStart(home.parent, Math.min(home.index, home.parent.childNodes.length));
    else if (before) range.setStartBefore(picture);
    else range.setStartAfter(picture);
    range.collapse(true);
    return range;
  }

  /** A picture's place in its block: its content offset (pictures take none) and its rank among
   *  the pictures at that offset, which is how the engine's occurrences line up with it. */
  private picturePlace(picture: HTMLElement, block: HTMLElement): { offset: number; rank: number } {
    const offsetOf = (p: HTMLElement): number => {
      const home = this.pictureHome(p);
      if (home) return contentOffsetOf(block, home.parent, home.index);
      const parent = p.parentNode!;
      return contentOffsetOf(block, parent, Array.prototype.indexOf.call(parent.childNodes, p));
    };
    const offset = offsetOf(picture);
    const same = this.picturesIn(block).filter((p) => offsetOf(p) === offset);
    return { offset, rank: Math.max(0, same.indexOf(picture)) };
  }

  private pictureAtPlace(block: HTMLElement, place: { offset: number; rank: number }): HTMLElement | null {
    const same = this.picturesIn(block).filter((p) => this.picturePlace(p, block).offset === place.offset);
    return same[place.rank] ?? null;
  }

  /** The engine's occurrence for a picture at `place` in the block whose anchor is `anchorId`. */
  private occurrenceAt(anchorId: string, place: { offset: number; rank: number }): ImageOccurrence | null {
    const candidates = this.listImages().filter((i) => i.anchorId === anchorId && i.span.start === place.offset);
    return candidates[place.rank] ?? null;
  }

  /** A click on a picture selects it (left or right button). True when it did. */
  private pictureMouseDown(event: MouseEvent): boolean {
    const target = event.target instanceof Element ? event.target : null;
    if (this.pictureOverlay?.owns(target)) return true;
    const picture = target?.closest<HTMLElement>(PICTURE_SELECTOR) ?? null;
    const block = picture ? this.editableBlockOf(this.pictureHome(picture)?.parent ?? picture) : null;
    if (!picture || !block || (event.button !== 0 && event.button !== 2)) {
      // A click elsewhere in the document drops the selection; one in the ribbon or a dialog keeps it.
      if (this.selectedPictureId && event.button === 0 && target && this.container.contains(target)) this.deselectPicture();
      return false;
    }
    if (event.button === 0) event.preventDefault();
    const selected = (this.selectedPictureId && this.pictureOverlay?.picture === picture)
      || this.selectPictureElement(picture, block);
    // A floating picture page view placed can be dragged somewhere else straight away.
    const shown = this.pictureOverlay?.picture;
    if (selected && event.button === 0 && shown && this.movablePicture(shown)) this.pictureOverlay!.startMove(event);
    return selected;
  }

  /** Where a floating picture in page view stands (its page box and the paragraph it is
   *  anchored in), when it can be moved by dragging; null otherwise (continuous view places
   *  floating pictures by CSS float, not by position). */
  private movablePicture(picture: HTMLElement): { page: HTMLElement; paragraph: HTMLElement } | null {
    const page = picture.closest<HTMLElement>("[data-page-number][data-section-index]");
    if (!page || !this.container.contains(page)) return null;
    const occurrence = this.selectedPictureId ? this.listImages().find((i) => i.id === this.selectedPictureId) : undefined;
    if (occurrence?.placement !== "floating" || !pictureInfoOf(occurrence).can.layout) return null;
    // Page view draws a top-and-bottom picture in its paragraph's flow, wherever its position says:
    // a drag would change the document without moving it on screen.
    const mode = occurrence.floatingLayout?.wrapMode;
    if (mode === "top_and_bottom" || mode === "unknown") return null;
    // A lifted picture is measured from the paragraph its hidden marker stands in.
    const homeParent = this.pictureHome(picture)?.parent;
    const at = homeParent ? (homeParent instanceof Element ? homeParent : homeParent.parentElement) : picture;
    const paragraph = at?.closest<HTMLElement>("p, h1, h2, h3, h4, h5, h6, [data-anchor]");
    return paragraph ? { page, paragraph } : null;
  }

  /**
   * Move the selected floating picture by a dragged distance: across from the page's edge, and
   * down from its own paragraph (so it still travels with its text), as Word keeps it. One undo step.
   */
  private movePictureBy(dxPixels: number, dyPixels: number): boolean {
    const picture = this.pictureOverlay?.picture;
    const id = this.selectedPictureId;
    const occurrence = id ? this.listImages().find((i) => i.id === id) : undefined;
    const where = picture ? this.movablePicture(picture) : null;
    if (!picture || !occurrence?.floatingLayout || !where) return false;
    const info = pictureInfoOf(occurrence);
    const rect = picture.getBoundingClientRect();
    const pixelsPerPoint = info.widthPoints && rect.width > 0 ? rect.width / info.widthPoints : 0;
    if (!(pixelsPerPoint > 0)) return false;
    // Across from the page's edge; down from the same paragraph box page view measures from.
    const left = (rect.left + dxPixels - where.page.getBoundingClientRect().left) / pixelsPerPoint;
    const top = (rect.top + dyPixels - where.paragraph.getBoundingClientRect().top) / pixelsPerPoint;
    const emu = (points: number): number => Math.round(points * 12700);
    const base = writableLayout(occurrence.floatingLayout);
    // Text wraps round a square, tight or through picture as round a float at a column edge,
    // so across the page it goes to whichever edge it was dropped nearer; down, it goes where it was dropped.
    const pageRect = where.page.getBoundingClientRect();
    const across: Partial<FloatingImageLayout> = base.wrapMode === "square" || base.wrapMode === "tight" || base.wrapMode === "through"
      ? { horizontalRelativeFrom: "column", horizontalOffsetEmu: null,
          horizontalAlignment: rect.left + dxPixels + rect.width / 2 <= pageRect.left + pageRect.width / 2 ? "left" : "right" }
      : { horizontalRelativeFrom: "page", horizontalOffsetEmu: emu(left), horizontalAlignment: null };
    const layout: FloatingImageLayout = {
      ...base, ...across,
      verticalRelativeFrom: "paragraph", verticalOffsetEmu: emu(top), verticalAlignment: null,
    };
    return this.changePicture((bridge, imageId) => bridge.SetImageFloatingLayout?.(this.handle, imageId, JSON.stringify(layout)));
  }

  private selectPictureElement(picture: HTMLElement, block: HTMLElement): boolean {
    // Commit the block's typing first: the change that follows renders from the session.
    let place = this.picturePlace(picture, block);
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    const idx = this.blockIndex(block);
    const synced = this.syncBlock(block, fullId);
    if (!synced) return false;
    fullId = synced;
    if (!block.isConnected) {
      const fresh = this.blockForAnchorId(fullId) ?? this.editableList()[idx];
      if (!fresh) return false;
      block = fresh;
    }
    const occurrence = this.occurrenceAt(fullId, place);
    const element = picture.isConnected ? picture : this.pictureAtPlace(block, place);
    if (!occurrence || !element) return false;
    place = { ...place };
    this.showPictureSelection(occurrence, element, block);
    return true;
  }

  private showPictureSelection(occurrence: ImageOccurrence, picture: HTMLElement, block: HTMLElement): void {
    this.selectedPictureId = occurrence.id;
    this.activeBlock = block;
    const info = pictureInfoOf(occurrence);
    this.overlay().show(picture, info.can.size);
    this.overlay().follow();
    if (this.movablePicture(picture)) picture.style.cursor = "move";
    // The caret waits just after the picture, so typing goes on from there.
    if (document.activeElement !== block) block.focus({ preventScroll: true });
    const sel = window.getSelection();
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(this.rangeAtPicture(picture, false));
    }
    this.options.onPictureChange?.(info);
  }

  /** Select the picture the engine calls `imageId` again, after a render replaced its element. */
  private reselectPicture(imageId: string): boolean {
    const occurrence = this.listImages().find((i) => i.id === imageId);
    const block = occurrence ? this.blockForAnchorId(occurrence.anchorId) : null;
    if (!occurrence || !block) { this.deselectPicture(); return false; }
    const rank = this.listImages()
      .filter((i) => i.anchorId === occurrence.anchorId && i.span.start === occurrence.span.start)
      .findIndex((i) => i.id === imageId);
    const picture = this.pictureAtPlace(block, { offset: occurrence.span.start, rank: Math.max(0, rank) });
    if (!picture) { this.deselectPicture(); return false; }
    this.showPictureSelection(occurrence, picture, block);
    return true;
  }

  private readonly onPictureKey = (ev: KeyboardEvent): void => {
    if (!this.selectedPictureId || this.closed) return;
    const picture = this.pictureOverlay?.picture;
    if (!picture?.isConnected) { this.deselectPicture(); return; }
    const active = document.activeElement;
    if (!active || !this.container.contains(active)) return;
    const caretBy = (before: boolean): void => {
      const sel = window.getSelection();
      if (!sel) return;
      sel.removeAllRanges();
      sel.addRange(this.rangeAtPicture(picture, before));
    };
    switch (ev.key) {
      case "Delete":
      case "Backspace":
        ev.preventDefault();
        ev.stopPropagation();
        this.removePicture();
        return;
      case "Escape":
        ev.preventDefault();
        this.deselectPicture();
        return;
      case "ArrowLeft":
      case "ArrowUp":
      case "ArrowRight":
      case "ArrowDown":
        ev.preventDefault();
        caretBy(ev.key === "ArrowLeft" || ev.key === "ArrowUp");
        this.deselectPicture();
        return;
      case "Shift": case "Control": case "Alt": case "Meta": case "ContextMenu":
        return;
      default:
        if (ev.ctrlKey || ev.metaKey || ev.altKey) return;
        this.deselectPicture();
    }
  };

  /** The selected picture, or null. */
  selectedPicture(): PictureInfo | null {
    if (!this.selectedPictureId) return null;
    const occurrence = this.listImages().find((i) => i.id === this.selectedPictureId);
    return occurrence ? pictureInfoOf(occurrence) : null;
  }

  /** Select the `index`th picture of the body (0 = the first), as a click on it does. */
  selectPicture(index = 0): boolean {
    this.assertOpen();
    let n = 0;
    for (const block of this.editableList()) {
      if (this.isBandBlock(block)) continue;
      for (const picture of this.picturesIn(block)) {
        if (n++ !== index) continue;
        return this.selectPictureElement(picture, block);
      }
    }
    return false;
  }

  /** Clear the picture selection (the caret stays where it is). */
  deselectPicture(): void {
    if (!this.selectedPictureId && !this.pictureOverlay?.picture) return;
    this.selectedPictureId = null;
    this.pictureOverlay?.hide();
    this.options.onPictureChange?.(null);
  }

  /** Run one engine change on the selected picture, re-render, and keep it selected. */
  private changePicture(
    call: (bridge: DocxEditorExports["DocxSessionBridge"], id: string, occurrence: ImageOccurrence) => string | undefined,
    keepSelected = true,
  ): boolean {
    if (this.closed) return false;
    const id = this.selectedPictureId;
    const occurrence = id ? this.listImages().find((i) => i.id === id) : undefined;
    if (!id || !occurrence) return false;
    const json = call(this.exports.DocxSessionBridge, id, occurrence);
    if (json === undefined) return false;
    const res = this.parseEdit(json);
    if (!res.success) return false;
    const block = this.blockForAnchorId(occurrence.anchorId);
    // Picture media and drawings are not in the block-render shell: only the full render shows them.
    // The selection is held quietly across it (no "deselected" for the ribbon to act on), then found again.
    this.selectedPictureId = null;
    this.remount(block ? this.blockIndex(block) : -1, false);
    if (keepSelected) this.reselectPicture(res.imageId ?? id);
    else this.deselectPicture();
    return true;
  }

  /**
   * Size the selected picture, in points. Give one side with `lockAspect` (the default) and the
   * other follows; give both to set both exactly.
   */
  setPictureSize(widthPoints?: number | null, heightPoints?: number | null, lockAspect = true): boolean {
    const w = widthPoints ?? undefined, h = heightPoints ?? undefined;
    if (w === undefined && h === undefined) return false;
    return this.changePicture((bridge, id) => bridge.SetImageDimensions?.(this.handle, id, JSON.stringify({
      ...(w !== undefined ? { widthPoints: w } : {}),
      ...(h !== undefined ? { heightPoints: h } : {}),
      preserveAspect: lockAspect && (w === undefined || h === undefined),
    })));
  }

  /** The picture's own size (Word's Reset Picture & Size, size only). */
  resetPictureSize(): boolean {
    const info = this.selectedPicture();
    if (!info?.originalWidthPoints || !info.originalHeightPoints) return false;
    return this.setPictureSize(info.originalWidthPoints, info.originalHeightPoints, false);
  }

  private resizePictureBy(widthRatio: number, heightRatio: number): void {
    const info = this.selectedPicture();
    if (!info?.widthPoints || !info.heightPoints) return;
    const round = (n: number): number => Math.round(n * 100) / 100;
    this.setPictureSize(round(info.widthPoints * widthRatio), round(info.heightPoints * heightRatio), false);
  }

  /** Word's Wrap Text for the selected picture. */
  setPictureWrap(wrap: PictureWrap): boolean {
    return this.changePicture((bridge, id, occurrence) => {
      if (wrap === "inline") {
        if (occurrence.placement === "inline") return JSON.stringify({ success: true, imageId: id });
        return bridge.SetImagePlacement?.(this.handle, id, "inline", null);
      }
      const floating = occurrence.placement === "floating";
      let base: FloatingImageLayout;
      if (floating && occurrence.floatingLayout) base = writableLayout(occurrence.floatingLayout);
      else {
        // Made to float: where it was (its paragraph, the column's edge), above the other floating
        // pictures, with Word's clearance at its sides.
        const highest = this.listImages().reduce((max, i) => Math.max(max, i.floatingLayout?.relativeHeight ?? 0), 251658240);
        base = {
          horizontalRelativeFrom: "column", horizontalOffsetEmu: 0, horizontalAlignment: null,
          verticalRelativeFrom: "paragraph", verticalOffsetEmu: 0, verticalAlignment: null,
          wrapSide: "both_sides", distanceTopEmu: 0, distanceBottomEmu: 0,
          distanceLeftEmu: FLOAT_SIDE_DISTANCE_EMU, distanceRightEmu: FLOAT_SIDE_DISTANCE_EMU,
          relativeHeight: Math.min(highest + 1, 0xfffffffe), locked: false, layoutInCell: true, allowOverlap: true,
        };
      }
      const mode = ({ square: "square", tight: "tight", through: "through", topAndBottom: "top_and_bottom",
        behind: "none", inFront: "none" } as const)[wrap];
      const layout: FloatingImageLayout = { ...base, wrapMode: mode, behindDocument: wrap === "behind" };
      if (layout.wrapSide === "unknown") layout.wrapSide = "both_sides";
      return floating
        ? bridge.SetImageFloatingLayout?.(this.handle, id, JSON.stringify(layout))
        : bridge.SetImagePlacement?.(this.handle, id, "floating", JSON.stringify(layout));
    });
  }

  /** Position and wrap a floating picture (the Layout tab of the picture dialog). */
  setPictureLayout(layout: FloatingImageLayout): boolean {
    return this.changePicture((bridge, id, occurrence) => occurrence.placement === "floating"
      ? bridge.SetImageFloatingLayout?.(this.handle, id, JSON.stringify(writableLayout(layout)))
      : bridge.SetImagePlacement?.(this.handle, id, "floating", JSON.stringify(writableLayout(layout))));
  }

  /** The selected picture's alt text (description) and title; null or "" clears one. */
  setPictureAltText(altText: string | null, title?: string | null): boolean {
    return this.changePicture((bridge, id, occurrence) => bridge.SetImageMetadata?.(this.handle, id,
      altText || null, title === undefined ? occurrence.title ?? null : title || null));
  }

  /** Put new picture bytes (base64) in place of the selected picture's, keeping its size and place. */
  replacePicture(imageBase64: string): boolean {
    return this.changePicture((bridge, id) => bridge.ReplaceImage?.(this.handle, id, imageBase64));
  }

  /** Read a picked file and put it in place of the selected picture. */
  async replacePictureFile(file: Blob): Promise<boolean> {
    const bytes = new Uint8Array(await file.arrayBuffer());
    let binary = "";
    for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return this.replacePicture(btoa(binary));
  }

  /** Delete the selected picture. */
  removePicture(): boolean {
    const id = this.selectedPictureId;
    const occurrence = id ? this.listImages().find((i) => i.id === id) : undefined;
    const ok = this.changePicture((bridge, imageId) => bridge.RemoveImage?.(this.handle, imageId), false);
    if (ok && occurrence) {
      const block = this.blockForAnchorId(occurrence.anchorId);
      if (block) {
        this.activeBlock = block;
        placeCaretAtOffset(block, occurrence.span.start);
      }
    }
    return ok;
  }

  /** Read a picked file and insert it as an inline image. */
  async insertImageFile(file: Blob, options: ImageInsertOptions = {}): Promise<boolean> {
    const bytes = new Uint8Array(await file.arrayBuffer());
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }
    return this.insertImage(btoa(binary), options);
  }

  /** Insert a table of contents (a real TOC field) before the active block. */
  insertTableOfContents(options: TableOfContentsOptions = {}): boolean {
    const block = this.activeBlock;
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || !block || typeof bridge.InsertTableOfContents !== "function") return false;
    if (this.isBandBlock(block) || block.closest("table, .footnotes, .endnotes")) return false;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return false;
    const res = this.parseEdit(
      bridge.InsertTableOfContents(this.handle, fullId, "before", JSON.stringify(options)),
    );
    if (!res.success) return false;
    this.refreshAfter(block, idx, false, /* forceRemount */ true);
    return true;
  }

  // ─── Table group (extended) ──────────────────────────────────────────

  /** Merge the active cell with `rowSpan`×`colSpan` neighbours (down and right). */
  mergeCells(rowSpan: number, colSpan: number): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.MergeCells !== "function") return;
    this.tableEdit((a) => bridge.MergeCells!(this.handle, a, rowSpan, colSpan, ""));
  }

  /** Split a merged cell back into its grid cells. */
  unmergeCells(): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.UnmergeCells !== "function") return;
    this.tableEdit((a) => bridge.UnmergeCells!(this.handle, a));
  }

  /** Set (or with `style: "none"` remove) the borders of the active cell's table. */
  setTableBorders(spec: TableBorderSpec): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetTableBorders !== "function") return;
    this.tableEdit((a) => bridge.SetTableBorders!(this.handle, a, JSON.stringify(spec)));
  }

  /** Shade the active cell (or its row/column/table) with a hex fill; `""` clears. */
  setCellShading(fill: string, scope: "cell" | "row" | "column" | "table" = "cell"): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetCellShading !== "function") return;
    this.tableEdit((a) => bridge.SetCellShading!(this.handle, a, fill.replace(/^#/, ""), scope));
  }

  /** Repeat the active cell's row at the top of every page the table spans. */
  setRepeatHeaderRow(repeat: boolean): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetRepeatHeaderRow !== "function") return;
    this.tableEdit((a) => bridge.SetRepeatHeaderRow!(this.handle, a, repeat));
  }

  /** The Table Properties dialog's values for the table, row and cell the caret is in, or null. */
  tableProperties(): TablePropertiesInfo | null {
    return this.tablePropertiesOf(this.activeBlock?.closest<HTMLElement>("td, th") ?? null);
  }

  /** Set the widths (twips) of the caret's table's grid columns, as one undo step (the ruler's
   *  column markers). */
  setTableColumnWidths(widthsTwips: number[]): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetColumnWidths !== "function" || widthsTwips.some((w) => !(w > 0))) return;
    const json = JSON.stringify(widthsTwips.map((w) => Math.round(w)));
    this.tableEdit((a) => bridge.SetColumnWidths!(this.handle, a, json));
  }

  private tablePropertiesOf(cell: HTMLElement | null): TablePropertiesInfo | null {
    const bridge = this.exports.DocxSessionBridge;
    const cellId = cell ? this.anchorIdOf(cell) : undefined;
    if (this.closed || !cellId || typeof bridge.GetTableProperties !== "function") return null;
    try {
      return JSON.parse(bridge.GetTableProperties!(this.handle, cellId)) as TablePropertiesInfo | null;
    } catch {
      return null;
    }
  }

  /** Word's Table Properties for the table the caret is in (width, alignment, indent, cell
   *  margins, alt text). One undo step. */
  setTableProperties(op: TablePropertiesOp): void {
    this.applyTableProperties({ table: op });
  }

  /** The vertical alignment (or other cell properties) of the caret's cell, row, column or table. */
  setCellProperties(op: CellPropertiesOp, scope: TableCellScope = "cell"): void {
    this.applyTableProperties({ cell: op, cellScope: scope });
  }

  /** Row options of the caret's row: height and its rule, break across pages, repeat as header. */
  setTableRowOptions(options: TableRowOptions): void {
    this.applyTableProperties({ row: options });
  }

  /** The width of the caret's column, twips (a merged cell's width is shared out over the grid
   *  columns it spans, in their proportions). */
  setColumnWidth(twips: number): void {
    this.applyTableProperties({ columnWidthTwips: twips });
  }

  /**
   * Apply the Table Properties dialog's changes to the table, row, column and cell the caret is in,
   * as one edit and one undo step. Each part is optional.
   */
  applyTableProperties(changes: {
    table?: TablePropertiesOp;
    row?: TableRowOptions;
    columnWidthTwips?: number;
    cell?: CellPropertiesOp;
    cellScope?: TableCellScope;
  }): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetTableProperties !== "function" || typeof bridge.SetCellProperties !== "function"
      || typeof bridge.SetTableRowOptions !== "function" || typeof bridge.SetColumnWidths !== "function") return;
    const info = changes.columnWidthTwips !== undefined ? this.tableProperties() : null;
    this.tableEdit((a) => {
      const tx = typeof bridge.BeginTransaction === "function" ? bridge.BeginTransaction(this.handle) : 0;
      let last = JSON.stringify({ success: true });
      let ok = true;
      const step = (json: string): void => {
        if (!ok) return;
        last = json;
        try { ok = (JSON.parse(json) as { success?: boolean }).success === true; } catch { ok = false; }
      };
      if (changes.columnWidthTwips !== undefined && info && changes.columnWidthTwips > 0) {
        const widths = info.columnWidthsTwips.slice();
        const span = widths.slice(info.columnIndex, info.columnIndex + info.columnSpan);
        const total = span.reduce((sum, w) => sum + w, 0) || span.length;
        let given = 0;
        span.forEach((w, i) => {
          const share = i === span.length - 1
            ? changes.columnWidthTwips! - given
            : Math.round((changes.columnWidthTwips! * (total === span.length ? 1 : w)) / total);
          widths[info.columnIndex + i] = Math.max(1, share);
          given += share;
        });
        step(bridge.SetColumnWidths!(this.handle, a, JSON.stringify(widths)));
      }
      if (changes.row) {
        const r = changes.row;
        step(bridge.SetTableRowOptions!(this.handle, a, r.repeatHeader ?? null, r.allowBreakAcrossPages ?? null,
          r.heightTwips ?? null, r.heightRule ?? "atLeast"));
      }
      if (changes.cell) step(bridge.SetCellProperties!(this.handle, a, JSON.stringify(changes.cell), changes.cellScope ?? "cell"));
      // The table's own properties last: SetColumnWidths sizes the table to its columns' sum.
      if (changes.table) step(bridge.SetTableProperties!(this.handle, a, JSON.stringify(changes.table)));
      if (tx) {
        if (ok) bridge.CommitTransaction!(tx);
        else bridge.RollbackTransaction?.(tx);
      }
      return last;
    });
  }

  /** Delete the whole table the caret is in. */
  deleteTable(): void {
    const block = this.activeBlock;
    const table = block?.closest<HTMLElement>("table");
    if (this.closed || !block || !table) return;
    const tableId = this.anchorIdOf(table);
    if (!tableId) return;
    const idx = this.blockIndex(block);
    const res = this.parseEdit(this.exports.DocxSessionBridge.DeleteBlock(this.handle, tableId));
    if (!res.success) return;
    this.refreshAfter(block, Math.max(0, idx - 1), true);
  }

  // ─── Review group ────────────────────────────────────────────────────

  /** How edits are being recorded right now. */
  get trackedChanges(): TrackedChangeMode {
    return this.options.trackedChanges;
  }

  /**
   * Switch tracked-changes mode mid-session (Word's "Track Changes" toggle). Undo history and
   * edits survive; the document re-renders, and unless `showRevisions` was set, revisions show
   * (or stop showing) inline with it.
   */
  setTrackedChanges(mode: TrackedChangeMode): void {
    this.assertOpen();
    if (this.options.trackedChanges === mode) return;
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetTrackedChanges !== "function") return;
    bridge.SetTrackedChanges(this.handle, mode);
    this.options.trackedChanges = mode;
    this.remount();
  }

  /** Whether the document's tracked revisions currently render inline. */
  get revisionsVisible(): boolean {
    return this.renderTrackedChanges;
  }

  /**
   * Show or hide the document's tracked revisions (Word's "All Markup" / "No Markup") without
   * changing how edits are recorded. The choice then holds across `setTrackedChanges`.
   */
  showRevisions(visible: boolean): void {
    this.assertOpen();
    const was = this.renderTrackedChanges;
    this.options.showRevisions = visible;
    if (was !== visible) this.remount();
  }

  /**
   * Change who later comments and replies are from (the gutter's included): `author`, with
   * `initials` (or, when omitted or empty, the first letter of each word of `author`).
   * Comments already in the document keep theirs.
   */
  setCommentAuthor(author: string, initials = ""): void {
    this.assertOpen();
    this.options.commentAuthor = author;
    this.options.commentInitials = initials;
  }

  /** The initials for a comment by `author`: the configured ones for the configured author. */
  private commentInitialsFor(author: string): string {
    return author === this.options.commentAuthor && this.options.commentInitials
      ? this.options.commentInitials
      : nameInitials(author);
  }

  setRevisionAuthor(author: string): void {
    this.assertOpen();
    this.options.revisionAuthor = author;
    this.exports.DocxSessionBridge.SetRevisionAuthor?.(this.handle, author);
  }

  /** Configure account ownership without changing the document or its history. */
  setRevisionIdentities(identities: readonly RevisionIdentity[] | null): void {
    this.assertOpen();
    const copied = identities?.map(identity => ({...identity})) ?? null;
    const bridge = this.exports.DocxSessionBridge;
    requireRevisionIdentitySupport(bridge, copied);
    bridge.SetRevisionIdentities?.(this.handle, JSON.stringify(copied));
    this.options.revisionIdentities = copied;
  }

  getRevisionAuthorNameStatus(author?: string | null): RevisionAuthorNameInfo {
    this.assertOpen();
    const query = this.exports.DocxSessionBridge.GetRevisionAuthorNameStatus;
    if (typeof query !== 'function') throw new Error('Revision author name queries require a renderer that supports them.');
    return JSON.parse(query(this.handle, author ?? '')) as RevisionAuthorNameInfo;
  }

  /** Every tracked revision in the document, from the session's registry. */
  listRevisions(): RevisionListEntry[] {
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || typeof bridge.ListRevisions !== "function") return [];
    try {
      const parsed = JSON.parse(bridge.ListRevisions(this.handle)) as RevisionListEntry[] | { error?: string };
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  acceptRevision(revisionId: string): boolean {
    return this.resolveRevision("AcceptRevision", revisionId);
  }

  rejectRevision(revisionId: string): boolean {
    return this.resolveRevision("RejectRevision", revisionId);
  }

  acceptAllRevisions(): boolean {
    return this.resolveRevision("AcceptAllRevisions");
  }

  rejectAllRevisions(): boolean {
    return this.resolveRevision("RejectAllRevisions");
  }

  private resolveRevision(
    op: "AcceptRevision" | "RejectRevision" | "AcceptAllRevisions" | "RejectAllRevisions",
    revisionId?: string,
  ): boolean {
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed) return false;
    const idx = this.activeBlock ? this.blockIndex(this.activeBlock) : -1;
    let json: string;
    if (op === "AcceptRevision" || op === "RejectRevision") {
      const fn = bridge[op];
      if (typeof fn !== "function" || !revisionId) return false;
      json = fn(this.handle, revisionId);
    } else {
      const fn = bridge[op];
      if (typeof fn !== "function") return false;
      json = fn(this.handle);
    }
    const res = this.parseEdit(json);
    if (!res.success) return false;
    // Resolution rewrites markup across the document; the reconciler proves what it can.
    this.reconcile(idx, false);
    return true;
  }

  /** Rendered revision marks in document order — what Previous/Next step through. */
  revisionElements(): HTMLElement[] {
    return Array.from(
      this.container.querySelectorAll<HTMLElement>(
        'ins[class*="-ins"], del[class*="-del"], ins[class*="move-to"], del[class*="move-from"], ' +
          'tr[class*="row-ins"], tr[class*="row-del"]',
      ),
    ).filter((el) => !el.closest("#pagination-staging, .docx-comment-gutter"));
  }

  /** The registry entry a rendered revision mark belongs to (matched by block, then by text). */
  revisionAt(el: HTMLElement): RevisionListEntry | null {
    const block = el.closest<HTMLElement>("[data-anchor]");
    const unid = block?.getAttribute("data-anchor");
    const candidates = this.listRevisions().filter((r) => r.resolutionStatus === "supported"
      && r.diagnostic?.code !== "inert_move_range" && r.diagnostic?.code !== "empty_wrapper_range"
      && (r.text.length > 0 || !!r.anchorId || r.affectedAnchors.length > 0));
    if (!unid) return candidates[0] ?? null;
    const inBlock = candidates.filter(
      (r) => r.anchorId?.endsWith(unid) || r.affectedAnchors?.some((a) => a.unid === unid),
    );
    if (inBlock.length === 0) return null;
    const text = (el.textContent ?? "").trim();
    return inBlock.find((r) => text && r.text.trim() === text) ?? inBlock[0];
  }

  // ─── Find & replace ──────────────────────────────────────────────────

  /**
   * Every hit of `query` in the editable blocks, in document order, with Word's find options
   * (see find-query.ts). The paragraphs of one story (the body between tables, one table cell,
   * one header or footer) are searched as one text with a paragraph mark between them, so ^p, a
   * wildcard's ^13 or a regex's \n find hits that run over several paragraphs; a hit never
   * crosses from one story into another. An invalid pattern finds nothing (`compileFind` says why).
   */
  find(query: string, options: FindOptions = {}): EditorMatch[] {
    if (!query) return [];
    const compiled = compileFind(query, options);
    if (!compiled.ok) return [];
    // Typing still in the focused block goes in first (it may re-render the block), so the hits
    // point at blocks that stay, and a replacement does not remount over text the session lacks.
    this.commitTyping();
    const re = compiled.regex;
    const out: EditorMatch[] = [];
    for (const segment of this.findSegments()) {
      const { blocks, starts, text } = segment;
      const blockAt = (pos: number): number => {
        let lo = 0, hi = starts.length - 1;
        while (lo < hi) {
          const mid = (lo + hi + 1) >> 1;
          if (starts[mid] <= pos) lo = mid; else hi = mid - 1;
        }
        return lo;
      };
      re.lastIndex = 0;
      for (let hit = re.exec(text); hit; hit = re.exec(text)) {
        if (hit[0].length === 0) {
          // A zero-length hit (^, \b, a lookaround) marks no text: step past it.
          re.lastIndex = hit.index + 1;
          if (re.lastIndex > text.length) break;
          continue;
        }
        const from = hit.index, to = hit.index + hit[0].length;
        const first = blockAt(from), last = blockAt(to);
        const start = from - starts[first];
        const found = hit;
        const expand = (replacement: string) => expandReplacement(replacement, found, options);
        if (first === last) {
          out.push({ block: blocks[first], start, length: to - from, text: hit[0], expand });
        } else {
          out.push({
            block: blocks[first], start, length: segment.lengths[first] - start,
            endBlock: blocks[last], endOffset: to - starts[last],
            between: blocks.slice(first + 1, last), text: hit[0], expand,
          });
        }
        if (out.length >= FIND_HIT_LIMIT) return out;
      }
    }
    return out;
  }

  /** Why `query` cannot be searched for with `options`, or null when it can. */
  findError(query: string, options: FindOptions = {}): string | null {
    if (!query) return null;
    const compiled = compileFind(query, options);
    return compiled.ok ? null : compiled.error;
  }

  /** The searchable stories: runs of consecutive editable blocks in the same container. */
  private findSegments(): Array<{ blocks: HTMLElement[]; starts: number[]; lengths: number[]; text: string }> {
    const blocks = this.editableList().concat(
      Array.from(this.container.querySelectorAll<HTMLElement>('[data-hf-band] [data-anchor][contenteditable="true"]')),
    );
    const segments: Array<{ blocks: HTMLElement[]; starts: number[]; lengths: number[]; text: string }> = [];
    let current: (typeof segments)[number] | null = null;
    let currentKey: Element | null | undefined;
    let previous: HTMLElement | null = null;
    for (const block of blocks) {
      const key = block.closest("td, th, [data-hf-band]");
      // A new story starts in another container, and after anything that is not a paragraph
      // (a table between two body paragraphs).
      const adjacent = previous !== null && key === currentKey && !key
        ? !this.tableBetween(previous, block)
        : key === currentKey;
      if (!current || !adjacent) {
        current = { blocks: [], starts: [], lengths: [], text: "" };
        segments.push(current);
        currentKey = key;
      }
      // An empty paragraph's placeholder space is no text (as plainBlockText reads it).
      const raw = blockStreamText(block);
      const text = trimSpaces(raw).length === 0 ? "" : raw;
      if (current.blocks.length > 0) current.text += "\r";
      current.starts.push(current.text.length);
      current.lengths.push(text.length);
      current.blocks.push(block);
      current.text += text;
      previous = block;
    }
    return segments;
  }

  /** True when a table (or anything else no find can cross) sits between two body blocks. */
  private tableBetween(a: HTMLElement, b: HTMLElement): boolean {
    const range = a.ownerDocument.createRange();
    try {
      range.setStartAfter(a);
      range.setEndBefore(b);
    } catch {
      return true;
    }
    const between = range.cloneContents();
    return between.querySelector("table") !== null;
  }

  /** The DOM range a hit covers (across blocks for a hit that runs over paragraph marks). */
  private matchRange(match: EditorMatch): Range | null {
    if (!match.endBlock) return contentRangeIn(match.block, match.start, match.length);
    if (!match.block.isConnected || !match.endBlock.isConnected) return null;
    const range = match.block.ownerDocument.createRange();
    const from = contentPositionIn(match.block, match.start);
    const to = contentPositionIn(match.endBlock, match.endOffset ?? 0);
    try {
      range.setStart(from.node, from.offset);
      range.setEnd(to.node, to.offset);
    } catch {
      return null;
    }
    return range;
  }

  /**
   * Select a match, focus its block and scroll it into view — the caret lands on the hit, so
   * typing continues in the document. A find box that is still being typed into wants
   * `showFindMatches` instead; this is the commit step (closing the bar, jumping in to edit).
   */
  selectMatch(match: EditorMatch): void {
    if (!match.block.isConnected) return;
    this.clearFindMatches();
    this.activeBlock = match.block;
    if (match.endBlock) {
      // Over paragraph marks: the editor's own selection across blocks (as Shift+Down makes).
      if (!match.endBlock.isConnected) return;
      this.selectBetween(
        contentPositionIn(match.block, match.start),
        contentPositionIn(match.endBlock, match.endOffset ?? 0),
      );
    } else {
      selectRange(match.block, match.start, match.length);
    }
    match.block.scrollIntoView({ block: "center", behavior: "smooth" });
  }

  /**
   * Paint `matches` and scroll the one at `activeIndex` into view WITHOUT moving focus or the
   * caret. This is what a find field calls on every keystroke: the document shows where the hits
   * are and rides to the current one, while the keyboard stays in the search box.
   *
   * Falls back to the document selection where the CSS Custom Highlight API is missing — still
   * without focusing the block, so the next character typed goes to the search field either way.
   */
  showFindMatches(matches: EditorMatch[], activeIndex: number): void {
    const active = matches[activeIndex];
    this.clearFindMatches();
    if (!active || !active.block.isConnected) return;
    this.activeBlock = active.block;
    const activeRange = this.matchRange(active);
    const registry = highlightRegistry();
    if (registry && activeRange) {
      ensureFindHighlightStyles(this.container.ownerDocument);
      const rest: Range[] = [];
      for (let i = 0; i < matches.length && rest.length < FIND_PAINT_LIMIT; i++) {
        if (i === activeIndex) continue;
        const range = this.matchRange(matches[i]);
        if (range) rest.push(range);
      }
      if (rest.length > 0) registry.set(FIND_HIGHLIGHT, makeHighlight(rest));
      registry.set(FIND_HIGHLIGHT_ACTIVE, makeHighlight([activeRange]));
      findPaintOwner = this;
    } else if (activeRange) {
      // No highlight registry (older Firefox/Safari): the document selection is the only paint
      // available. It is not registry-owned, so `clearFindMatches` leaves it — the next
      // `showFindMatches` or the closing `selectMatch` overwrites it.
      const sel = this.container.ownerDocument.defaultView?.getSelection();
      sel?.removeAllRanges();
      sel?.addRange(activeRange);
    }
    active.block.scrollIntoView({ block: "center", behavior: "smooth" });
  }

  /** Drop the find painting (closing the find bar, or committing a match to the caret). */
  clearFindMatches(): void {
    if (findPaintOwner !== this) return;
    const registry = highlightRegistry();
    registry?.delete(FIND_HIGHLIGHT);
    registry?.delete(FIND_HIGHLIGHT_ACTIVE);
    findPaintOwner = null;
  }

  /**
   * Replace one match. `replacement` is read as the find that made the match reads it (^p, ^&,
   * \1, $1 ...: see `expandReplacement`); a match made by hand ({block, start, length}) takes it
   * as it is. Formatting of the surrounding run is inherited. `focus: false` leaves the keyboard
   * where it is — a Replace button pressed from the find bar must not drag the caret into the
   * document mid-search.
   */
  replaceMatch(
    match: EditorMatch,
    replacement: string,
    options: { focus?: boolean } = {},
  ): boolean {
    const block = match.block;
    if (this.closed || !block.isConnected) return false;
    const text = match.expand ? match.expand(replacement) : replacement;
    if (match.endBlock || text.includes("\r")) return this.replaceMatches([{ match, text }]) > 0;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return false;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return false;
    const span = trimmedSpan(block, { start: match.start, length: match.length });
    if (span.length === 0) return false;
    const runText = streamToRunText(text);
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ReplaceTextAtSpan(
        this.handle, fullId, span.start, span.length, runText,
      ),
    );
    if (!res.success) return false;
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh && options.focus !== false) selectRange(fresh, span.start, runText.length);
    return true;
  }

  /** Replace every hit of `query` (read with `options`, as `find` reads it); returns how many
   *  were replaced. One undo step. */
  replaceAll(
    query: string,
    replacement: string,
    options: FindOptions & { focus?: boolean } = {},
  ): number {
    if (this.closed || !query) return 0;
    const matches = this.find(query, options);
    if (matches.length === 0) return 0;
    return this.replaceMatches(matches.map((match) => ({
      match, text: match.expand ? match.expand(replacement) : replacement,
    })));
  }

  /**
   * Make `items` (hits in document order, each with its text in stream characters) in the session
   * as ONE transaction — one undo step — working from the last to the first so the offsets of the
   * earlier ones stay true; then show the result: the blocks swapped in place when every hit stayed
   * in its paragraph, else the document again. Returns how many were made (0 when the engine
   * refused one: then none is).
   */
  private replaceMatches(items: Array<{ match: EditorMatch; text: string }>): number {
    const bridge = this.exports.DocxSessionBridge;
    this.commitTyping();
    const live = items.filter((it) => it.match.block.isConnected && (!it.match.endBlock || it.match.endBlock.isConnected));
    if (live.length === 0) return 0;
    if (live.some((it) => it.match.endBlock && [it.match.block, it.match.endBlock].some((b) => b.closest("table"))
      && it.match.block.closest("td, th") !== it.match.endBlock!.closest("td, th"))) return 0;
    const structural = live.some((it) => it.match.endBlock || it.text.includes("\r"));
    const ids = new Map<HTMLElement, string>();
    const results = new Map<HTMLElement, EditResultLite>();
    const tracked = this.options.trackedChanges === TrackedChangeMode.RenderInline;
    const firstIndex = this.editableList().indexOf(live[0].match.block);
    const tx = typeof bridge.BeginTransaction === "function" ? bridge.BeginTransaction(this.handle) : 0;
    let ok = true;
    let made = 0;
    try {
      const step = (json: string) => {
        const res = this.parseEdit(json);
        if (!res.success) throw new Error("the engine refused the edit");
        return res;
      };
      // Typing not yet committed goes in first, so the offsets match the session.
      const idOf = (block: HTMLElement): string => {
        const known = ids.get(block);
        if (known) return known;
        const fullId = this.anchorIdOf(block);
        const synced = fullId ? this.syncBlock(block, fullId) : undefined;
        if (!synced) throw new Error("a block could not be committed");
        ids.set(block, synced);
        return synced;
      };
      for (const it of live.slice().reverse()) {
        const m = it.match;
        const lines = it.text.split("\r").map(streamToRunText);
        const firstSpan = trimmedSpan(m.block, { start: m.start, length: m.length });
        let current = idOf(m.block);
        if (!m.endBlock) {
          if (firstSpan.length === 0 && m.length > 0) continue;
          if (firstSpan.length === 0 && !lines[0] && lines.length === 1) continue;
          const res = step(bridge.ReplaceTextAtSpan(this.handle, current, firstSpan.start, firstSpan.length, lines[0], ""));
          current = res.modified?.[0]?.id ?? current;
          results.set(m.block, res);
        } else {
          const last = m.endBlock;
          let lastId = idOf(last);
          const between = (m.between ?? []).map(idOf);
          if (firstSpan.length > 0 || lines[0]) {
            current = step(bridge.ReplaceTextAtSpan(this.handle, current, firstSpan.start, firstSpan.length, lines[0], ""))
              .modified?.[0]?.id ?? current;
          }
          const lastSpan = trimmedSpan(last, { start: 0, length: m.endOffset ?? 0 });
          if (lastSpan.length > 0) {
            lastId = step(bridge.ReplaceTextAtSpan(this.handle, lastId, 0, lastSpan.length, "", "")).modified?.[0]?.id ?? lastId;
          }
          if (between.length > 0) {
            if (typeof bridge.DeleteRange === "function") step(bridge.DeleteRange(this.handle, between[0], lastId));
            else for (const id of between) step(bridge.DeleteBlock(this.handle, id));
          }
          current = step(bridge.MergeParagraphs(this.handle, current, lastId)).modified?.[0]?.id ?? current;
          // MergeParagraphs puts a space between the two when neither side has one; a replaced
          // paragraph mark joins the two sides as they are (Word's ^p -> nothing).
          const kept = blockContentText(m.block).slice(0, firstSpan.start) + lines[0];
          const rest = blockContentText(last).slice(lastSpan.start + lastSpan.length);
          if (!tracked && kept && rest && !/\s$/.test(kept) && !/^\s/.test(rest)) {
            current = step(bridge.ReplaceTextAtSpan(this.handle, current, kept.length, 1, "", "")).modified?.[0]?.id ?? current;
          }
          ids.delete(last);
        }
        // Lines after the first become paragraphs of their own.
        let head = current;
        let offset = firstSpan.start + lines[0].length;
        for (const line of lines.slice(1)) {
          const split = step(bridge.SplitParagraph(this.handle, current, offset));
          if (current === head) head = split.modified?.[0]?.id ?? head;
          const second = split.created?.[0]?.id;
          if (!second) throw new Error("the engine made no new paragraph");
          current = line ? step(bridge.ReplaceTextAtSpan(this.handle, second, 0, 0, line, "")).modified?.[0]?.id ?? second : second;
          offset = line.length;
        }
        // The hits before this one in the same paragraph go on in its head.
        ids.set(m.block, head);
        made++;
      }
    } catch {
      ok = false;
      made = 0;
    } finally {
      if (tx) {
        if (ok) bridge.CommitTransaction!(tx);
        else bridge.RollbackTransaction?.(tx);
      }
    }
    if (!ok) {
      // Without a transaction, part of the work may be in the session: show what it holds.
      if (!tx) this.remount(firstIndex);
      return 0;
    }
    if (structural) {
      this.remount(firstIndex);
    } else {
      const targets = Array.from(results, ([block, res]) => ({
        block, unid: block.getAttribute("data-anchor"), span: null, res,
      }));
      this.finishMultiBlockOp(targets, false);
    }
    const target = this.editableList()[firstIndex];
    if (target) this.options.onEdit?.({ anchorId: this.anchorIdOf(target) ?? "", unid: target.getAttribute("data-anchor") ?? "" });
    return made;
  }

  // ─── View ────────────────────────────────────────────────────────────

  /** Set the author-pinned zoom (1 = 100%). Fit-to-width still caps it on narrow hosts. */
  setZoom(scale: number): void {
    this.assertOpen();
    this.viewport.setScale(scale);
    this.gutter?.schedule();
  }

  /** The zoom the user asked for (what a zoom control shows), before fit-to-width caps it; in
   *  page width, the zoom that fits. */
  get requestedZoom(): number {
    return this.viewport.requestedScale;
  }

  /**
   * Word's Page Width: zoom so the widest page fills the editor's width, and keep it filling it
   * as the width changes (window resized, view switched), until `setZoom` sets a zoom again.
   */
  fitPageWidth(): void {
    this.assertOpen();
    this.viewport.fitPageWidth();
    this.gutter?.schedule();
  }

  /** Whether the zoom follows the editor's width (fitPageWidth). */
  get fitsPageWidth(): boolean {
    return this.viewport.fitsPageWidth;
  }

  /**
   * Word's One Page: zoom so a whole page shows (its width and its height), and keep it so as
   * the editor's size changes, until `setZoom` sets a zoom again.
   */
  fitOnePage(): void {
    this.assertOpen();
    this.viewport.fitOnePage();
    this.gutter?.schedule();
  }

  /** Whether the zoom keeps one whole page in view (fitOnePage). */
  get fitsOnePage(): boolean {
    return this.viewport.fitsOnePage;
  }

  /** Word count over the body's editable text. */
  wordCount(): number {
    let count = 0;
    for (const block of this.editableList()) {
      // Notes are not body words in either view (the paginator's per-page note block carries
      // the "page-footnotes" class; the continuous view's sections are ".footnotes"/".endnotes").
      if (block.closest('.footnotes, .endnotes, [class$="-footnotes"], [class$="-endnotes"]')) continue;
      const text = blockContentText(block);
      const words = text.match(/[\p{L}\p{N}][\p{L}\p{N}'’.-]*/gu);
      count += words ? words.length : 0;
    }
    return count;
  }

  /** In page view, the active block's page and the page total; null in continuous view. */
  pageInfo(): { page: number; total: number } | null {
    if (!this.options.paginated) return null;
    const boxes = this.editRoot.querySelectorAll<HTMLElement>(".page-box:not([data-section-filler])");
    const total = boxes.length;
    const box = this.activeBlock?.closest<HTMLElement>(".page-box");
    const page = box ? parseInt(box.dataset.pageNumber || "1", 10) : 1;
    return { page, total };
  }

  // ─── Layout: section page setup ──────────────────────────────────────

  /** The section governing the active block (or the first body block). */
  sectionInfo(): SectionInfo | null {
    if (this.closed) return null;
    const block =
      this.activeBlock && !this.isBandBlock(this.activeBlock) ? this.activeBlock : this.editableList()[0];
    const unid = block?.getAttribute("data-anchor");
    const fullId = unid ? this.unidToFullId.get(unid) : undefined;
    if (!fullId) return null;
    try {
      const parsed = JSON.parse(this.exports.DocxSessionBridge.GetSectionInfo(this.handle, fullId)) as SectionInfo | null;
      return parsed && typeof parsed.sectionUnid === "string" ? parsed : null;
    } catch {
      return null;
    }
  }

  /** Word's Page Setup on the section holding the caret: size, orientation, margins. */
  setPageSetup(op: EditorPageSetup): boolean {
    this.assertOpen();
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetPageSetup !== "function") return false;
    const info = this.sectionInfo();
    if (!info) return false;
    const res = this.parseEdit(bridge.SetPageSetup(this.handle, info.anchorId, JSON.stringify(op)));
    if (!res.success) return false;
    // Geometry is whole-document context (sheet width, page boxes).
    this.remount(this.activeBlock ? this.blockIndex(this.activeBlock) : -1, false);
    return true;
  }

  /**
   * Word's Layout > Columns: set the text columns ({@link ColumnsOp}). `applyTo` is Word's
   * "Apply to"; left out, it is `selection` when text is selected (continuous section breaks go
   * in around it, as Word does) and `section` (the caret's) otherwise. `forward` puts a
   * continuous break at the caret. False, changing nothing, where the engine refuses: columns
   * narrower than half an inch, a header, footer or text box, or a new section while changes
   * are tracked.
   */
  setColumns(op: ColumnsOp, applyTo?: ColumnsApplyTo): boolean {
    this.assertOpen();
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetColumns !== "function") return false;
    const blocks = this.selectedBlocks();
    const first = blocks[0];
    if (!first || this.isBandBlock(first)) return false;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    const selected = !!sel && sel.rangeCount > 0 && !sel.isCollapsed
      && blocks.some((b) => b.contains(sel.getRangeAt(0).startContainer) || b.contains(sel.getRangeAt(0).endContainer));
    const scope: ColumnsApplyTo = applyTo ?? (selected ? "selection" : "section");
    const last = blocks[blocks.length - 1];

    // Offsets are measured before the blocks are synced (syncing re-renders nothing here, but
    // the spans read the live DOM).
    const firstSpan = selected ? this.blockSpanForSelection(first) : null;
    const lastSpan = selected && last !== first ? this.blockSpanForSelection(last) : firstSpan;
    const caret = trimmedSplitOffset(first, caretOffsetIn(first) ?? 0);
    let startId = this.anchorIdOf(first);
    if (!startId) return false;
    startId = this.syncBlock(first, startId);
    if (!startId) return false;
    let endId = startId;
    if (last !== first) {
      const id = this.anchorIdOf(last);
      endId = id ? this.syncBlock(last, id) ?? "" : "";
      if (!endId) return false;
    }

    let startOffset = 0, endAnchor = "", endOffset = 0;
    if (scope === "forward") startOffset = caret;
    else if (scope === "selection") {
      if (!selected) return false;
      startOffset = firstSpan?.start ?? 0;
      endAnchor = endId;
      endOffset = last === first
        ? (firstSpan ? firstSpan.start + firstSpan.length : 0)
        : (lastSpan ? lastSpan.start + lastSpan.length : 0);
    }
    const res = this.parseEdit(bridge.SetColumns(this.handle, startId, startOffset, JSON.stringify(op),
      scope, endAnchor, endOffset));
    if (!res.success) return false;
    // Columns and any new sections are section wrappers, which only the full render lays out.
    this.remount(this.blockIndex(first), false);
    return true;
  }

  /** Which body section a block is in: page view stamps it on the block or on the column wrapper
   *  around it (a page box carries only its first section's index); the continuous view keeps blocks
   *  inside their section's element. */
  private sectionKeyOf(block: HTMLElement): string {
    return block.closest<HTMLElement>("[data-pagination-section]")?.dataset.paginationSection
      ?? block.closest<HTMLElement>("[data-section-index]")?.dataset.sectionIndex ?? "0";
  }

  /** The section-break anchor that ends the section `block` is in (`sec:body:<unid>`), or null in
   *  the last section (its properties end the body, not a break). */
  private sectionBreakAfter(block: HTMLElement): string | null {
    const id = this.anchorIdOf(block);
    if (!id) return null;
    try {
      const info = JSON.parse(this.exports.DocxSessionBridge.GetSectionInfo(this.handle, id)) as SectionInfo | null;
      if (!info?.sectionUnid) return null;
      const last = this.editableList().filter((b) => !this.isBandBlock(b)).pop();
      if (last && this.sectionKeyOf(last) === this.sectionKeyOf(block)) return null;
      return `sec:body:${info.sectionUnid}`;
    } catch {
      return null;
    }
  }

  /** Remove the section break `breakAnchor` (Word's Delete on a section-break mark) and re-render;
   *  the caret stays in `focusBlock`'s place. */
  private removeBreak(breakAnchor: string, focusBlock: HTMLElement | null): boolean {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.RemoveSectionBreak !== "function") return false;
    const idx = focusBlock ? this.blockIndex(focusBlock) : -1;
    const caret = focusBlock ? caretOffsetIn(focusBlock) ?? 0 : 0;
    if (focusBlock) {
      const id = this.anchorIdOf(focusBlock);
      if (id && !this.syncBlock(focusBlock, id)) return false;
    }
    const res = this.parseEdit(bridge.RemoveSectionBreak(this.handle, breakAnchor));
    if (!res.success) return false;
    const removedFocus = !!focusBlock && (res.removed ?? []).some((a) => a.unid === focusBlock.getAttribute("data-anchor"));
    this.remount(removedFocus ? Math.max(0, idx - 1) : idx, false);
    const target = this.editableList()[removedFocus ? Math.max(0, idx - 1) : idx];
    if (target) placeCaretAtOffset(target, removedFocus ? blockContentText(target).length : caret);
    return true;
  }

  /**
   * Layout > Breaks > Remove section break: remove the break that ends the caret's section (in the
   * last section, the break before it). As in Word, the text before the break joins the section
   * after it and takes that section's layout. False (nothing changed) when there is no break.
   */
  removeSectionBreak(): boolean {
    this.assertOpen();
    const block = this.activeBlock && !this.isBandBlock(this.activeBlock) ? this.activeBlock : null;
    if (!block) return false;
    let breakAnchor = this.sectionBreakAfter(block);
    if (!breakAnchor) {
      // The last section: the break before it ends the section before.
      const body = this.editableList().filter((b) => !this.isBandBlock(b));
      const key = this.sectionKeyOf(block);
      const first = body.findIndex((b) => this.sectionKeyOf(b) === key);
      const before = first > 0 ? body[first - 1] : null;
      breakAnchor = before ? this.sectionBreakAfter(before) : null;
    }
    return breakAnchor ? this.removeBreak(breakAnchor, block) : false;
  }

  /** The line numbering of the section holding the caret, or null when it has none. */
  lineNumbering(): LineNumbering | null {
    return this.sectionInfo()?.lineNumbering ?? null;
  }

  /**
   * Word's Line Numbers: `lineNumbering` (`countBy: 0` turns them off) for the caret's section, or
   * with `applyTo: "document"` for every section, as one edit. They show in the page view.
   */
  setLineNumbering(lineNumbering: LineNumbering, applyTo: "section" | "document" = "section"): boolean {
    this.assertOpen();
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.SetPageSetup !== "function") return false;
    let anchors: string[];
    if (applyTo === "document") {
      // The first body block of each section stands for it: page view stamps a block's section on
      // it, the continuous view keeps blocks inside their section's element.
      const seen = new Map<string, string>();
      for (const block of this.editableList()) {
        if (this.isBandBlock(block) || block.closest("td, th")) continue;
        const key = this.sectionKeyOf(block);
        const id = this.anchorIdOf(block);
        if (id && !seen.has(key)) seen.set(key, id);
      }
      anchors = [...seen.values()];
    } else {
      const info = this.sectionInfo();
      anchors = info ? [info.anchorId] : [];
    }
    if (!anchors.length) return false;
    const op = JSON.stringify({ lineNumbering });
    const tx = typeof bridge.BeginTransaction === "function" ? bridge.BeginTransaction(this.handle) : 0;
    let ok = true;
    for (const anchor of anchors) {
      if (!this.parseEdit(bridge.SetPageSetup(this.handle, anchor, op)).success) { ok = false; break; }
    }
    if (tx) {
      if (ok) bridge.CommitTransaction!(tx);
      else bridge.RollbackTransaction?.(tx);
    }
    if (!ok) return false;
    this.remount(this.activeBlock ? this.blockIndex(this.activeBlock) : -1, false);
    return true;
  }

  // ─── Header & footer (extended) ──────────────────────────────────────

  /** Word's "Different first page" (`first`) / "Different odd & even pages" (`even`). Under
   *  tracked changes `even` is refused (reported through `onEditFailed`, returns false). */
  setHeaderFooterKindEnabled(kind: "first" | "even", enabled: boolean): boolean {
    this.assertOpen();
    return this.region?.setKindEnabled(kind, enabled) ?? false;
  }

  headerFooterKindEnabled(kind: "first" | "even"): boolean {
    return this.region?.kindEnabled(kind) ?? false;
  }

  /** "First Page Header", "Footer", … for the story a band (or the active page area) shows. */
  headerFooterStoryLabel(which: BandWhich): string {
    return this.region?.storyLabel(which) ?? (which === "header" ? "Header" : "Footer");
  }

  /** Put the caret in the header or footer (Word's "Go to Header / Go to Footer"). */
  goToHeaderFooter(which: BandWhich): boolean {
    this.assertOpen();
    return this.region?.focusStory(which, this.activeBlock) ?? false;
  }

  /** Leave the header/footer and put the caret back in the body. */
  closeHeaderFooter(): void {
    this.assertOpen();
    this.region?.close();
    const body = this.editableList();
    const target =
      this.activeBlock && !this.isBandBlock(this.activeBlock) && this.activeBlock.isConnected
        ? this.activeBlock
        : body[0];
    if (target) placeCaretAtOffset(target, 0);
  }

  /** Which story the caret is in, or null in the body. */
  get activeStoryKind(): BandWhich | null {
    return this.activeStory;
  }

  // ─── Introspection for chrome ────────────────────────────────────────

  /** The document's style definitions (for a styles gallery); empty on older bundles. */
  styles(): StyleInfo[] {
    const bridge = this.exports.DocxSessionBridge;
    if (this.closed || typeof bridge.ListStyles !== "function") return [];
    try {
      const parsed = JSON.parse(bridge.ListStyles(this.handle)) as StyleInfo[] | { error?: string };
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  /** The active block's paragraph style id (from the render), or null. */
  styleAtCaret(): string | null {
    const info = this.paragraphFormatting();
    return info?.effectiveParagraph.styleId ?? info?.directParagraph.styleId ?? null;
  }

  /** Which inline formats the current selection carries — for ribbon button highlighting. */
  queryFormatState(): Record<FormatKey, boolean> {
    const block = this.activeBlock ?? this.editRoot;
    return {
      bold: selectionHasFormat("bold", block),
      italic: selectionHasFormat("italic", block),
      underline: selectionHasFormat("underline", block),
      strike: selectionHasFormat("strike", block),
      code: selectionHasFormat("code", block),
      superscript: selectionHasFormat("superscript", block),
      subscript: selectionHasFormat("subscript", block),
    };
  }

  // ─── Commands by name (editor-commands.ts) ───────────────────────────

  /**
   * Run an editing command by name, with its arguments as a plain object, e.g.
   * `exec("fontSize", { pts: 14 })` or `exec("insertText", { text: "Dear Sir,\n" })`. Returns the
   * command's result (a count, some text; null for most). For hosts that drive the editor from
   * outside: a scripting bridge, a command palette, a macro recorder. Throws an
   * `EditorCommandError` (`unknown_command`, `bad_arguments`, `closed`, `refused`).
   */
  exec(name: string, args: EditorCommandArgs = {}): unknown {
    if (this.closed) throw new EditorCommandError("closed", name, `${name}: the editor is closed`);
    return runEditorCommand(this, name, args);
  }

  /** The commands `exec` knows: name, group, description, arguments and result. */
  static commands(): EditorCommandInfo[] {
    return editorCommands();
  }

  /**
   * Type `text` at the caret in the active block, replacing the selection there, through the
   * engine (one undo step, not simulated keystrokes). Each line break ("\n", "\r\n") starts a
   * new paragraph, as Enter does. With no caret in the active block the text goes at its end.
   * False when there is no active block, the selection spans paragraphs, or the engine refused.
   */
  insertText(text: string): boolean {
    if (this.closed || !this.activeBlock) return false;
    const lines = text.replace(/\r\n?/g, "\n").split("\n");
    for (let i = 0; i < lines.length; i++) {
      if (i > 0) {
        const el: HTMLElement | null = this.activeBlock;
        if (!el?.isConnected) return false;
        if (caretOffsetIn(el) === null) placeCaretAtOffset(el, blockContentText(el).length);
        this.splitAtCaret(el);
        // (the split leaves the caret at the start of the new paragraph, which is now the active one)
        const next = this.focusedBlock();
        if (!next || next === el) return false;
        this.activeBlock = next;
      }
      if (lines[i] !== "" && !this.insertInBlock(lines[i])) return false;
    }
    return true;
  }

  // ─── Paste ───────────────────────────────────────────────────────────

  /**
   * Paste `content` at the caret in the active block, replacing the selection there, as one undo
   * step: its HTML through the engine (paragraphs, headings, lists, tables, links and pictures, as
   * Word pastes them), else its text, else its picture. `formatting` defaults to
   * `options.pasteFormatting` ("merge"); "text" pastes the text only. A selection across
   * paragraphs is refused. Resolves false when nothing was pasted.
   */
  async pasteContent(content: PasteContent, formatting?: PasteFormatting): Promise<boolean> {
    if (this.closed) return false;
    const mode = formatting ?? this.options.pasteFormatting ?? "merge";
    const bridge = this.exports.DocxSessionBridge;
    const at = this.pasteTarget();
    if (!at) return false;
    const html = mode === "text" ? "" : content.html ?? "";
    if (html && typeof bridge.InsertHtml === "function") {
      const cleaned = cleanPastedHtml(html, GENERATED_CHROME_SELECTOR);
      if (!cleaned.empty) {
        const images = new Map<string, string | null>();
        if (cleaned.pendingImages.length) {
          // A picture copied from a page comes as its file as well: a lone picture takes those bytes.
          const lone = content.image && cleaned.pendingImages.length === 1 && !(cleaned.doc.body.textContent ?? "").trim()
            ? await blobToDataUrl(content.image) : null;
          for (const src of cleaned.pendingImages) images.set(src, lone ?? await this.fetchPasteImage(src));
          if (!this.stillAt(at)) return false;
        }
        const res = this.parseEdit(bridge.InsertHtml(this.handle, at.fullId, at.start, at.length, serializePaste(cleaned, images),
          JSON.stringify({ formatting: mode === "keep" ? "keep" : "merge", css: cleaned.css })), { expectRejection: !!content.text });
        if (res.success) return this.finishPaste(at, res, cleaned.hasImages);
        // The engine could not take this HTML: its text still pastes, and the host hears why.
        if (content.text) this.reportEditFailure(res.error);
      }
    }
    if (content.text && typeof bridge.InsertPastedText === "function") {
      const res = this.parseEdit(bridge.InsertPastedText(this.handle, at.fullId, at.start, at.length, pastedPlainText(content.text)));
      return res.success && this.finishPaste(at, res, false);
    }
    if (content.image && at.length === 0) {
      this.activeBlock = at.block;
      return await this.insertImageFile(content.image);
    }
    return false;
  }

  /** Paste HTML at the caret (as `pasteContent`). */
  pasteHtml(html: string, formatting?: PasteFormatting): Promise<boolean> {
    return this.pasteContent({ html }, formatting);
  }

  /** Paste plain text at the caret, taking the formatting where it lands; each line a paragraph. */
  pasteText(text: string): Promise<boolean> {
    return this.pasteContent({ text }, "text");
  }

  /**
   * Paste what the system clipboard holds (the ribbon's Paste): read with the async Clipboard
   * API, which the page must be allowed to use (a WebView2 host grants "clipboard-read").
   */
  async pasteFromClipboard(formatting?: PasteFormatting): Promise<boolean> {
    if (this.closed) return false;
    const clip = typeof navigator !== "undefined" ? navigator.clipboard : undefined;
    const content: PasteContent = {};
    try {
      if (clip && typeof clip.read === "function") {
        for (const item of await clip.read()) {
          if (content.html === undefined && item.types.includes("text/html")) content.html = await (await item.getType("text/html")).text();
          if (content.text === undefined && item.types.includes("text/plain")) content.text = await (await item.getType("text/plain")).text();
          const imageType = item.types.find((t) => t.startsWith("image/"));
          if (!content.image && imageType) content.image = await item.getType(imageType);
        }
      } else if (clip) content.text = await clip.readText();
    } catch (error) {
      this.reportEditFailure({ message: `The clipboard could not be read: ${(error as Error)?.message ?? error}` } as EditResultLite["error"]);
      return false;
    }
    return this.pasteContent(content, formatting);
  }

  /** Where a paste lands: the block holding the selection (else the active one), its typing
   *  committed, the selection measured in the engine's offsets. */
  private pasteTarget(): PasteTarget | null {
    const selection = this.container.ownerDocument.defaultView?.getSelection();
    const block = (selection?.anchorNode ? this.editableBlockOf(selection.anchorNode) : null) ?? this.activeBlock;
    if (!block?.isConnected) return null;
    // (the live selection decides: the remembered cross-block one may be stale once the caret moved)
    const crossBlock = !!selection && !selection.isCollapsed && selection.focusNode
      && this.editableBlockOf(selection.focusNode) !== block;
    if (crossBlock) {
      this.reportEditFailure({ message: "Pasting over a selection that spans paragraphs is not supported yet: select within one paragraph, or delete the selection first" } as EditResultLite["error"]);
      return null;
    }
    const unid = block.getAttribute("data-anchor");
    let fullId = unid ? this.anchorIdOf(block) : undefined;
    if (!unid || !fullId) return null;
    const caret = caretOffsetIn(block);
    const span = selectionSpanIn(block)
      ?? { start: trimmedSplitOffset(block, caret ?? blockContentText(block).length), length: 0 };
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return null;
    const live = block.isConnected ? block : this.blockForAnchorId(fullId);
    if (!live) return null;
    return { block: live, unid: live.getAttribute("data-anchor") ?? unid, fullId, start: span.start, length: span.length, text: blockContentText(live) };
  }

  /** True when the paste's block is as it was when the paste began (pictures took a while). */
  private stillAt(at: PasteTarget): boolean {
    return at.block.isConnected && blockContentText(at.block) === at.text;
  }

  /** Show the pasted content and put the caret at its end. */
  private finishPaste(at: PasteTarget, res: EditResultLite, pictures: boolean): boolean {
    const caret = res.caret;
    if (!res.created?.length && !pictures) {
      const fresh = this.swapBlock(at.block, at.unid, res.modified?.[0]) ?? (at.block.isConnected ? at.block : null);
      if (fresh) {
        this.activeBlock = fresh;
        placeCaretAtOffset(fresh, caret?.offset ?? 0);
      }
    } else {
      // Image parts are not in the block-render shell: only a full render shows pasted pictures.
      this.refreshAfter(at.block, this.blockIndex(at.block), false, pictures);
      const el = caret ? this.blockForAnchorId(caret.anchorId) : null;
      if (el) {
        this.activeBlock = el;
        placeCaretAtOffset(el, caret!.offset);
      }
    }
    const first = res.modified?.[0];
    this.options.onEdit?.({ anchorId: first?.id ?? at.fullId, unid: first?.unid ?? at.unid });
    return true;
  }

  /** The editable block an anchor id names, after a render. */
  private blockForAnchorId(anchorId: string): HTMLElement | null {
    const unid = anchorId.slice(anchorId.lastIndexOf(":") + 1);
    return this.editableList().find((el) => el.getAttribute("data-anchor") === unid || this.anchorIdOf(el) === anchorId) ?? null;
  }

  /** A pasted picture's bytes as a `data:` URL: fetched from the web, else from the host. */
  private async fetchPasteImage(src: string): Promise<string | null> {
    if (/^https?:/i.test(src)) {
      const abort = new AbortController();
      const timer = setTimeout(() => abort.abort(), 5000);
      try {
        const response = await fetch(src, { signal: abort.signal });
        const blob = response.ok ? await response.blob() : null;
        if (blob?.type.startsWith("image/")) return await blobToDataUrl(blob);
      } catch {
        /* a page that does not allow it (CORS): the host may still fetch it */
      } finally {
        clearTimeout(timer);
      }
    }
    try {
      return (await this.options.resolvePasteImage?.(src)) ?? null;
    } catch {
      return null;
    }
  }

  /** One line of insertText: replace the selection, or insert at the caret, in the active block. */
  private insertInBlock(text: string): boolean {
    const block = this.activeBlock;
    if (!block?.isConnected || this.selectedBlocks().length > 1) return false;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return false;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return false;
    // (measured before the sync, which commits pending typing; both in the committed text's offsets)
    const caret = caretOffsetIn(block);
    const span = selectionSpanIn(block)
      ?? { start: trimmedSplitOffset(block, caret ?? blockContentText(block).length), length: 0 };
    const { leading } = trimBounds(block);
    const committed = blockContentText(block).slice(leading);
    fullId = this.syncBlock(block, fullId);
    if (!fullId) return false;
    const bridge = this.exports.DocxSessionBridge;
    let start = span.start;
    let res = this.parseEdit(bridge.ReplaceTextAtSpan(this.handle, fullId, span.start, span.length, text),
      span.length === 0 ? { expectRejection: true } : undefined);
    if (!res.success && span.length === 0 && committed.length > 0) {
      // The engine inserts only at a run boundary or inside a plain run (not inside a link or a
      // field): replace a neighbouring character with itself plus the text, as typing does.
      if (start > 0) {
        start -= 1;
        res = this.parseEdit(bridge.ReplaceTextAtSpan(this.handle, fullId, start, 1, committed[start] + text));
        start += 1;
      } else {
        res = this.parseEdit(bridge.ReplaceTextAtSpan(this.handle, fullId, 0, 1, text + committed[0]));
      }
    } else if (!res.success && span.length === 0) {
      this.reportEditFailure(res.error);
    }
    if (!res.success) return false;
    const fresh = this.swapBlock(block, unid, res.modified?.[0]) ?? (block.isConnected ? block : null);
    if (!fresh) return false;
    this.activeBlock = fresh;
    placeCaretAtOffset(fresh, start + text.length);
    return true;
  }

  /** The editable block holding the focus, if any (after a command that moved it). */
  private focusedBlock(): HTMLElement | null {
    const active = this.container.ownerDocument.activeElement;
    return active instanceof HTMLElement && active.isContentEditable && active.hasAttribute("data-anchor")
      && this.container.contains(active) ? active : null;
  }

  /**
   * The body's text: a line for each paragraph and table cell (in document order; notes, headers
   * and footers left out). An empty paragraph is an empty line.
   */
  documentText(): string {
    if (this.closed) return "";
    return this.editableList()
      .filter((b) => !b.closest('.footnotes, .endnotes, [class$="-footnotes"], [class$="-endnotes"]'))
      .map((b) => plainBlockText(b))
      .join("\n");
  }

  /** The text of the paragraph (or cell) at the caret; "" when there is none. */
  paragraphText(): string {
    return this.closed || !this.activeBlock ? "" : plainBlockText(this.activeBlock);
  }

  /** The selected text, a line for each paragraph it covers; "" when nothing is selected in the
   *  document. */
  selectedText(): string {
    if (this.closed) return "";
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return "";
    const range = sel.getRangeAt(0);
    if (!this.container.contains(range.commonAncestorContainer)) return "";
    return this.selectedBlocks().map((b) => {
      const text = blockContentText(b);
      const from = b.contains(range.startContainer) ? contentOffsetOf(b, range.startContainer, range.startOffset) : 0;
      const to = b.contains(range.endContainer) ? contentOffsetOf(b, range.endContainer, range.endOffset) : text.length;
      return text.slice(Math.min(from, to), Math.max(from, to)).replace(/\u00a0/g, " ");
    }).join("\n");
  }

  /** Put the caret at the start or end of the body or of the active paragraph. False when there
   *  is no paragraph to go to. */
  moveCaret(to: "documentStart" | "documentEnd" | "paragraphStart" | "paragraphEnd"): boolean {
    if (this.closed) return false;
    const blocks = this.editableList();
    const target = to === "documentStart" ? blocks[0]
      : to === "documentEnd" ? blocks[blocks.length - 1]
        : this.activeBlock;
    if (!target?.isConnected) return false;
    this.activeBlock = target;
    const { leading } = trimBounds(target);
    placeCaretAtOffset(target, to.endsWith("Start") ? leading : blockContentText(target).length);
    return true;
  }

  /** Re-render one block from the live session by EditResult ref, swapping it in place. */
  /** The drawings page view lifted out of `block` into its page box, in their paragraph order. */
  private liftedDrawingsOf(block: HTMLElement): HTMLElement[] {
    const doc = block.ownerDocument;
    const point = (el: HTMLElement): Range => {
      const home = liftedDrawingHomes.get(el)!;
      const r = doc.createRange();
      r.setStart(home.parent, Math.min(home.index, home.parent.childNodes.length));
      return r;
    };
    return Array.from(this.container.querySelectorAll<HTMLElement>("[data-docx-drawing-anchor]"))
      .filter((el) => { const home = liftedDrawingHomes.get(el); return !!home && block.contains(home.parent) && !block.contains(el); })
      .sort((a, b) => point(a).compareBoundaryPoints(Range.START_TO_START, point(b)));
  }

  /**
   * A block re-rendered in page view brings its floating drawings back inside it, while the copies
   * page view lifted out of the old block stay in the page box. Put the fresh drawings where the old
   * ones stood (the next full layout places them exactly) and drop the old ones; when they no longer
   * pair up, lay the page out again.
   */
  private reliftDrawings(lifted: HTMLElement[], fresh: Node[]): void {
    const drawings = fresh.flatMap((node) => node.nodeType === 1
      ? Array.from((node as HTMLElement).querySelectorAll<HTMLElement>("[data-docx-drawing-anchor]")) : [])
      .filter((el) => !el.parentElement?.closest("[data-docx-drawing-anchor]"));
    // Same drawings, same placement (a text edit): reuse where page view put them (the host it
    // records is page view's own note, not placement). Anything else
    // (undo of a move, a new wrap) needs the layout run again.
    const placement = (el: HTMLElement): string => Array.from(el.attributes)
      .filter((a) => a.name.startsWith("data-docx-anchor-") && a.name !== "data-docx-anchor-host").map((a) => `${a.name}=${a.value}`).sort().join(";");
    if (drawings.length !== lifted.length || drawings.some((d, i) => placement(d) !== placement(lifted[i]))) {
      for (const old of lifted) old.remove();
      if (!this.pendingDrawingLayout) {
        this.pendingDrawingLayout = true;
        queueMicrotask(() => {
          this.pendingDrawingLayout = false;
          if (this.closed) return;
          // Lay the pages out again, the caret staying where it was.
          const block = this.activeBlock?.isConnected ? this.activeBlock : null;
          const anchorId = block ? this.anchorIdOf(block) : null;
          const offset = block ? caretOffsetIn(block) : null;
          this.remount(-1, false);
          const back = anchorId ? this.blockForAnchorId(anchorId) : null;
          if (back) {
            this.activeBlock = back;
            back.focus({ preventScroll: true });
            placeCaretAtOffset(back, offset ?? 0);
          }
        });
      }
      return;
    }
    drawings.forEach((drawing, i) => {
      const old = lifted[i];
      const parent = drawing.parentNode!;
      liftedDrawingHomes.set(drawing, { parent, index: Array.prototype.indexOf.call(parent.childNodes, drawing) });
      for (const attr of Array.from(old.attributes)) drawing.setAttribute(attr.name, attr.value);
      old.replaceWith(drawing);
    });
  }
  private pendingDrawingLayout = false;

  private swapBlock(oldEl: HTMLElement, oldUnid: string, ref?: AnchorRef): HTMLElement | null {
    const inBand = this.isBandBlock(oldEl);
    const anchorId = ref?.id ?? this.anchorIdOf(oldEl);
    const newUnid = ref?.unid ?? oldUnid;
    if (!anchorId) return null;
    const fresh = this.renderInto(anchorId);
    if (fresh && this.remountSplitParagraph(oldEl, fresh)) {
      this.options.onEdit?.({ anchorId, unid: newUnid });
      return this.activeBlock;
    }
    if (!fresh || !this.replaceNode(oldEl, fresh)) return null;
    // A band block resolves through its stamped `data-hf-anchor`, never the unid map — and its
    // unid can collide with another part's, so writing it here would corrupt that entry.
    if (!inBand) {
      this.unidToFullId.delete(oldUnid);
      this.unidToFullId.set(newUnid, anchorId);
    }
    // Stamp BEFORE wiring so wireBlock's own resolution sees the authoritative id.
    if (inBand) this.region!.adoptBlock(fresh, anchorId);
    this.wireBlock(fresh);
    this.activeBlock = fresh;
    // The throwaway render numbers citation markers from 1 — repair in place.
    this.maybeRenumberNotes(fresh);
    if (inBand) this.region!.afterStoryEdit(fresh);
    this.options.onEdit?.({ anchorId, unid: newUnid });
    return fresh;
  }

  /**
   * Full-document HTML from the live session. Prefers the session-attached
   * `RenderHtml` bridge — the saved bytes never cross the JS/WASM boundary
   * (two multi-MB copies per remount on a large doc) — and falls back to
   * Save + ConvertDocxToHtmlComplete for older WASM bundles. Both paths use
   * the same option profile, so the rendered HTML is identical.
   */
  /** The editor's render profile, as the comment-aware bridge endpoints take it. */
  private editorRenderProfile(): string {
    return JSON.stringify({
      cssPrefix: this.options.cssPrefix,
      fabricateClasses: this.options.fabricateClasses,
      paginated: this.options.paginated,
      scale: this.options.scale,
      renderTrackedChanges: this.renderTrackedChanges,
      comments: this.options.comments,
    });
  }

  /**
   * Full-document HTML from the live session. Prefers the session-attached
   * `RenderEditorHtml` (comment-aware) or `RenderHtml` bridge — the saved bytes never cross the
   * JS/WASM boundary — and falls back to `ConvertDocxToHtmlComplete` over `bytes` (the opened
   * document on first paint, else a fresh save) for older WASM bundles. Every path uses the
   * same option profile, so the rendered HTML is identical.
   */
  private renderFullHtml(bytes?: Uint8Array): string {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.RenderEditorHtml === "function") {
      const html = bridge.RenderEditorHtml(this.handle, this.editorRenderProfile());
      if (html.charCodeAt(0) !== 0x7b) return html;
    }
    if (this.options.comments) {
      // Only the bytes path can render comment markup on a bundle without RenderEditorHtml.
      const source =
        bytes ??
        (typeof bridge.SaveWithAnchorIds === "function"
          ? bridge.SaveWithAnchorIds(this.handle)
          : bridge.Save(this.handle));
      return this.exports.DocumentConverter.ConvertDocxToHtmlComplete(
        ...completeArgs(
          source, this.options.cssPrefix, this.options.fabricateClasses,
          this.options.paginated, this.options.scale, this.renderTrackedChanges, true,
        ),
      );
    }
    if (typeof bridge.RenderHtmlForReview === "function") {
      const html = bridge.RenderHtmlForReview(
        this.handle,
        this.options.cssPrefix,
        this.options.fabricateClasses,
        this.options.paginated,
        this.options.scale,
        this.renderTrackedChanges,
      );
      if (html.charCodeAt(0) !== 0x7b) return html;
    }
    if (typeof bridge.RenderHtml === "function") {
      const html = bridge.RenderHtml(
        this.handle,
        this.options.cssPrefix,
        this.options.fabricateClasses,
        this.options.paginated,
        this.options.scale,
      );
      if (html.charCodeAt(0) !== 0x7b /* not an error object */) return html;
    }
    // Fallback only (no RenderHtml on this bridge, or it errored). These bytes are re-rendered and
    // discarded, and the re-render has to resolve to the SAME anchors the live session holds — a
    // content change re-derives a block's content-hashed unid, which would leave it unwired. So ask
    // for the Unid-bearing save here, and here only; DocxEditor.save() stays clean.
    const source =
      bytes ??
      (typeof bridge.SaveWithAnchorIds === "function"
        ? bridge.SaveWithAnchorIds(this.handle)
        : bridge.Save(this.handle));
    return this.exports.DocumentConverter.ConvertDocxToHtmlComplete(
      ...completeArgs(
        source, this.options.cssPrefix, this.options.fabricateClasses,
        this.options.paginated, this.options.scale, this.renderTrackedChanges, false,
      ),
    );
  }

  /** Editable BODY blocks in document order (band blocks are enumerated by `ownerRoot`). In page
   *  view a live story sits INSIDE a page box, so it is excluded here explicitly — a remount
   *  rebuilds the pages without it, and a focus index that counted it would land one block off. */
  private editableList(): HTMLElement[] {
    return Array.from(
      this.editRoot.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    ).filter((el) => !el.closest("[data-hf-band]"));
  }

  private blockIndex(el: HTMLElement): number {
    return Array.from(
      this.ownerRoot(el).querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    ).indexOf(el);
  }

  /**
   * True when an edit produced or touched a list item (kind "li"). List markers and
   * numbering CONTINUATION need whole-document context, which a single-block render lacks
   * (every item would render as "1."), so such edits re-render the whole document.
   */
  private affectsList(res: EditResultLite): boolean {
    return [...(res.modified ?? []), ...(res.created ?? [])].some((r) => r.kind === "li");
  }

  // ─── Incremental structural reconcile ─────────────────────────────────
  //
  // After a structural op (insert table/row/col, footnote, delete block, undo/redo)
  // the DOM is patched from a unit-sequence diff against the session's render plan
  // instead of remounting the whole document (~3 s of full-document conversion on a
  // 350-block file). Full remount remains the universal FALLBACK: any ambiguity,
  // unsupported bridge, paginated mode, list-membership change, or thrown error
  // lands there — correctness never depends on the diff being right.

  /** True when the bridge carries the reconcile trio and the mode allows patching. */
  private canReconcile(): boolean {
    const b = this.exports.DocxSessionBridge;
    return (
      !this.options.paginated &&
      typeof b.ListBlocks === "function" &&
      typeof b.RenderBlocksHtml === "function" &&
      typeof b.ListNotes === "function"
    );
  }

  /** The body's top-level unit nodes in document order: `[data-anchor]` elements not
   *  nested in another unit (cell paragraphs collapse into their table) and not in the
   *  notes sections. */
  private bodyUnitNodes(): HTMLElement[] {
    const all = Array.from(this.editRoot.querySelectorAll<HTMLElement>("[data-anchor]"));
    return all.filter((el) => {
      if (el.closest("section.footnotes, section.endnotes")) return false;
      const ancestor = el.parentElement?.closest("[data-anchor]");
      return !(ancestor && this.editRoot.contains(ancestor));
    });
  }

  /** The DOM diff token for a body unit node (see editor-reconcile.tokenOf). */
  private static domTokenOf(el: HTMLElement): string {
    const unid = el.getAttribute("data-anchor") ?? "";
    const sig = el.getAttribute("data-render-sig");
    return sig ? `${unid}|${sig}` : unid;
  }

  /** The kind a body unit node would have in the plan (only 'li'/'tbl' matter to the
   *  remount guard). */
  private static domKindOf(el: HTMLElement): string {
    if (el.tagName === "TABLE") return "tbl";
    return el.querySelector(":scope > [data-list-marker]") ? "li" : "p";
  }

  private static listMarkerText(el: Element | null): string | null {
    const m = el?.querySelector(":scope > [data-list-marker]");
    return m ? m.textContent : null;
  }

  /**
   * Incrementally patch the DOM from the session's render plan; falls back to
   * {@link remount} whenever it cannot prove the patch correct. Same focus contract
   * as remount.
   */
  private reconcile(focusIndex = -1, caretAtEnd = false): void {
    const inPlace = !this.canReconcile();
    if (inPlace && !this.canReconcileInPlace()) {
      this.remount(focusIndex, caretAtEnd);
      return;
    }
    try {
      if (!(inPlace ? this.reconcileInPlace() : this.reconcileCore())) {
        this.remount(focusIndex, caretAtEnd);
        return;
      }
    } catch (err) {
      this.lastReconcileFallback = `threw: ${err instanceof Error ? err.message : String(err)}`;
      this.remount(focusIndex, caretAtEnd);
      return;
    }
    if (focusIndex >= 0) {
      const blocks = this.editableList();
      const target = blocks[Math.min(focusIndex, blocks.length - 1)];
      if (target) {
        this.activeBlock = target;
        placeCaretAtOffset(target, caretAtEnd ? (target.textContent ?? "").length : 0);
      }
    }
    this.syncRegionToBody(this.activeBlock ?? undefined);
    this.gutter?.schedule();
  }

  /** Page view's reconcile needs the plan and the single-block render. */
  private canReconcileInPlace(): boolean {
    const b = this.exports.DocxSessionBridge;
    return !!this.options.paginated && typeof b.ListBlocks === "function";
  }

  /** The notes as the last full render or reconcile in page view saw them (their tokens). */
  private planNotesKey: string | null = null;

  private static notesKeyOf(plan: RenderPlan): string {
    return JSON.stringify([plan.footnotes.map(tokenOf), plan.endnotes.map(tokenOf)]);
  }

  /**
   * Page view's reconcile (after Undo or Redo, or an op that re-renders): when every change is
   * a body paragraph changed in place — an undone format or case change, say — swap those
   * paragraphs where they are, as an edit in one paragraph does in page view (the pages are not
   * laid out again). Anything a swap cannot show asks for the full render: paragraphs added,
   * removed or moved, a table or a list item changed, a paragraph split across two pages, the
   * notes changed, or very many paragraphs changed.
   */
  private reconcileInPlace(): boolean {
    this.refreshAnchorMap();
    const plan = JSON.parse(this.renderPlanJson()) as RenderPlan & { error?: string };
    if (plan.error) return this.bail(`plan error: ${plan.error}`);
    const notesKey = DocxEditor.notesKeyOf(plan);
    if (this.planNotesKey === null || notesKey !== this.planNotesKey) return this.bail("page view: notes changed (or not known)");
    // A paragraph split across pages is there twice (one piece on each page): count it once.
    const count = new Map<string, number>();
    const oldNodes: HTMLElement[] = [];
    for (const el of this.bodyUnitNodes()) {
      const unid = el.getAttribute("data-anchor") ?? "";
      const n = (count.get(unid) ?? 0) + 1;
      count.set(unid, n);
      if (n === 1) oldNodes.push(el);
    }
    const diff = diffUnits(oldNodes.map(DocxEditor.domTokenOf), plan.body);
    const n = diff.substituted.length;
    if (diff.moved.length > 0 || diff.added.length !== n || diff.removed.length !== n) {
      return this.bail(`page view: +${diff.added.length} -${diff.removed.length} ~${n} moved=${diff.moved.length}`);
    }
    if (n > 200) return this.bail(`page view: ${n} paragraphs changed`);
    const swaps: Array<{ old: HTMLElement; fresh: HTMLElement }> = [];
    for (const { oldIndex, newIndex } of diff.substituted) {
      const old = oldNodes[oldIndex]!;
      if ((count.get(old.getAttribute("data-anchor") ?? "") ?? 0) > 1) return this.bail("page view: a changed paragraph is split across pages");
      if (DocxEditor.domKindOf(old) !== "p" || ["li", "tbl", "tr", "tc"].includes(plan.body[newIndex]!.kind)) return this.bail("page view: a table or list item changed");
      const fresh = this.renderInto(plan.body[newIndex]!.id);
      if (!fresh || !fresh.hasAttribute("data-anchor")) return this.bail(`page view: unrenderable ${plan.body[newIndex]!.id}`);
      fresh.removeAttribute("xmlns");
      swaps.push({ old, fresh });
    }
    for (const { old, fresh } of swaps) {
      const wasActive = this.activeBlock === old;
      if (!this.replaceNode(old, fresh)) return this.bail("page view: swap failed");
      this.wireBlock(fresh);
      this.maybeRenumberNotes(fresh);
      if (wasActive) this.activeBlock = fresh;
    }
    if (this.activeBlock && !this.activeBlock.isConnected) this.activeBlock = null;
    this.lastReconcileFallback = null;
    return true;
  }

  /** The patch itself. Returns false to request the remount fallback. */
  private reconcileCore(): boolean {
    const bridge = this.exports.DocxSessionBridge;
    // Refresh the unid → anchor map FIRST: wiring freshly rendered nodes (wireBlock)
    // resolves through it, and the map must reflect the post-op session.
    this.refreshAnchorMap();
    const plan = JSON.parse(this.renderPlanJson()) as RenderPlan & { error?: string };
    if (plan.error) return this.bail(`plan error: ${plan.error}`);

    const oldNodes = this.bodyUnitNodes();
    const oldTokens = oldNodes.map(DocxEditor.domTokenOf);
    const oldKinds = oldNodes.map(DocxEditor.domKindOf);
    const bodyDiff = diffUnits(oldTokens, plan.body);
    // A section break added or taken away (Undo of Layout > Breaks, a merge across a break)
    // changes the section wrappers, which only the full render lays out.
    if (plan.body.some((u) => u.section !== undefined)) {
      const planSections = new Set(plan.body.map((u) => u.section ?? 0)).size;
      const shownSections = this.editRoot.querySelectorAll("[data-section-index]").length;
      if (shownSections > 0 && planSections !== shownSections) {
        return this.bail(`sections changed: ${shownSections} shown, ${planSections} in the document`);
      }
    }
    if (needsRemount(bodyDiff, plan.body, oldKinds)) {
      // Name the shape, not just the verdict: "churn" and "a list item moved in or out" are very
      // different findings when a repaint unexpectedly costs a whole-document render.
      return this.bail(
        `needsRemount (li change or churn): +${bodyDiff.added.length} -${bodyDiff.removed.length} ` +
          `~${bodyDiff.substituted.length} moved=${bodyDiff.moved.length} of ${plan.body.length}`,
      );
    }

    const fnState = this.notesDiff("footnotes", plan.footnotes);
    const enState = this.notesDiff("endnotes", plan.endnotes);
    if (fnState === null || enState === null) return this.bail("notes container unstampable/missing");

    // One batch render for everything that needs fresh HTML.
    const movedBodyNew = new Set(bodyDiff.moved.map((m) => m.newIndex));
    const addedBodyIds = bodyDiff.added
      .filter((j) => !movedBodyNew.has(j))
      .map((j) => plan.body[j].id);
    const addedNoteIds = fnState.diff.added
      .map((j) => plan.footnotes[j].id)
      .concat(enState.diff.added.map((j) => plan.endnotes[j].id));
    const allIds = addedBodyIds.concat(addedNoteIds);
    let rendered: Record<string, string | null> = {};
    if (allIds.length > 0) {
      const idsJson = JSON.stringify(allIds);
      const renderedJson = this.renderBlocksJson(idsJson);
      rendered = JSON.parse(renderedJson);
      if ((rendered as { error?: string }).error) return this.bail(`render error: ${(rendered as { error?: string }).error}`);
      for (const id of allIds) if (!rendered[id]) return this.bail(`unrenderable: ${id}`);
    }

    // A substituted list item may only swap in place if its rendered marker matches
    // the old node's — a marker change (level/membership/numbering) means sibling
    // numbers shifted too, which only a remount repaints.
    const parse = (html: string): HTMLElement | null => {
      const el = new DOMParser().parseFromString(html, "text/html").body
        .firstElementChild as HTMLElement | null;
      // Per-block converter output carries the XHTML xmlns; a full render only has it
      // on the document root, so drop it to keep reconciled DOM ≡ remounted DOM.
      el?.removeAttribute("xmlns");
      return el;
    };
    const freshBody = new Map<number, HTMLElement>();
    for (const j of bodyDiff.added.filter((j) => !movedBodyNew.has(j))) {
      const el = parse(rendered[plan.body[j].id]!);
      if (!el) return this.bail(`unparseable render: ${plan.body[j].id}`);
      freshBody.set(j, el);
    }
    for (const { oldIndex, newIndex } of bodyDiff.substituted) {
      const freshRoot = freshBody.get(newIndex);
      const oldMarker = DocxEditor.listMarkerText(oldNodes[oldIndex]);
      const newMarker = DocxEditor.listMarkerText(
        freshRoot ? DocxEditor.anchorElOf(freshRoot) : null,
      );
      if (oldMarker !== newMarker) return this.bail("substituted li marker drift");
    }

    const bodyApply = this.applyBodyDiff(oldNodes, plan.body, bodyDiff, freshBody);
    if (bodyApply !== true) return this.bail(`applyBodyDiff bail: ${bodyApply}`);
    this.applyNotesDiff("footnotes", plan.footnotes, fnState, rendered);
    this.applyNotesDiff("endnotes", plan.endnotes, enState, rendered);

    // Note chrome (marker sup text / hrefs / li values) is position-derived and NOT
    // covered by the unit diff — renumber whenever notes changed or any fresh body
    // node carries a citation marker.
    const freshHasMarker = [...freshBody.values()].some((el) =>
      el.querySelector("a.footnote-ref, a.endnote-ref"),
    );
    if (fnState.diff.added.length + fnState.diff.removed.length > 0 || freshHasMarker)
      this.renumberNoteChrome("footnote");
    if (enState.diff.added.length + enState.diff.removed.length > 0 || freshHasMarker)
      this.renumberNoteChrome("endnote");
    this.lastReconcileFallback = null;
    return true;
  }

  private bail(reason: string): false {
    this.lastReconcileFallback = reason;
    return false;
  }

  /** The generated single-child wrapper chain around a unit node (a table's alignment
   *  <div>). Climbs while the parent is an anchor-less DIV whose ONLY element child is
   *  the current node — never a section div (multi-child) or the edit root. */
  private unitWrapperOf(el: HTMLElement): HTMLElement {
    let n: HTMLElement = el;
    while (
      n.parentElement &&
      n.parentElement !== this.editRoot &&
      n.parentElement.tagName === "DIV" &&
      !n.parentElement.hasAttribute("data-section-index") &&
      !n.parentElement.hasAttribute("data-anchor") &&
      n.parentElement.childElementCount === 1
    ) {
      n = n.parentElement;
    }
    return n;
  }

  /** The `[data-anchor]` element of a fresh render root (the root itself for a leaf
   *  block, its descendant for a wrapper-shaped render like a table's align div). */
  private static anchorElOf(root: HTMLElement): HTMLElement | null {
    return root.hasAttribute("data-anchor")
      ? root
      : root.querySelector<HTMLElement>("[data-anchor]");
  }

  /** Insert/remove/swap body unit nodes per the diff. Returns `true` on success or a
   *  bail-reason string (parent ambiguity, order violation, wrapper semantics) — the
   *  session is already correct, so bailing just means a full repaint. */
  private applyBodyDiff(
    oldNodes: HTMLElement[],
    units: RenderUnit[],
    diff: UnitDiff,
    fresh: Map<number, HTMLElement>,
  ): true | string {
    // In-place substitutions first: replace at WRAPPER level so a table swaps with its
    // alignment div. A LEAF render replacing a wrapped node would break the wrapper's
    // semantics (border <div> grouping) — that is remount territory.
    const subOldByNew = new Map(diff.substituted.map((s) => [s.newIndex, s.oldIndex]));
    for (const [nj, oi] of subOldByNew) {
      const freshRoot = fresh.get(nj)!;
      const oldWrapper = this.unitWrapperOf(oldNodes[oi]);
      // A render that differs from the live node ONLY in <img> attributes (a host replacing
      // an image's media through the session — the arcade's Doom cartridge does it every
      // frame) patches those attributes on the elements already on screen instead of
      // swapping nodes. Firefox and WebKit paint a freshly inserted <img> as an empty box
      // until its data URI is decoded, so a per-frame swap strobes white between frames;
      // an in-place src change keeps the previous bitmap up until the new one is ready.
      // The live node stays wired, so only its signature stamp needs refreshing. Any
      // other difference (or a border-wrapper change) falls through to the swap below.
      const leaf = freshRoot.hasAttribute("data-anchor");
      const imagePairs = imageOnlyDelta(leaf ? oldNodes[oi] : oldWrapper, freshRoot);
      if (imagePairs) {
        for (const [live, next] of imagePairs) patchImageAttributes(live, next);
        const sig = units[nj].sig;
        if (sig) oldNodes[oi].setAttribute("data-render-sig", sig);
        continue;
      }
      if (!freshRoot.hasAttribute("data-anchor")) {
        oldWrapper.replaceWith(freshRoot); // wrapper-shaped render (table) ⇄ wrapper
      } else if (oldWrapper === oldNodes[oi]) {
        oldNodes[oi].replaceWith(freshRoot);
      } else if (
        !oldNodes[oi].hasAttribute("data-render-sig") &&
        unidOf(units[nj].id) === oldNodes[oi].getAttribute("data-anchor")
      ) {
        // Same block, leaf render, wrapped slot, and the old node carries NO signature:
        // a sig-less node is one a SINGLE-BLOCK swap placed (mount and reconcile both
        // stamp signatures), and those swaps never change border grouping — every
        // border-changing op forces a full remount. The wrapper is therefore still
        // valid; swap the leaf inside it, exactly as the single-block path itself does
        // for a bordered paragraph. A STAMPED wrapped slot whose content changed stays
        // remount territory — the change could be the border itself.
        oldNodes[oi].replaceWith(freshRoot);
      } else {
        // border-div semantics, remount
        return `leaf render into wrapped slot (unit ${units[nj].id.slice(0, 24)})`;
      }
      this.wireUnit(freshRoot, units[nj]);
    }

    const movedOldByNew = new Map(diff.moved.map((m) => [m.newIndex, m.oldIndex]));
    const movedNew = new Set(diff.moved.map((m) => m.newIndex));
    const movedOld = new Set(diff.moved.map((m) => m.oldIndex));
    const pureAdded = diff.added.filter((j) => !subOldByNew.has(j) && !movedNew.has(j));
    const pureRemoved = diff.removed.filter(
      (i) => !diff.substituted.some((s) => s.oldIndex === i) && !movedOld.has(i),
    );
    // DOMParser roots report isConnected=true because they belong to their detached
    // HTMLDocument. Track only fresh additions actually inserted into the live editor;
    // otherwise a later not-yet-applied sibling looks like a neighbor in another parent.
    const appliedFresh = new Set<number>();
    const nodeAt = (j: number): HTMLElement | null => {
      const oi = diff.keep.get(j);
      if (oi !== undefined) return oldNodes[oi];
      const movedOi = movedOldByNew.get(j);
      if (movedOi !== undefined) return oldNodes[movedOi];
      if (subOldByNew.has(j)) return fresh.get(j)!;
      const f = fresh.get(j);
      return f && appliedFresh.has(j) ? f : null;
    };

    // Preserve the established incremental insertion path: body render output may
    // legitimately use several wrapper parents (for example border groups), so it is
    // not valid to normalize the whole body against one common parent.
    for (const j of pureAdded) {
      const el = fresh.get(j)!;
      let prev: HTMLElement | null = null;
      for (let k = j - 1; k >= 0 && !prev; k--) prev = nodeAt(k);
      let next: HTMLElement | null = null;
      for (let k = j + 1; k < units.length && !next; k++) next = nodeAt(k);
      const prevW = prev ? this.unitWrapperOf(prev) : null;
      const nextW = next ? this.unitWrapperOf(next) : null;
      if (prevW && nextW && prevW.parentElement !== nextW.parentElement)
        return "insert neighbors in different parents";
      if (prevW) prevW.after(el);
      else if (nextW) nextW.before(el);
      else return "insert with no anchored neighbor";
      appliedFresh.add(j);
      this.wireUnit(el, units[j]);
    }

    // Pure removals take their now-empty generated wrappers with them. Exact-token
    // move sources remain live even though the LCS also reports them as removed.
    for (const i of pureRemoved) {
      this.unitWrapperOf(oldNodes[i]).remove();
    }

    // A block move changes one exact-token unit's slot. Move only that existing
    // wrapper beside its final neighbor; this preserves identity without disturbing
    // unrelated nodes or assuming every rendered unit shares one DOM parent.
    for (const { newIndex } of [...diff.moved].sort((a, b) => a.newIndex - b.newIndex)) {
      const moving = nodeAt(newIndex);
      if (!moving?.isConnected) return "move source is not connected";
      const movingW = this.unitWrapperOf(moving);
      let prevW: HTMLElement | null = null;
      for (let k = newIndex - 1; k >= 0 && !prevW; k--) {
        const n = nodeAt(k);
        if (n?.isConnected) prevW = this.unitWrapperOf(n);
      }
      let nextW: HTMLElement | null = null;
      for (let k = newIndex + 1; k < units.length && !nextW; k++) {
        const n = nodeAt(k);
        if (n?.isConnected) nextW = this.unitWrapperOf(n);
      }
      if (prevW && prevW !== movingW && prevW.parentElement === movingW.parentElement) {
        prevW.after(movingW);
      } else if (nextW && nextW !== movingW && nextW.parentElement === movingW.parentElement) {
        nextW.before(movingW);
      } else {
        return "move neighbors in different parents";
      }
    }
    return true;
  }

  /** Wire a freshly rendered unit root (and its nested blocks) and stamp the unit's
   *  content signature on its `[data-anchor]` element — the element the next
   *  reconcile's DOM walk reads tokens from. */
  private wireUnit(root: HTMLElement, unit: RenderUnit): void {
    const anchorEl = DocxEditor.anchorElOf(root);
    if (anchorEl && unit.sig) anchorEl.setAttribute("data-render-sig", unit.sig);
    if (anchorEl) this.wireBlock(anchorEl);
    root.querySelectorAll<HTMLElement>("[data-anchor]").forEach((b) => this.wireBlock(b));
  }

  /** Old-sequence diff state for one notes section. `null` requests remount when an
   *  existing list is not stampable; an absent list is the valid first-note case. */
  private notesDiff(
    sectionClass: "footnotes" | "endnotes",
    units: RenderUnit[],
  ): { lis: HTMLElement[]; diff: UnitDiff } | null {
    const ol = this.editRoot.querySelector<HTMLElement>(`section.${sectionClass} > ol`);
    const lis = ol ? (Array.from(ol.children).filter((c) => c.tagName === "LI") as HTMLElement[]) : [];
    if (units.length === 0 && lis.length === 0) return { lis, diff: diffUnits([], []) };
    // Full render omits empty note sections. Treat the first note as an insertion into
    // an empty list; applyNotesDiff creates the converter-equivalent section + list.
    if (!ol) return { lis, diff: diffUnits([], units) };
    const tokens: string[] = [];
    for (const li of lis) {
      const unid = li.getAttribute("data-note-anchor");
      if (!unid) return null; // unstamped DOM (older mount) — remount restamps
      const sig = li.getAttribute("data-render-sig");
      tokens.push(sig ? `${unid}|${sig}` : unid);
    }
    return { lis, diff: diffUnits(tokens, units) };
  }

  /** Apply a notes-section diff: rebuild the `<ol>`'s li list, preserving kept nodes. */
  private applyNotesDiff(
    sectionClass: "footnotes" | "endnotes",
    units: RenderUnit[],
    state: { lis: HTMLElement[]; diff: UnitDiff },
    rendered: Record<string, string | null>,
  ): void {
    if (state.diff.added.length === 0 && state.diff.removed.length === 0) return;
    // Removing the LAST note removes the whole section — a full render emits no
    // section for a document without notes, and equivalence with remount is the pin.
    if (units.length === 0) {
      this.editRoot.querySelector(`section.${sectionClass}`)?.remove();
      return;
    }
    let section = this.editRoot.querySelector<HTMLElement>(`:scope > section.${sectionClass}`);
    let ol = section?.querySelector<HTMLOListElement>(`:scope > ol`) ?? null;
    if (!ol) {
      if (!section) {
        section = document.createElement("section");
        section.className = sectionClass;
        const following = sectionClass === "footnotes"
          ? this.editRoot.querySelector<HTMLElement>(`:scope > section.endnotes`)
          : null;
        if (following) following.before(section);
        else this.editRoot.append(section);
      }
      if (!section.querySelector(":scope > hr")) section.prepend(document.createElement("hr"));
      ol = document.createElement("ol");
      section.append(ol);
    }
    const prefix = sectionClass === "footnotes" ? "fn" : "en";
    const nodes: HTMLElement[] = [];
    for (let j = 0; j < units.length; j++) {
      const oi = state.diff.keep.get(j);
      if (oi !== undefined) {
        nodes.push(state.lis[oi]);
        continue;
      }
      nodes.push(this.buildNoteLi(prefix, units[j], rendered[units[j].id]!));
    }
    ol.replaceChildren(...nodes);
  }

  /** Build a notes-section `<li>` for a freshly rendered note — replicating the
   *  converter's chrome (id/value are re-stamped by the renumber pass; the backref
   *  goes inside the last paragraph, matching RenderFootnoteItem). */
  private buildNoteLi(prefix: "fn" | "en", unit: RenderUnit, html: string): HTMLElement {
    const li = document.createElement("li");
    li.setAttribute("data-note-anchor", unidOf(unit.id));
    if (unit.sig) li.setAttribute("data-render-sig", unit.sig);
    li.innerHTML = html;
    const paras = li.querySelectorAll<HTMLElement>(":scope > p");
    const last = paras[paras.length - 1];
    if (last) {
      const backref = document.createElement("a");
      backref.setAttribute("class", `${prefix}-backref`);
      backref.setAttribute("contenteditable", "false");
      // The separator space must live INSIDE the chrome element. As a bare text
      // node in the paragraph it was counted by blockContentText, giving the
      // committed baseline a phantom trailing character the session's flat text
      // does not have — every later span commit then computed offsets one past
      // the session text and died with offset_out_of_range, silently reverting
      // the user's note edit (#583).
      backref.textContent = " ↩";
      last.append(backref);
    }
    li.querySelectorAll<HTMLElement>("[data-anchor]").forEach((b) => this.wireBlock(b));
    return li;
  }

  /**
   * Rewrite position-derived note chrome from the session's citation-ordered note
   * list: the k-th marker in document order IS note k (ids ascend in reference
   * order), so marker sup text, hrefs/ids, li ids/values and backref hrefs are all
   * re-derived positionally. Pure attribute/text patching of generated chrome.
   */
  private renumberNoteChrome(kind: "footnote" | "endnote"): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.ListNotes !== "function") return;
    const prefix = kind === "footnote" ? "fn" : "en";
    let notes: Array<{ id: string; defAnchorId: string; ordinal: number }>;
    try {
      notes = JSON.parse(bridge.ListNotes(this.handle, kind === "endnote"));
    } catch {
      return;
    }
    if (!Array.isArray(notes)) return;
    const markers = Array.from(
      this.editRoot.querySelectorAll<HTMLElement>(`a.${kind}-ref`),
    ).filter((a) => !a.closest("section.footnotes, section.endnotes"));
    markers.forEach((a, k) => {
      const n = notes[k];
      if (!n) return;
      a.setAttribute("href", `#${prefix}-${n.id}`);
      a.id = `${prefix}-ref-${n.id}`;
      if (kind === "footnote") a.setAttribute("data-footnote-id", n.id);
      const sup = a.querySelector("sup");
      if (sup) sup.textContent = String(n.ordinal);
    });
    // Match list items by their stamped note anchor, NOT position: the section can
    // hold rendered-but-never-cited notes (Word's continuationNotice) that ListNotes
    // — a citation walk — does not list; positional pairing would relabel them.
    const byUnid = new Map(notes.map((n) => [unidOf(n.defAnchorId), n]));
    const lis = this.editRoot.querySelectorAll<HTMLElement>(`section.${kind}s > ol > li`);
    lis.forEach((li) => {
      const unid = li.getAttribute("data-note-anchor");
      const n = unid ? byUnid.get(unid) : undefined;
      if (!n) return;
      li.id = `${prefix}-${n.id}`;
      li.setAttribute("value", String(n.ordinal));
      li.querySelectorAll<HTMLElement>(`a.${prefix}-backref`).forEach((b) =>
        b.setAttribute("href", `#${prefix}-ref-${n.id}`),
      );
    });
  }

  /** After an incremental block swap, stale marker chrome in the swapped node (the
   *  throwaway render numbers citations from 1) is repaired in place. */
  private maybeRenumberNotes(fresh: HTMLElement): void {
    if (fresh.querySelector("a.footnote-ref")) this.renumberNoteChrome("footnote");
    if (fresh.querySelector("a.endnote-ref")) this.renumberNoteChrome("endnote");
  }

  /** Stamp the DOM state the reconciler diffs against: container signatures on body
   *  tables and `data-note-anchor` + signature on notes-section items. Called after
   *  every full mount; reconcile stamps its own insertions. `planJson`: a render plan the
   *  caller has already read at the session's current version (the windowed mount's). */
  private stampPlanState(planJson?: string): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.ListBlocks !== "function") return;
    try {
      const plan = JSON.parse(planJson ?? this.renderPlanJson()) as RenderPlan & { error?: string };
      if (plan.error) { this.planNotesKey = null; return; }
      this.planNotesKey = DocxEditor.notesKeyOf(plan);
      // Positional pairing: a fresh full mount renders exactly the plan's units in
      // order (verified invariant). On any mismatch, leave unstamped — an unstamped
      // unit just diffs as changed and re-renders once.
      const nodes = this.bodyUnitNodes();
      if (nodes.length === plan.body.length) {
        nodes.forEach((el, k) => {
          if (plan.body[k].sig && el.getAttribute("data-anchor") === unidOf(plan.body[k].id))
            el.setAttribute("data-render-sig", plan.body[k].sig!);
        });
      }
      const stampNotes = (sectionClass: string, units: RenderUnit[]): void => {
        const lis = this.editRoot.querySelectorAll<HTMLElement>(`section.${sectionClass} > ol > li`);
        if (lis.length !== units.length) return; // inconsistent — leave unstamped (reconcile will remount)
        lis.forEach((li, k) => {
          li.setAttribute("data-note-anchor", unidOf(units[k].id));
          if (units[k].sig) li.setAttribute("data-render-sig", units[k].sig!);
        });
      };
      stampNotes("footnotes", plan.footnotes);
      stampNotes("endnotes", plan.endnotes);
    } catch {
      /* stamping is best-effort; unstamped DOM just falls back to remount */
    }
  }

  /**
   * Full re-render from current session state (after undo/redo, and after list edits where
   * single-block rendering can't compute numbering). Optionally focus the editable block at
   * `focusIndex` (caret at start, or end if `caretAtEnd`) — addressed by index because a
   * block's content-hashed unid changes across the save/reproject a remount performs.
   */
  private remount(focusIndex = -1, caretAtEnd = false): void {
    this.container.classList.toggle("docx-paginated", !!this.options.paginated);
    this.teardownBlockDrag();
    this.refreshAnchorMap();
    const fullHtml = this.renderFullHtml();
    this.activeBlock = null;
    if (this.options.paginated) this.mountPaginated(fullHtml);
    else this.mountHtml(fullHtml);
    if (focusIndex >= 0) {
      const blocks = this.editableList();
      const target = blocks[Math.min(focusIndex, blocks.length - 1)];
      if (target) {
        this.activeBlock = target;
        placeCaretAtOffset(target, caretAtEnd ? (target.textContent ?? "").length : 0);
      }
    }
    // A remount rebuilds the body from the live session; re-resolve the section so undo/redo of a
    // section-affecting edit (or a pagination toggle) leaves the bands describing the right one.
    this.syncRegionToBody(this.activeBlock ?? undefined);
    this.setupBlockDrag();
    this.gutter?.schedule();
    if (this.selectedPictureId && !this.pictureOverlay?.picture?.isConnected) this.deselectPicture();
  }

  /** Re-append the gutter after a mount emptied the container (mounts replace `innerHTML`). */
  private readoptGutter(): void {
    this.gutter?.reattach();
  }
}

const commentGutterStyledDocuments = new WeakSet<Document>();
function ensureCommentGutterStyles(doc: Document): void {
  if (commentGutterStyledDocuments.has(doc)) return;
  commentGutterStyledDocuments.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusCommentGutter = "true";
  style.textContent = COMMENT_GUTTER_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}
