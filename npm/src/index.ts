/** Full browser API. Use docxodus/core for the engine without editor dependencies. */
export * from "./core.js";

export { DocxEditor } from "./editor.js";
export { EditorCommandError } from "./editor-commands.js";
export type { EditorCommandArgs, EditorCommandErrorCode, EditorCommandInfo } from "./editor-commands.js";
export type {
  DocxEditorOptions,
  DocxEditorExports,
  EditorAlignment,
  EditorMatch,
  EditorPageSetup,
  FormatKey,
} from "./editor.js";
// The comment gutter and header/footer region are the editor's own; exported for hosts that
// build their own chrome and want to drive them directly.
export { CommentGutter } from "./editor-comments.js";
export type { CommentGutterHost, CommentGutterOptions } from "./editor-comments.js";
export type { BandWhich } from "./editor-headerfooter.js";

// The ribbon is the editor's UI shell: tabbed chrome, anchor rail, table picker and
// loading overlay wired onto DocxEditor's command surface. `createRibbonEditor`
// (docxodus/embed) is the one-call version that also boots WASM.
export { mountRibbon } from "./ribbon.js";
// Change Case's Headline Case word lists, for a host that keeps them in a file to be edited.
export { DEFAULT_HEADLINE_CASE_TEXT, parseHeadlineCaseLists } from "./change-case.js";
export type { CaseMode, HeadlineCaseLists } from "./change-case.js";
export { compileFind, expandReplacement, specialFindCodes, specialReplaceCodes, findOptionApplies, normalizeFindOptions, FIND_OPTION_KEYS } from "./find-query.js";
export type { FindOptions, CompiledFind, SpecialCode } from "./find-query.js";
export type {
  RibbonEditor,
  RibbonOptions,
  RibbonRecentEntry,
  RibbonCommand,
  RibbonFileMenuItem,
  RibbonHistoryOptions,
  RibbonHistoryBinding,
  RibbonChromeMode,
  RibbonState,
  RibbonLoader,
  RibbonLoaderOptions,
  RibbonLoaderStage,
  RibbonLoaderFeature,
} from "./ribbon.js";
