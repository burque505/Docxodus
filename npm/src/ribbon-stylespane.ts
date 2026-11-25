/**
 * The ribbon's Styles pane — Word's Home > Styles launcher: the document's paragraph and character
 * styles in a list beside it. A click applies a style; the caret paragraph's style is marked; the
 * filter shows the recommended styles, the ones in use or all; Show Preview draws each name in its
 * style's font. A row's menu (right-click, the context-menu key, or its ▾ button) is Word's: Update
 * to Match Selection, Modify…, Delete (custom styles), Add to / Remove from Style Gallery. New Style…
 * sits at the foot.
 */
import type { StyleInfo } from "./types.js";

export const STYLES_PANE_CSS = `
.dxr-stylespane { flex: 0 0 250px; display: flex; flex-direction: column; min-height: 0; overflow: hidden;
  background: var(--dxr-chrome); border-left: 1px solid var(--dxr-rule); font-size: 12.5px; }
.dxr-stylespane[hidden] { display: none; }
.dxr-stylespane .dxr-navhead select { font: inherit; font-weight: 400; font-size: 12px; max-width: 120px; }
.dxr-stylelist { flex: 1 1 auto; overflow: auto; padding: 4px 0 8px; }
.dxr-styleitem { display: flex; align-items: center; gap: 6px; padding: 3px 4px 3px 9px; cursor: pointer;
  color: var(--dxr-ink); border-left: 3px solid transparent; user-select: none; min-height: 22px; }
.dxr-styleitem:hover { background: var(--dxr-chrome-sunk); }
.dxr-styleitem:focus-visible { outline: 2px solid var(--dxr-accent); outline-offset: -2px; }
.dxr-styleitem[data-current] { background: var(--dxr-wash, #f0fdfa); border-left-color: var(--dxr-accent); }
.dxr-stylename { flex: 1 1 auto; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; line-height: 1.3; }
.dxr-styleglyph { flex: 0 0 16px; text-align: center; color: var(--dxr-muted); font-size: 12px; }
.dxr-stylemore { flex: 0 0 auto; border: 0; background: none; color: var(--dxr-muted); cursor: pointer; padding: 0 4px;
  visibility: hidden; font-size: 11px; }
.dxr-styleitem:hover .dxr-stylemore, .dxr-styleitem:focus-within .dxr-stylemore, .dxr-styleitem[data-current] .dxr-stylemore { visibility: visible; }
.dxr-stylefoot { display: flex; flex-direction: column; gap: 6px; padding: 8px 10px; border-top: 1px solid var(--dxr-rule); }
.dxr-stylefoot label { display: flex; align-items: center; gap: 6px; color: var(--dxr-ink); }
.dxr-stylefoot button { align-self: flex-start; font: inherit; font-size: 12px; padding: 3px 10px; cursor: pointer;
  border: 1px solid var(--dxr-rule-strong); border-radius: 4px; background: var(--dxr-chrome); color: var(--dxr-ink); }
.dxr-stylefoot button:hover { background: var(--dxr-chrome-sunk); }
.dxr[data-chrome="compact"] .dxr-stylespane { flex-basis: 200px; }
.dxr-navmenu button[hidden] { display: none; }
`;

/** The pane's markup: after the document in the ribbon's work area (see RIBBON_HTML). */
export const STYLES_PANE_HTML = `
<aside class="dxr-stylespane" data-dxr="stylespane" hidden aria-label="Styles">
  <div class="dxr-navhead"><span>Styles</span>
    <select data-sp="filter" title="Which styles the pane lists">
      <option value="recommended">Recommended</option>
      <option value="inuse">In use</option>
      <option value="all">All styles</option>
    </select>
    <button type="button" data-sp="close" title="Close the Styles pane" aria-label="Close">&times;</button></div>
  <div class="dxr-stylelist" data-sp="list" role="listbox" aria-label="Styles"></div>
  <p class="dxr-navempty" data-sp="empty" hidden>No styles to show.</p>
  <div class="dxr-stylefoot">
    <label title="Show each style's name in its own font"><input type="checkbox" data-sp="preview" /> Show Preview</label>
    <button type="button" data-sp="new" title="Create a new style from the selection's formatting">New Style&#8230;</button>
  </div>
  <span hidden data-sp-text="paragraph">Paragraph style</span>
  <span hidden data-sp-text="character">Character style</span>
  <span hidden data-sp-text="more">Style options</span>
</aside>
<div class="dxr-navmenu" data-dxr="stylespanemenu" role="menu">
  <button type="button" role="menuitem" data-spcmd="update"><span data-sp-label>Update to Match Selection</span></button>
  <button type="button" role="menuitem" data-spcmd="modify">Modify&#8230;</button>
  <hr />
  <button type="button" role="menuitem" data-spcmd="delete"><span data-sp-label>Delete&#8230;</span></button>
  <button type="button" role="menuitem" data-spcmd="addgallery">Add to Style Gallery</button>
  <button type="button" role="menuitem" data-spcmd="removegallery">Remove from Style Gallery</button>
</div>
`;

export type StylesPaneCommand = "update" | "modify" | "delete" | "addgallery" | "removegallery";
export type StylesPaneFilter = "recommended" | "inuse" | "all";

export interface StylesPaneHost {
  /** The document's styles (with the built-ins Word always offers). */
  styles(): StyleInfo[];
  /** Goes up whenever the styles change (redraws the list). */
  version(): number;
  /** The caret paragraph's style id (and the caret run's character style, when it has one). */
  current(): { paragraph: string | null; character: string | null };
  apply(style: StyleInfo): void;
  command(style: StyleInfo, command: StylesPaneCommand): void;
  newStyle(): void;
  /** The pane's close button. */
  close(): void;
  /** Show Preview or the filter changed, for a host to keep. */
  onSettingsChange?(filter: StylesPaneFilter, preview: boolean): void;
}

/** Which styles the filter shows, sorted as Word sorts them (priority, then name). */
export function stylesFor(styles: StyleInfo[], filter: StylesPaneFilter): StyleInfo[] {
  const shown = styles.filter((s) => (s.type === "paragraph" || s.type === "character")
    && (filter === "all" || (filter === "inuse" ? !!s.inUse : !s.semiHidden || !!s.inUse)));
  return shown.sort((a, b) => filter === "all"
    ? a.name.localeCompare(b.name)
    : (a.uiPriority ?? 99) - (b.uiPriority ?? 99) || a.name.localeCompare(b.name));
}

/** A style's name as Word shows it: built-in names lower-case in the file ("heading 1") are capitalised. */
export function styleDisplayName(style: StyleInfo): string {
  return style.isCustom ? style.name : style.name.replace(/^[a-z]/, (c) => c.toUpperCase());
}

/** The Styles pane over its markup (`STYLES_PANE_HTML`, already in the page). */
export class StylesPane {
  private readonly list: HTMLElement;
  private readonly empty: HTMLElement;
  private readonly filterBox: HTMLSelectElement;
  private readonly previewBox: HTMLInputElement;
  private styles: StyleInfo[] = [];
  private signature = "";
  private menuFor: StyleInfo | null = null;

  constructor(private readonly pane: HTMLElement, private readonly menu: HTMLElement, private readonly host: StylesPaneHost) {
    this.list = pane.querySelector<HTMLElement>('[data-sp="list"]')!;
    this.empty = pane.querySelector<HTMLElement>('[data-sp="empty"]')!;
    this.filterBox = pane.querySelector<HTMLSelectElement>('[data-sp="filter"]')!;
    this.previewBox = pane.querySelector<HTMLInputElement>('[data-sp="preview"]')!;
    pane.querySelector('[data-sp="close"]')!.addEventListener("click", () => this.host.close());
    pane.querySelector('[data-sp="new"]')!.addEventListener("click", () => this.host.newStyle());
    const settings = () => { this.refresh(true); this.host.onSettingsChange?.(this.filter, this.preview); };
    this.filterBox.addEventListener("change", settings);
    this.previewBox.addEventListener("change", settings);
    // A click on a row applies its style; mousedown keeps the document's selection.
    this.list.addEventListener("mousedown", (event) => {
      if ((event.target as HTMLElement).closest(".dxr-styleitem")) event.preventDefault();
    });
    this.list.addEventListener("click", (event) => {
      const row = (event.target as HTMLElement).closest<HTMLElement>(".dxr-styleitem");
      const style = row ? this.styleOf(row) : null;
      if (!row || !style) return;
      if ((event.target as HTMLElement).closest(".dxr-stylemore")) {
        const r = row.getBoundingClientRect();
        this.openMenu(style, r.right - 8, r.bottom);
        return;
      }
      this.host.apply(style);
    });
    this.list.addEventListener("keydown", (event) => this.onKey(event));
    this.list.addEventListener("contextmenu", (event) => {
      const row = (event.target as HTMLElement).closest<HTMLElement>(".dxr-styleitem");
      const style = row ? this.styleOf(row) : null;
      event.preventDefault();
      if (style) this.openMenu(style, event.clientX, event.clientY);
    });
    menu.addEventListener("mousedown", (event) => event.preventDefault());
    menu.addEventListener("click", (event) => {
      const button = (event.target as HTMLElement).closest<HTMLButtonElement>("button");
      if (!button || button.disabled) return;
      const style = this.menuFor;
      this.closeMenu();
      if (style && button.dataset.spcmd) this.host.command(style, button.dataset.spcmd as StylesPaneCommand);
    });
    menu.addEventListener("keydown", (event) => {
      const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button:not(:disabled)"));
      const at = items.indexOf(event.target as HTMLButtonElement);
      if (event.key === "ArrowDown") { event.preventDefault(); items[(at + 1) % items.length]?.focus(); }
      else if (event.key === "ArrowUp") { event.preventDefault(); items[(at - 1 + items.length) % items.length]?.focus(); }
    });
    const doc = pane.ownerDocument;
    doc.addEventListener("mousedown", (event) => {
      if (this.menu.hasAttribute("data-open") && !this.menu.contains(event.target as Node)) this.closeMenu();
    });
    doc.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && this.menu.hasAttribute("data-open")) { event.preventDefault(); this.closeMenu(true); }
    });
  }

  get visible(): boolean {
    return !this.pane.hidden;
  }

  get filter(): StylesPaneFilter {
    return this.filterBox.value as StylesPaneFilter;
  }

  get preview(): boolean {
    return this.previewBox.checked;
  }

  /** Set the filter and Show Preview (a host's kept settings). */
  setSettings(filter: StylesPaneFilter | undefined, preview: boolean | undefined): void {
    if (filter && ["recommended", "inuse", "all"].includes(filter)) this.filterBox.value = filter;
    if (preview !== undefined) this.previewBox.checked = preview;
    this.refresh(true);
  }

  setVisible(on: boolean): void {
    this.pane.hidden = !on;
    if (!on) this.closeMenu();
    else this.refresh(true);
  }

  /** Read the styles and the caret's style again; redraw only when either changed. */
  refresh(force = false): void {
    if (this.pane.hidden) return;
    const current = this.host.current();
    const signature = `${this.host.version()}\u0001${current.paragraph ?? ""}\u0001${current.character ?? ""}\u0001${this.filter}\u0001${this.preview}`;
    if (!force && signature === this.signature) return;
    this.signature = signature;
    this.styles = stylesFor(this.host.styles(), this.filter);
    this.render(current);
  }

  private render(current: { paragraph: string | null; character: string | null }): void {
    const doc = this.list.ownerDocument;
    const focusedId = (doc.activeElement as HTMLElement | null)?.closest?.<HTMLElement>(".dxr-styleitem")?.dataset.id;
    this.list.replaceChildren();
    this.empty.hidden = this.styles.length > 0;
    const text = (key: string, fallback: string) => this.pane.querySelector<HTMLElement>(`[data-sp-text="${key}"]`)?.textContent ?? fallback;
    const kinds = { paragraph: text("paragraph", "Paragraph style"), character: text("character", "Character style") };
    const more = text("more", "Style options");
    const marked = current.character ?? current.paragraph;
    for (const style of this.styles) {
      const row = doc.createElement("div");
      row.className = "dxr-styleitem";
      row.dataset.id = style.id;
      row.dataset.type = style.type;
      row.setAttribute("role", "option");
      row.tabIndex = -1;
      const name = doc.createElement("span");
      name.className = "dxr-stylename";
      name.textContent = styleDisplayName(style);
      if (this.preview) this.paintPreview(name, style);
      const glyph = doc.createElement("span");
      glyph.className = "dxr-styleglyph";
      glyph.textContent = style.type === "character" ? "a" : "¶";
      glyph.title = style.type === "character" ? kinds.character : kinds.paragraph;
      const button = doc.createElement("button");
      button.type = "button";
      button.className = "dxr-stylemore";
      button.tabIndex = -1;
      button.textContent = "▾";
      button.title = more;
      button.setAttribute("aria-label", more);
      row.title = `${styleDisplayName(style)} (${style.type === "character" ? kinds.character : kinds.paragraph})`;
      if (style.id === marked) { row.dataset.current = ""; row.setAttribute("aria-selected", "true"); }
      row.append(name, glyph, button);
      this.list.appendChild(row);
    }
    const rows = Array.from(this.list.children) as HTMLElement[];
    const tabStop = rows.find((r) => r.dataset.id === focusedId) ?? rows.find((r) => r.hasAttribute("data-current")) ?? rows[0];
    if (tabStop) tabStop.tabIndex = 0;
    if (focusedId && tabStop?.dataset.id === focusedId) tabStop.focus({ preventScroll: true });
    this.list.querySelector<HTMLElement>("[data-current]")?.scrollIntoView({ block: "nearest" });
  }

  /** The style's look on its name: font, size (kept between 8 and 20 pt), bold, italic, colour. */
  private paintPreview(el: HTMLElement, style: StyleInfo): void {
    const run = style.resolvedRun;
    if (!run) return;
    if (run.fontFamily) el.style.fontFamily = `"${run.fontFamily}", sans-serif`;
    if (run.fontSizePts) el.style.fontSize = `${Math.min(20, Math.max(8, run.fontSizePts))}pt`;
    if (run.bold) el.style.fontWeight = "700";
    if (run.italic) el.style.fontStyle = "italic";
    const lines = [run.underline ? "underline" : "", run.strike ? "line-through" : ""].filter(Boolean).join(" ");
    if (lines) el.style.textDecoration = lines;
    if (run.caps) el.style.textTransform = "uppercase";
    if (run.smallCaps) el.style.fontVariant = "small-caps";
    if (run.color && /^[0-9a-f]{6}$/i.test(run.color)) el.style.color = `#${run.color}`;
  }

  private styleOf(row: HTMLElement): StyleInfo | null {
    return this.styles.find((s) => s.id === row.dataset.id) ?? null;
  }

  /** Up/Down move between styles, Enter applies, the context-menu key (or Shift+F10) opens the menu. */
  private onKey(event: KeyboardEvent): void {
    const row = (event.target as HTMLElement).closest<HTMLElement>(".dxr-styleitem");
    const style = row ? this.styleOf(row) : null;
    if (!row || !style) return;
    const rows = Array.from(this.list.children) as HTMLElement[];
    const i = rows.indexOf(row);
    const focus = (r: HTMLElement | undefined) => {
      if (!r) return;
      for (const x of rows) x.tabIndex = -1;
      r.tabIndex = 0;
      r.focus();
    };
    const menuHere = () => { const r = row.getBoundingClientRect(); this.openMenu(style, r.left + 24, r.bottom); };
    switch (event.key) {
      case "ArrowDown": event.preventDefault(); focus(rows[i + 1]); return;
      case "ArrowUp": event.preventDefault(); focus(rows[i - 1]); return;
      case "Home": event.preventDefault(); focus(rows[0]); return;
      case "End": event.preventDefault(); focus(rows[rows.length - 1]); return;
      case "Enter": event.preventDefault(); this.host.apply(style); return;
      case "ContextMenu": event.preventDefault(); menuHere(); return;
      default:
        if (event.key === "F10" && event.shiftKey) { event.preventDefault(); menuHere(); }
    }
  }

  private openMenu(style: StyleInfo, x: number, y: number): void {
    this.menuFor = style;
    const name = styleDisplayName(style);
    const item = (cmd: StylesPaneCommand) => this.menu.querySelector<HTMLButtonElement>(`[data-spcmd="${cmd}"]`)!;
    const deletable = style.isCustom && !style.isDefault;
    item("delete").disabled = !deletable;
    item("addgallery").hidden = !!style.quickFormat;
    item("removegallery").hidden = !style.quickFormat;
    // (the labels name the style, as Word's do: "Update Heading 1 to Match Selection", "Delete Quote...")
    for (const cmd of ["update", "delete"] as const) {
      const label = item(cmd).querySelector<HTMLElement>("[data-sp-label]")!;
      label.dataset.template ??= label.textContent ?? "";
      label.textContent = this.labelFor(cmd, label.dataset.template, name);
    }
    this.menu.setAttribute("data-open", "");
    const view = this.menu.ownerDocument.defaultView;
    const w = this.menu.offsetWidth, h = this.menu.offsetHeight;
    this.menu.style.left = `${Math.max(4, Math.min(x, (view?.innerWidth ?? x + w) - w - 4))}px`;
    this.menu.style.top = `${Math.max(4, Math.min(y, (view?.innerHeight ?? y + h) - h - 4))}px`;
    this.menu.querySelector<HTMLButtonElement>("button:not(:disabled):not([hidden])")?.focus();
  }

  /** The label with the style's name in it: a translated template may hold "{0}"; an English one
   *  gets the name after its first word ("Update Heading 1 to Match Selection", "Delete Quote..."). */
  private labelFor(cmd: "update" | "delete", template: string, name: string): string {
    if (template.includes("{0}")) return template.replace("{0}", name);
    const space = template.indexOf(" ");
    if (cmd === "update" && space > 0) return `${template.slice(0, space)} ${name}${template.slice(space)}`;
    if (cmd === "delete") return `${template.replace(/…$|\.\.\.$/, "")} ${name}…`;
    return template;
  }

  private closeMenu(refocus = false): void {
    const style = this.menuFor;
    this.menu.removeAttribute("data-open");
    this.menuFor = null;
    if (refocus && style) this.list.querySelector<HTMLElement>(`.dxr-styleitem[data-id="${CSS.escape(style.id)}"]`)?.focus();
  }
}
