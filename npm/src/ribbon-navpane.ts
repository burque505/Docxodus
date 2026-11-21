/**
 * The ribbon's Navigation Pane — Word's View > Navigation Pane, Headings: the document's headings
 * as a tree beside it. A click goes to a heading; the heading the caret is under is marked; a drag
 * moves a heading with everything under it; the right-click menu is Word's (Promote, Demote, New
 * Heading Before / After, New Subheading, Delete, Select Heading and Content, Expand All, Collapse
 * All, Show Heading Levels). Collapsing a heading here hides its subheadings in the pane only.
 */
import type { OutlineEntry } from "./types.js";

export const NAV_PANE_CSS = `
.dxr-work { display: flex; flex: 1 1 auto; min-height: 0; }
.dxr-workmain { display: flex; flex-direction: column; flex: 1 1 auto; min-width: 0; min-height: 0; }
.dxr-navpane { flex: 0 0 240px; display: flex; flex-direction: column; min-height: 0; overflow: hidden;
  background: var(--dxr-chrome); border-right: 1px solid var(--dxr-rule); font-size: 12.5px; }
.dxr-navpane[hidden] { display: none; }
.dxr-navhead { display: flex; align-items: center; justify-content: space-between; padding: 8px 10px 6px;
  font-weight: 600; color: var(--dxr-ink); border-bottom: 1px solid var(--dxr-rule); }
.dxr-navhead button { border: 0; background: none; cursor: pointer; color: var(--dxr-muted); font-size: 14px; padding: 0 4px; }
.dxr-navlist { flex: 1 1 auto; overflow: auto; padding: 4px 0 12px; }
.dxr-navempty { padding: 10px 12px; color: var(--dxr-muted); line-height: 1.4; }
.dxr-navitem { display: flex; align-items: flex-start; gap: 2px; padding: 3px 8px 3px 0; cursor: pointer;
  color: var(--dxr-ink); border-left: 3px solid transparent; line-height: 1.35; user-select: none; }
.dxr-navitem:hover { background: var(--dxr-chrome-sunk); }
.dxr-navitem:focus-visible { outline: 2px solid var(--dxr-accent); outline-offset: -2px; }
.dxr-navitem[data-current] { background: var(--dxr-wash, #f0fdfa); border-left-color: var(--dxr-accent); font-weight: 600; }
.dxr-navitem[data-dragging] { opacity: .45; }
.dxr-navitem[data-drop="before"] { box-shadow: inset 0 2px 0 var(--dxr-accent); }
.dxr-navitem[data-drop="end"] { box-shadow: inset 0 -2px 0 var(--dxr-accent); }
.dxr-navtoggle { flex: 0 0 14px; text-align: center; color: var(--dxr-muted); font-size: 10px; line-height: 17px; }
.dxr-navtext { flex: 1 1 auto; min-width: 0; overflow: hidden; text-overflow: ellipsis; display: -webkit-box;
  -webkit-line-clamp: 2; -webkit-box-orient: vertical; word-break: break-word; }
.dxr-navitem[data-level="1"] .dxr-navtext { font-weight: 600; }
.dxr-navtext[data-blank] { color: var(--dxr-muted); font-style: italic; }
.dxr-navmenu { position: fixed; z-index: 40; display: none; min-width: 210px; padding: 4px 0; background: var(--dxr-chrome);
  border: 1px solid var(--dxr-rule-strong); border-radius: 6px; box-shadow: 0 6px 20px rgba(15, 23, 42, .12); }
.dxr-navmenu[data-open] { display: block; }
.dxr-navmenu button { display: block; width: 100%; text-align: left; border: 0; background: none; padding: 5px 14px;
  font: inherit; font-size: 12.5px; color: var(--dxr-ink); cursor: pointer; }
.dxr-navmenu button:hover:not(:disabled), .dxr-navmenu button:focus-visible { background: var(--dxr-chrome-sunk); outline: none; }
.dxr-navmenu button:disabled { color: var(--dxr-muted); cursor: default; }
.dxr-navmenu hr { border: 0; border-top: 1px solid var(--dxr-rule); margin: 4px 0; }
.dxr-navmenu .dxr-navlevels { display: none; }
.dxr-navmenu .dxr-navlevels[data-open] { display: block; padding-left: 10px; }
.dxr[data-chrome="compact"] .dxr-navpane { flex-basis: 180px; }
`;

/** The pane's markup: it goes where the ribbon's work area starts (see RIBBON_HTML). */
export const NAV_PANE_HTML = `
<aside class="dxr-navpane" data-dxr="navpane" hidden aria-label="Navigation">
  <div class="dxr-navhead"><span>Navigation</span><button type="button" data-nav="close" title="Close the Navigation Pane" aria-label="Close">&times;</button></div>
  <div class="dxr-navlist" data-nav="list" role="tree"></div>
  <p class="dxr-navempty" data-nav="empty" hidden>The document has no headings. Give a paragraph a Heading style to see it here.</p>
  <p class="dxr-navempty" data-nav="hidden" hidden>No headings at the levels shown. Right-click here and choose Show Heading Levels to see more.</p>
  <span hidden data-nav-text="blank">(empty heading)</span>
</aside>
<div class="dxr-navmenu" data-dxr="navmenu" role="menu">
  <button type="button" role="menuitem" data-navcmd="promote">Promote</button>
  <button type="button" role="menuitem" data-navcmd="demote">Demote</button>
  <hr />
  <button type="button" role="menuitem" data-navcmd="before">New Heading Before</button>
  <button type="button" role="menuitem" data-navcmd="after">New Heading After</button>
  <button type="button" role="menuitem" data-navcmd="sub">New Subheading</button>
  <hr />
  <button type="button" role="menuitem" data-navcmd="delete">Delete</button>
  <button type="button" role="menuitem" data-navcmd="select">Select Heading and Content</button>
  <hr />
  <button type="button" role="menuitem" data-navcmd="expand">Expand All</button>
  <button type="button" role="menuitem" data-navcmd="collapse">Collapse All</button>
  <button type="button" role="menuitem" data-navcmd="levels" aria-haspopup="true">Show Heading Levels &#9656;</button>
  <div class="dxr-navlevels" data-nav="levels">
    <button type="button" role="menuitem" data-navlevel="1">Show Heading 1</button>
    <button type="button" role="menuitem" data-navlevel="2">Show Headings 1-2</button>
    <button type="button" role="menuitem" data-navlevel="3">Show Headings 1-3</button>
    <button type="button" role="menuitem" data-navlevel="4">Show Headings 1-4</button>
    <button type="button" role="menuitem" data-navlevel="5">Show Headings 1-5</button>
    <button type="button" role="menuitem" data-navlevel="6">Show Headings 1-6</button>
    <button type="button" role="menuitem" data-navlevel="7">Show Headings 1-7</button>
    <button type="button" role="menuitem" data-navlevel="8">Show Headings 1-8</button>
    <button type="button" role="menuitem" data-navlevel="9">Show All Headings</button>
  </div>
</div>
`;

export type NavPaneCommand = "promote" | "demote" | "before" | "after" | "sub" | "delete" | "select";

export interface NavPaneHost {
  outline(): OutlineEntry[];
  /** The heading the caret is under, or null. */
  current(): string | null;
  goTo(entry: OutlineEntry): void;
  /** Move `entry`'s part before the heading `before`, or to the end of the document (null). */
  move(entry: OutlineEntry, before: OutlineEntry | null): void;
  command(entry: OutlineEntry, command: NavPaneCommand): void;
  /** The pane's close button. */
  close(): void;
}

/** The Navigation Pane over its markup (`NAV_PANE_HTML`, already in the page). */
export class NavPane {
  private readonly list: HTMLElement;
  private readonly empty: HTMLElement;
  /** Headings whose subheadings the pane hides, by anchor (anchors outlive edits elsewhere). */
  private readonly collapsed = new Set<string>();
  private maxLevel = 9;
  private entries: OutlineEntry[] = [];
  private signature = "";
  private dragged: OutlineEntry | null = null;
  private menuFor: OutlineEntry | null = null;

  constructor(private readonly pane: HTMLElement, private readonly menu: HTMLElement, private readonly host: NavPaneHost) {
    this.list = pane.querySelector<HTMLElement>('[data-nav="list"]')!;
    this.empty = pane.querySelector<HTMLElement>('[data-nav="empty"]')!;
    pane.querySelector('[data-nav="close"]')!.addEventListener("click", () => this.host.close());
    this.list.addEventListener("click", (event) => {
      const row = (event.target as HTMLElement).closest<HTMLElement>(".dxr-navitem");
      const entry = row ? this.entryOf(row) : null;
      if (!row || !entry) return;
      if ((event.target as HTMLElement).closest(".dxr-navtoggle")?.textContent) { this.toggle(entry); return; }
      this.host.goTo(entry);
    });
    this.list.addEventListener("keydown", (event) => this.onKey(event));
    // Anywhere in the pane: on a heading, the menu acts on it; elsewhere (or with every heading
    // hidden by Show Heading Levels) its heading commands are off, Expand All and the levels stay.
    pane.addEventListener("contextmenu", (event) => {
      const row = (event.target as HTMLElement).closest<HTMLElement>(".dxr-navitem");
      event.preventDefault();
      this.openMenu(row ? this.entryOf(row) : null, event.clientX, event.clientY);
    });
    this.wireDrag();
    menu.addEventListener("click", (event) => {
      const button = (event.target as HTMLElement).closest<HTMLButtonElement>("button");
      if (!button || button.disabled) return;
      const entry = this.menuFor;
      if (button.dataset.navcmd === "levels") {
        menu.querySelector('[data-nav="levels"]')!.toggleAttribute("data-open");
        // (taller now: kept inside the window)
        this.placeMenu(parseFloat(menu.style.left), parseFloat(menu.style.top));
        return;
      }
      this.closeMenu();
      if (button.dataset.navlevel) { this.showLevels(Number(button.dataset.navlevel)); return; }
      const cmd = button.dataset.navcmd;
      if (cmd === "expand") { this.collapsed.clear(); this.render(true); return; }
      if (cmd === "collapse") { this.collapseAll(); return; }
      if (entry && cmd) this.host.command(entry, cmd as NavPaneCommand);
    });
    const doc = pane.ownerDocument;
    doc.addEventListener("mousedown", (event) => {
      if (this.menu.hasAttribute("data-open") && !this.menu.contains(event.target as Node)) this.closeMenu();
    });
    doc.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && this.menu.hasAttribute("data-open")) { event.preventDefault(); this.closeMenu(); }
    });
  }

  get visible(): boolean {
    return !this.pane.hidden;
  }

  setVisible(on: boolean): void {
    this.pane.hidden = !on;
    if (!on) this.closeMenu();
    else this.refresh(true);
  }

  /** Read the outline and the caret's heading again; redraw only when either changed. */
  refresh(force = false): void {
    if (this.pane.hidden) return;
    this.entries = this.host.outline();
    const current = this.host.current();
    const signature = this.entries.map((e) => `${e.anchorId}\u0001${e.level}\u0001${e.text}`).join("\u0002")
      + `\u0003${current ?? ""}\u0003${this.maxLevel}\u0003${[...this.collapsed].join(",")}`;
    if (!force && signature === this.signature) return;
    this.signature = signature;
    this.render(false, current);
  }

  /** Which headings show: none under a collapsed one, none deeper than Show Heading Levels. */
  private shown(): OutlineEntry[] {
    const out: OutlineEntry[] = [];
    let hideBelow = Infinity;
    for (const entry of this.entries) {
      if (entry.level > hideBelow) continue;
      hideBelow = Infinity;
      if (entry.level > this.maxLevel) continue;
      out.push(entry);
      if (this.collapsed.has(entry.anchorId)) hideBelow = entry.level;
    }
    return out;
  }

  private render(reread: boolean, current: string | null = this.host.current()): void {
    if (reread) { this.signature = ""; this.refresh(); return; }
    const doc = this.list.ownerDocument;
    const focusedId = (doc.activeElement as HTMLElement | null)?.closest?.<HTMLElement>(".dxr-navitem")?.dataset.id;
    this.list.replaceChildren();
    this.empty.hidden = this.entries.length > 0;
    const shown = this.shown();
    this.pane.querySelector<HTMLElement>('[data-nav="hidden"]')!.hidden = !(this.entries.length > 0 && shown.length === 0);
    // (a heading has children when the next one in the whole outline is deeper)
    const hasChildren = new Set<string>();
    for (let i = 0; i + 1 < this.entries.length; i++)
      if (this.entries[i + 1].level > this.entries[i].level && this.entries[i + 1].level <= this.maxLevel) hasChildren.add(this.entries[i].anchorId);
    for (const entry of shown) {
      const row = doc.createElement("div");
      row.className = "dxr-navitem";
      row.dataset.id = entry.anchorId;
      row.dataset.level = String(entry.level);
      row.setAttribute("role", "treeitem");
      row.setAttribute("aria-level", String(entry.level));
      row.tabIndex = -1;
      row.draggable = true;
      row.style.paddingLeft = `${4 + (entry.level - 1) * 14}px`;
      const toggle = doc.createElement("span");
      toggle.className = "dxr-navtoggle";
      if (hasChildren.has(entry.anchorId)) {
        const closed = this.collapsed.has(entry.anchorId);
        toggle.textContent = closed ? "▸" : "▾";
        row.setAttribute("aria-expanded", String(!closed));
      }
      const text = doc.createElement("span");
      text.className = "dxr-navtext";
      const label = entry.text.trim();
      if (label) text.textContent = label;
      else { text.textContent = this.blankLabel(); text.dataset.blank = ""; }
      row.title = label;
      if (entry.anchorId === current) row.dataset.current = "";
      row.append(toggle, text);
      this.list.appendChild(row);
    }
    const rows = Array.from(this.list.children) as HTMLElement[];
    const tabStop = rows.find((r) => r.dataset.id === focusedId) ?? rows.find((r) => r.hasAttribute("data-current")) ?? rows[0];
    if (tabStop) tabStop.tabIndex = 0;
    if (focusedId && tabStop?.dataset.id === focusedId) tabStop.focus({ preventScroll: true });
    this.list.querySelector<HTMLElement>("[data-current]")?.scrollIntoView({ block: "nearest" });
  }

  /** The text for an empty heading, from the markup (so a translated ribbon has it in its language). */
  private blankLabel(): string {
    return this.menu.ownerDocument.querySelector<HTMLElement>('[data-nav-text="blank"]')?.textContent ?? "(empty heading)";
  }

  private entryOf(row: HTMLElement): OutlineEntry | null {
    return this.entries.find((e) => e.anchorId === row.dataset.id) ?? null;
  }

  private toggle(entry: OutlineEntry): void {
    if (this.collapsed.has(entry.anchorId)) this.collapsed.delete(entry.anchorId);
    else this.collapsed.add(entry.anchorId);
    this.render(true);
  }

  private collapseAll(): void {
    this.collapsed.clear();
    for (let i = 0; i + 1 < this.entries.length; i++)
      if (this.entries[i + 1].level > this.entries[i].level) this.collapsed.add(this.entries[i].anchorId);
    this.render(true);
  }

  private showLevels(level: number): void {
    this.maxLevel = level;
    this.collapsed.clear();
    this.render(true);
  }

  /** Up/Down move between headings, Left/Right collapse and expand, Enter goes there, the
   *  context-menu key (or Shift+F10) opens the menu. */
  private onKey(event: KeyboardEvent): void {
    const row = (event.target as HTMLElement).closest<HTMLElement>(".dxr-navitem");
    const entry = row ? this.entryOf(row) : null;
    if (!row || !entry) return;
    const rows = Array.from(this.list.children) as HTMLElement[];
    const i = rows.indexOf(row);
    const focus = (r: HTMLElement | undefined) => {
      if (!r) return;
      for (const x of rows) x.tabIndex = -1;
      r.tabIndex = 0;
      r.focus();
    };
    switch (event.key) {
      case "ArrowDown": event.preventDefault(); focus(rows[i + 1]); return;
      case "ArrowUp": event.preventDefault(); focus(rows[i - 1]); return;
      case "Home": event.preventDefault(); focus(rows[0]); return;
      case "End": event.preventDefault(); focus(rows[rows.length - 1]); return;
      case "ArrowLeft": if (row.getAttribute("aria-expanded") === "true") { event.preventDefault(); this.toggle(entry); } return;
      case "ArrowRight": if (row.getAttribute("aria-expanded") === "false") { event.preventDefault(); this.toggle(entry); } return;
      case "Enter": event.preventDefault(); this.host.goTo(entry); return;
      case "ContextMenu": {
        event.preventDefault();
        const r = row.getBoundingClientRect();
        this.openMenu(entry, r.left + 24, r.bottom);
        return;
      }
      default:
        if (event.key === "F10" && event.shiftKey) {
          event.preventDefault();
          const r = row.getBoundingClientRect();
          this.openMenu(entry, r.left + 24, r.bottom);
        }
    }
  }

  private openMenu(entry: OutlineEntry | null, x: number, y: number): void {
    this.menuFor = entry;
    const set = (cmd: string, enabled: boolean) => {
      const b = this.menu.querySelector<HTMLButtonElement>(`[data-navcmd="${cmd}"]`);
      if (b) b.disabled = !enabled;
    };
    for (const cmd of ["before", "after", "delete", "select"]) set(cmd, !!entry);
    set("promote", !!entry && entry.level > 1);
    set("demote", !!entry && entry.level < 9);
    set("sub", !!entry && entry.level < 9);
    this.menu.querySelector('[data-nav="levels"]')!.removeAttribute("data-open");
    this.menu.setAttribute("data-open", "");
    this.placeMenu(x, y);
    this.menu.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus();
  }

  /** Put the menu at (x, y), moved up and left as far as it takes to stay inside the window. */
  private placeMenu(x: number, y: number): void {
    const view = this.menu.ownerDocument.defaultView;
    const w = this.menu.offsetWidth, h = this.menu.offsetHeight;
    this.menu.style.left = `${Math.max(4, Math.min(x, (view?.innerWidth ?? x + w) - w - 4))}px`;
    this.menu.style.top = `${Math.max(4, Math.min(y, (view?.innerHeight ?? y + h) - h - 4))}px`;
  }

  private closeMenu(): void {
    this.menu.removeAttribute("data-open");
    this.menuFor = null;
  }

  /**
   * Drag a heading onto another: dropped on a heading's upper half, it goes before that heading;
   * on the lower half, before the next heading the pane shows (after the last: to the end). Never
   * into its own part.
   */
  private wireDrag(): void {
    const clear = () => {
      for (const r of Array.from(this.list.querySelectorAll<HTMLElement>("[data-drop], [data-dragging]"))) {
        r.removeAttribute("data-drop");
        r.removeAttribute("data-dragging");
      }
    };
    /** Where a drop at `event` puts the dragged heading: before `before` (null: the end), shown on `row`. */
    const dropAt = (event: DragEvent): { row: HTMLElement; mark: "before" | "end"; before: OutlineEntry | null } | null => {
      const dragged = this.dragged;
      const row = (event.target as HTMLElement).closest<HTMLElement>(".dxr-navitem");
      if (!dragged || !row) return null;
      const rows = Array.from(this.list.children) as HTMLElement[];
      const box = row.getBoundingClientRect();
      const lower = event.clientY > box.top + box.height / 2;
      const index = rows.indexOf(row);
      const beforeRow = lower ? rows[index + 1] : row;
      const before = beforeRow ? this.entryOf(beforeRow) : null;
      // (not into its own part: from the dragged heading to the next one at its level or above)
      const all = this.entries;
      const from = all.findIndex((e) => e.anchorId === dragged.anchorId);
      let to = all.length;
      for (let k = from + 1; k < all.length; k++) if (all[k].level <= dragged.level) { to = k; break; }
      const at = before ? all.findIndex((e) => e.anchorId === before.anchorId) : all.length;
      if (at > from && at <= to) return null;
      return beforeRow ? { row: beforeRow, mark: "before", before } : { row, mark: "end", before: null };
    };
    this.list.addEventListener("dragstart", (event) => {
      const row = (event.target as HTMLElement).closest<HTMLElement>(".dxr-navitem");
      const entry = row ? this.entryOf(row) : null;
      if (!row || !entry || !event.dataTransfer) return;
      this.dragged = entry;
      row.dataset.dragging = "";
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", entry.text);
    });
    this.list.addEventListener("dragover", (event) => {
      if (!this.dragged) return;
      const at = dropAt(event);
      for (const r of Array.from(this.list.querySelectorAll<HTMLElement>("[data-drop]"))) r.removeAttribute("data-drop");
      if (!at) { if (event.dataTransfer) event.dataTransfer.dropEffect = "none"; return; }
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
      at.row.dataset.drop = at.mark;
    });
    this.list.addEventListener("dragleave", (event) => {
      if (!this.list.contains(event.relatedTarget as Node)) for (const r of Array.from(this.list.querySelectorAll<HTMLElement>("[data-drop]"))) r.removeAttribute("data-drop");
    });
    this.list.addEventListener("drop", (event) => {
      const dragged = this.dragged;
      const at = dropAt(event);
      clear();
      this.dragged = null;
      if (!dragged || !at) return;
      event.preventDefault();
      this.host.move(dragged, at.before);
    });
    this.list.addEventListener("dragend", () => { clear(); this.dragged = null; });
  }
}
