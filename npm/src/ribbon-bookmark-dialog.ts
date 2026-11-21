/**
 * The ribbon's Bookmark dialog — Word's Insert > Bookmark: the document's bookmarks, sorted by name
 * or by location, Word's own hidden ones on request; Add (a name the document has moves that
 * bookmark to the selection), Delete and Go To.
 */
import { makeDialogMovable, restoreDialogPlace } from "./ribbon-paragraph-dialog.js";

export const BOOKMARK_DIALOG_CSS = `
.dxr-bmdialog { width: min(360px, calc(100vw - 32px)); }
.dxr-bmdialog .dxr-bmbody { display: grid; grid-template-columns: 1fr auto; gap: 8px 12px; padding: 10px 16px 4px; }
.dxr-bmdialog .dxr-bmname { grid-column: 1 / -1; display: flex; flex-direction: column; gap: 4px; }
.dxr-bmdialog .dxr-bmname input { width: 100%; box-sizing: border-box; }
.dxr-bmdialog select[data-bm="list"] { width: 100%; min-height: 10em; box-sizing: border-box; }
.dxr-bmdialog .dxr-bmbuttons { display: flex; flex-direction: column; gap: 6px; }
.dxr-bmdialog .dxr-bmbuttons button { min-width: 6em; }
.dxr-bmdialog .dxr-bmopts { grid-column: 1 / -1; display: flex; flex-wrap: wrap; gap: 4px 14px; align-items: center; }
.dxr-bmdialog .dxr-bmempty { color: var(--dxr-ink-soft, #667085); font-style: italic; }
`;

export const BOOKMARK_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-bmdialog" data-dxr="bookmarkdialog">
  <form method="dialog" novalidate>
    <h2 data-bm-title>Bookmark</h2>
    <div class="dxr-bmbody">
      <label class="dxr-bmname">Bookmark name <input type="text" autocomplete="off" spellcheck="false" maxlength="40" data-bm="name" /></label>
      <select size="10" data-bm="list" aria-label="Bookmarks"></select>
      <div class="dxr-bmbuttons">
        <button type="submit" value="add" data-bm="add">Add</button>
        <button type="button" data-bm="delete">Delete</button>
        <button type="button" data-bm="goto">Go To</button>
      </div>
      <div class="dxr-bmopts" role="radiogroup">
        <span>Sort by</span>
        <label><input type="radio" name="bmsort" value="name" data-bm="sort-name" checked /> Name</label>
        <label><input type="radio" name="bmsort" value="location" data-bm="sort-location" /> Location</label>
        <label class="dxr-dcheck"><input type="checkbox" data-bm="hidden" /> Hidden bookmarks</label>
      </div>
    </div>
    <div class="dxr-derror" role="alert"><span data-bm-msg="name" hidden>A bookmark name starts with a letter and has only letters, digits and _ (up to 40).</span><span data-bm-msg="inUse" hidden>A link or cross-reference uses this bookmark, so it is not deleted.</span><span data-bm-msg="refused" hidden>The bookmark cannot be put there.</span><span data-bm-msg="notShown" hidden>This bookmark is in a note or a comment, which Go To cannot show.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="cancel" formnovalidate>Close</button>
    </div>
    <span hidden data-bm-text="none">(no bookmarks)</span>
  </form>
</dialog>
`;

/** A bookmark name Word's dialog takes: a letter, then letters, digits and underscores, 40 at most. */
export function isBookmarkName(name: string): boolean {
  return /^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(name);
}

export interface BookmarkDialogHost {
  /** The document's bookmarks in document order (`hidden`: Word's own `_…` ones too). */
  list(hidden: boolean): Array<{ name: string; text: string }>;
  /** Add (or move) bookmark `name` to the selection the dialog was opened on. */
  add(name: string): { ok: boolean; code?: string };
  remove(name: string): { ok: boolean; code?: string };
  goTo(name: string): boolean;
}

/** The Bookmark dialog over its markup (`BOOKMARK_DIALOG_HTML`, already in the page). */
export class BookmarkDialog {
  /** The bookmark Go To last showed: the document's selection is put back on it when the dialog closes. */
  private wentTo: string | null = null;

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: BookmarkDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-bm-title]")!);
    const form = dialog.querySelector("form")!;
    const name = this.input("name");
    const list = this.list;
    name.addEventListener("input", () => {
      name.removeAttribute("aria-invalid");
      this.showError(null);
      const typed = name.value.trim();
      const match = Array.from(list.options).find((o) => o.value.toLowerCase() === typed.toLowerCase());
      list.value = match ? match.value : "";
      this.enable();
    });
    list.addEventListener("change", () => {
      if (list.value) name.value = list.value;
      this.showError(null);
      this.enable();
    });
    list.addEventListener("dblclick", () => { if (list.value) this.goTo(); });
    for (const box of Array.from(dialog.querySelectorAll<HTMLInputElement>('input[name="bmsort"], [data-bm="hidden"]')))
      box.addEventListener("change", () => this.fill());
    this.button("delete").addEventListener("click", () => this.remove());
    this.button("goto").addEventListener("click", () => this.goTo());
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter && submitter.value !== "add") { this.close(); return; }
      this.add();
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.close(); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  /** Open on the document's bookmarks, the Name box holding `suggested` (Word: the bookmark at the caret). */
  show(suggested = ""): void {
    this.wentTo = null;
    this.input("hidden").checked = false;
    this.fill();
    const name = this.input("name");
    name.value = suggested;
    this.list.value = Array.from(this.list.options).some((o) => o.value === suggested) ? suggested : "";
    name.removeAttribute("aria-invalid");
    this.showError(null);
    this.enable();
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    name.focus();
    name.select();
  }

  private close(): void {
    this.dialog.close("close");
    // (the dialog gave the focus back to where it came from: the selection Go To made goes with it)
    if (this.wentTo) this.host.goTo(this.wentTo);
    this.wentTo = null;
  }

  private add(): void {
    const name = this.input("name");
    const value = name.value.trim();
    if (!isBookmarkName(value)) { this.invalid(); return; }
    this.dialog.close("add");
    this.wentTo = null;
    const result = this.host.add(value);
    if (!result.ok) {
      // (back open, with the reason: the selection it was opened on still stands)
      this.dialog.showModal();
      restoreDialogPlace(this.dialog);
      if (result.code === "invalid_bookmark_name") this.invalid();
      else this.showError("refused");
    }
  }

  private remove(): void {
    const value = this.input("name").value.trim();
    const known = this.find(value);
    if (!known) return;
    const result = this.host.remove(known);
    if (!result.ok) { this.showError(result.code === "bookmark_in_use" ? "inUse" : "refused"); return; }
    if (this.wentTo === known) this.wentTo = null;
    this.fill();
    this.input("name").value = "";
    this.enable();
    this.input("name").focus();
  }

  private goTo(): void {
    const known = this.find(this.input("name").value.trim());
    if (!known) return;
    // (a modal dialog leaves the document inert, where no selection can be made: Go To closes it,
    // then selects; Word's stays open)
    this.dialog.close("goto");
    this.wentTo = null;
    if (this.host.goTo(known)) return;
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.showError("notShown");
  }

  /** The listed bookmark named `name` (as Word, case aside), or null. */
  private find(name: string): string | null {
    if (!name) return null;
    return Array.from(this.list.options).find((o) => o.value.toLowerCase() === name.toLowerCase())?.value ?? null;
  }

  private fill(): void {
    const list = this.list;
    const keep = list.value;
    const marks = this.host.list(this.input("hidden").checked);
    const byName = this.input("sort-name").checked;
    const ordered = byName ? [...marks].sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" })) : marks;
    list.replaceChildren();
    for (const mark of ordered) {
      const option = list.ownerDocument.createElement("option");
      option.value = mark.name;
      option.textContent = mark.name;
      option.title = mark.text;
      list.appendChild(option);
    }
    if (ordered.length === 0) {
      const empty = list.ownerDocument.createElement("option");
      empty.disabled = true;
      empty.value = "";
      empty.className = "dxr-bmempty";
      // (from the markup, so a translated ribbon has it in its language)
      empty.textContent = this.dialog.querySelector<HTMLElement>('[data-bm-text="none"]')?.textContent ?? "(no bookmarks)";
      list.appendChild(empty);
    }
    list.value = ordered.some((m) => m.name === keep) ? keep : "";
    this.enable();
  }

  private enable(): void {
    const value = this.input("name").value.trim();
    const known = !!this.find(value);
    this.button("add").disabled = value === "";
    this.button("delete").disabled = !known;
    this.button("goto").disabled = !known;
  }

  private invalid(): void {
    const name = this.input("name");
    name.setAttribute("aria-invalid", "true");
    this.showError("name");
    name.focus();
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-bm-msg]")))
      span.hidden = span.dataset.bmMsg !== message;
  }

  private get list(): HTMLSelectElement {
    return this.dialog.querySelector<HTMLSelectElement>('[data-bm="list"]')!;
  }

  private button(name: string): HTMLButtonElement {
    return this.dialog.querySelector<HTMLButtonElement>(`[data-bm="${name}"]`)!;
  }

  private input(name: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`[data-bm="${name}"]`)!;
  }
}
