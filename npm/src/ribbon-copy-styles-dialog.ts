/**
 * The ribbon's Copy Styles dialog — Word's Organizer (Manage Styles > Import/Export), from a file
 * into the document: Browse picks a document or template, its styles are listed beside this
 * document's, and Copy brings the chosen ones in. A style this document has by the same name is
 * replaced only after Word's question (Yes, Yes to All, No, Cancel). The dialog stays open after a
 * copy, for the next one, as the Organizer does.
 */
import type { StyleInfo } from "./types.js";
import { makeDialogMovable, restoreDialogPlace } from "./ribbon-paragraph-dialog.js";

export const COPY_STYLES_DIALOG_CSS = `
.dxr-csdialog { width: min(620px, calc(100vw - 32px)); }
.dxr-csdialog .dxr-csbody { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 14px; padding: 10px 16px 4px; }
.dxr-csdialog .dxr-csfile { grid-column: 1 / -1; display: flex; align-items: center; gap: 10px; min-width: 0; }
.dxr-csdialog .dxr-csfile [data-cs="filename"] { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--dxr-muted); }
.dxr-csdialog .dxr-cscol { display: flex; flex-direction: column; gap: 4px; min-width: 0; }
.dxr-csdialog select[multiple], .dxr-csdialog select[size] { width: 100%; height: 16em; box-sizing: border-box; }
.dxr-csdialog .dxr-csopts { grid-column: 1 / -1; display: flex; flex-wrap: wrap; align-items: center; gap: 6px 12px; }
.dxr-csdialog .dxr-csdesc { grid-column: 1 / -1; min-height: 2.6em; margin: 0; color: var(--dxr-muted); font-size: 11.5px; }
.dxr-csdialog .dxr-dfoot [data-cs="all"] { margin-right: auto; }
.dxr-csask { width: min(420px, calc(100vw - 32px)); }
.dxr-csask p { margin: 12px 16px; }
`;

export const COPY_STYLES_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-csdialog" data-dxr="copystylesdialog">
  <form method="dialog" novalidate>
    <h2 data-cs-title>Copy Styles from Document</h2>
    <div class="dxr-csbody">
      <div class="dxr-csfile">
        <label class="dxr-btn" tabindex="0" data-cs="browse">Browse&#8230;<input data-cs="file" type="file" accept=".docx,.dotx,.docm,.dotm" hidden /></label>
        <span data-cs="filename"></span>
      </div>
      <div class="dxr-cscol">
        <label for="dxr-cs-from">In the other file</label>
        <select id="dxr-cs-from" multiple data-cs="from"></select>
      </div>
      <div class="dxr-cscol">
        <label for="dxr-cs-to">In this document</label>
        <select id="dxr-cs-to" size="16" data-cs="to"></select>
      </div>
      <div class="dxr-csopts">
        <label for="dxr-cs-show">Show</label>
        <select id="dxr-cs-show" data-cs="show">
          <option value="">All styles</option>
          <option value="paragraph">Paragraph styles</option>
          <option value="character">Character styles</option>
          <option value="table">Table styles</option>
          <option value="numbering">List styles</option>
        </select>
        <span>&#10003; = this document has a style of that name</span>
      </div>
      <p class="dxr-csdesc" data-cs="desc"></p>
    </div>
    <div class="dxr-derror" role="alert"><span data-cs-msg="notWord" hidden>This file is not a Word document or template.</span><span data-cs-msg="refused" hidden></span><span data-cs-msg="none" hidden>Choose the styles to copy.</span></div>
    <div class="dxr-dfoot">
      <button type="button" data-cs="all">Select All</button>
      <button type="submit" value="copy" data-cs="copy">Copy</button>
      <button type="submit" value="cancel" formnovalidate>Close</button>
    </div>
    <span hidden data-cs-text="pick">Browse for a document or template to copy styles from.</span>
    <span hidden data-cs-text="empty">(no styles)</span>
    <span hidden data-cs-text="paragraph">Paragraph style</span>
    <span hidden data-cs-text="character">Character style</span>
    <span hidden data-cs-text="table">Table style</span>
    <span hidden data-cs-text="numbering">List style</span>
    <span hidden data-cs-text="basedOn">based on {0}</span>
    <span hidden data-cs-text="replaces">This document has a style of this name: copying replaces it.</span>
    <span hidden data-cs-text="copied">Styles copied: {0}</span>
  </form>
</dialog>
<dialog class="dxr-dialog dxr-csask" data-dxr="copystylesask">
  <form method="dialog">
    <h2>Copy Styles from Document</h2>
    <p data-cs-ask></p>
    <div class="dxr-dfoot">
      <button type="submit" value="yes">Yes</button>
      <button type="submit" value="all">Yes to All</button>
      <button type="submit" value="no">No</button>
      <button type="submit" value="cancel">Cancel</button>
    </div>
    <span hidden data-cs-text="ask">Do you want to replace the existing style {0}?</span>
  </form>
</dialog>
`;

export interface CopyStylesDialogHost {
  /** This document's styles. */
  styles(): StyleInfo[];
  /** The other file's styles (its bytes as base64), or null when it is not a Word package. */
  read(base64: string): StyleInfo[] | null;
  /** Copy the other file's styles `ids` into the document; the engine's message when it refuses,
   *  else null. */
  copy(base64: string, ids: string[]): string | null;
}

const key = (s: StyleInfo) => `${s.type}\u0000${s.name.toLowerCase()}`;

/** A file's bytes as base64. */
async function fileBase64(file: Blob): Promise<string> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

/** The Copy Styles dialog over its markup (`COPY_STYLES_DIALOG_HTML`, already in the page). */
export class CopyStylesDialog {
  private source: StyleInfo[] = [];
  private base64 = "";
  private readonly ask: HTMLDialogElement;

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: CopyStylesDialogHost) {
    this.ask = (dialog.getRootNode() as Document | ShadowRoot).querySelector<HTMLDialogElement>('[data-dxr="copystylesask"]')!;
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-cs-title]")!);
    const file = this.el<HTMLInputElement>("file");
    const browse = this.el<HTMLElement>("browse");
    browse.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") { event.preventDefault(); file.click(); }
    });
    file.addEventListener("change", () => {
      const picked = file.files?.[0];
      file.value = "";
      if (picked) void this.load(picked);
    });
    this.el<HTMLSelectElement>("show").addEventListener("change", () => this.fill());
    this.el<HTMLSelectElement>("from").addEventListener("change", () => { this.showError(null); this.describe(); });
    this.el<HTMLButtonElement>("all").addEventListener("click", () => {
      for (const o of Array.from(this.el<HTMLSelectElement>("from").options)) o.selected = !o.disabled;
      this.describe();
    });
    dialog.querySelector("form")!.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "copy") void this.copy();
      else dialog.close("cancel");
    });
  }

  get open(): boolean { return this.dialog.open; }

  show(): boolean {
    this.source = [];
    this.base64 = "";
    this.el<HTMLElement>("filename").textContent = "";
    this.showError(null);
    this.fill();
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.el<HTMLElement>("browse").focus();
    return true;
  }

  /** Read a picked file's styles (also for tests and scripts: a Blob). */
  async load(file: Blob, name?: string): Promise<boolean> {
    this.showError(null);
    const base64 = await fileBase64(file);
    const styles = this.host.read(base64);
    this.el<HTMLElement>("filename").textContent = name ?? (file as File).name ?? "";
    if (!styles) {
      this.source = [];
      this.base64 = "";
      this.fill();
      this.showError("notWord");
      return false;
    }
    this.source = styles;
    this.base64 = base64;
    this.fill();
    this.el<HTMLSelectElement>("from").focus();
    return true;
  }

  private el<T extends HTMLElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-cs="${name}"]`)!;
  }

  private text(name: string): string {
    return (this.dialog.querySelector<HTMLElement>(`[data-cs-text="${name}"]`)
      ?? this.ask.querySelector<HTMLElement>(`[data-cs-text="${name}"]`))?.textContent ?? name;
  }

  private showError(which: "notWord" | "refused" | "none" | null, message?: string): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-cs-msg]"))) {
      span.hidden = span.dataset.csMsg !== which;
      if (which === "refused" && span.dataset.csMsg === "refused") span.textContent = message ?? "";
    }
  }

  private fill(): void {
    const show = this.el<HTMLSelectElement>("show").value;
    const byName = (a: StyleInfo, b: StyleInfo) => a.name.localeCompare(b.name);
    const mine = this.host.styles();
    const here = new Set(mine.map(key));
    const option = (value: string, label: string) => Object.assign(document.createElement("option"), { value, textContent: label });
    const placeholder = (label: string) => Object.assign(option("", label), { disabled: true });

    const from = this.el<HTMLSelectElement>("from");
    const chosen = new Set(Array.from(from.selectedOptions).map((o) => o.value));
    const sourceShown = this.source.filter((s) => !show || s.type === show).sort(byName);
    from.replaceChildren(...(this.base64 === ""
      ? [placeholder(this.text("pick"))]
      : sourceShown.length === 0
        ? [placeholder(this.text("empty"))]
        : sourceShown.map((s) => Object.assign(option(s.id, here.has(key(s)) ? `${s.name}  ✓` : s.name),
          { selected: chosen.has(s.id) }))));

    const to = this.el<HTMLSelectElement>("to");
    const mineShown = mine.filter((s) => !show || s.type === show).sort(byName);
    to.replaceChildren(...(mineShown.length ? mineShown.map((s) => option(s.id, s.name)) : [placeholder(this.text("empty"))]));
    this.describe();
  }

  /** The chosen style's type and basis, and whether copying it replaces one of this document's. */
  private describe(): void {
    const from = this.el<HTMLSelectElement>("from");
    const desc = this.el<HTMLElement>("desc");
    const ids = Array.from(from.selectedOptions).map((o) => o.value);
    const style = ids.length === 1 ? this.source.find((s) => s.id === ids[0]) : undefined;
    if (!style) { desc.textContent = ""; return; }
    const parts = [this.text(style.type)];
    const basis = style.basedOn ? this.source.find((s) => s.id === style.basedOn)?.name ?? style.basedOn : null;
    if (basis) parts[0] += ", " + this.text("basedOn").replace("{0}", basis);
    parts[0] += ".";
    if (new Set(this.host.styles().map(key)).has(key(style))) parts.push(this.text("replaces"));
    desc.textContent = parts.join(" ");
  }

  /** Word's question for one style that is already in the document. */
  private question(name: string): Promise<"yes" | "all" | "no" | "cancel"> {
    return new Promise((resolve) => {
      this.ask.querySelector<HTMLElement>("[data-cs-ask]")!.textContent = this.text("ask").replace("{0}", name);
      const done = () => {
        this.ask.removeEventListener("close", done);
        const v = this.ask.returnValue;
        resolve(v === "yes" || v === "all" || v === "no" ? v : "cancel");
      };
      this.ask.addEventListener("close", done);
      this.ask.returnValue = "";
      this.ask.showModal();
      this.ask.querySelector<HTMLButtonElement>('button[value="yes"]')!.focus();
    });
  }

  private async copy(): Promise<void> {
    this.showError(null);
    const from = this.el<HTMLSelectElement>("from");
    const chosen = Array.from(from.selectedOptions).map((o) => this.source.find((s) => s.id === o.value)).filter((s): s is StyleInfo => !!s);
    if (chosen.length === 0) { this.showError("none"); return; }
    const here = new Set(this.host.styles().map(key));
    const ids: string[] = [];
    let all = false;
    for (const style of chosen) {
      if (here.has(key(style)) && !all) {
        const answer = await this.question(style.name);
        if (answer === "cancel") { from.focus(); return; }
        if (answer === "no") continue;
        if (answer === "all") all = true;
      }
      ids.push(style.id);
    }
    if (ids.length === 0) { from.focus(); return; }
    const refused = this.host.copy(this.base64, ids);
    if (refused !== null) { this.showError("refused", refused); return; }
    for (const o of Array.from(from.options)) o.selected = false;
    this.fill();
    this.el<HTMLElement>("desc").textContent = this.text("copied").replace("{0}", String(ids.length));
    from.focus();
  }
}
