/**
 * The ribbon's Cross-reference dialog — Word's Insert > Cross-reference: a reference type
 * (heading, bookmark, numbered item), what the reference shows, "Insert as hyperlink" and
 * "Include above/below", and the list of targets. Insert leaves the dialog open for the next one,
 * as Word's does; Cancel becomes Close after the first.
 */
import { makeDialogMovable, restoreDialogPlace } from "./ribbon-paragraph-dialog.js";
import {
  CROSS_REFERENCE_INSERTS,
  crossReferenceInsertLabel,
  crossReferenceTakesPosition,
  type CrossReferenceInsert,
  type CrossReferenceType,
} from "./cross-reference.js";

export const CROSSREF_DIALOG_CSS = `
.dxr-xrdialog { width: min(440px, calc(100vw - 32px)); }
.dxr-xrdialog .dxr-xrbody { display: grid; grid-template-columns: 1fr 1fr; gap: 8px 12px; padding: 10px 16px 4px; }
.dxr-xrdialog .dxr-xrfield { display: flex; flex-direction: column; gap: 4px; min-width: 0; }
.dxr-xrdialog .dxr-xrfield select { width: 100%; box-sizing: border-box; }
.dxr-xrdialog .dxr-xrwide { grid-column: 1 / -1; }
.dxr-xrdialog select[data-xr="list"] { min-height: 12em; }
.dxr-xrdialog .dxr-xrempty { color: var(--dxr-ink-soft, #667085); font-style: italic; }
`;

const LABELS = [
  "Heading text", "Bookmark text", "Paragraph text", "Page number", "Heading number",
  "Heading number (no context)", "Heading number (full context)", "Paragraph number",
  "Paragraph number (no context)", "Paragraph number (full context)", "Above/below",
];

export const CROSSREF_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-xrdialog" data-dxr="crossrefdialog">
  <form method="dialog" novalidate>
    <h2 data-xr-title>Cross-reference</h2>
    <div class="dxr-xrbody">
      <label class="dxr-xrfield">Reference type <select data-xr="type">
        <option value="heading">Heading</option>
        <option value="bookmark">Bookmark</option>
        <option value="numbered">Numbered item</option>
      </select></label>
      <label class="dxr-xrfield">Insert reference to <select data-xr="insert"></select></label>
      <label class="dxr-dcheck"><input type="checkbox" data-xr="hyperlink" checked /> Insert as hyperlink</label>
      <label class="dxr-dcheck"><input type="checkbox" data-xr="position" /> Include above/below</label>
      <label class="dxr-xrfield dxr-xrwide"><span data-xr-for="heading">For which heading:</span><span data-xr-for="bookmark" hidden>For which bookmark:</span><span data-xr-for="numbered" hidden>For which numbered item:</span>
        <select size="12" data-xr="list"></select></label>
    </div>
    <div class="dxr-derror" role="alert"><span data-xr-msg="refused" hidden>The cross-reference cannot be put there.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="insert" data-xr="insertbtn">Insert</button>
      <button type="submit" value="cancel" formnovalidate data-xr="close"><span data-xr-close="cancel">Cancel</span><span data-xr-close="close" hidden>Close</span></button>
    </div>
    <span hidden data-xr-text="none">(none)</span>
    ${LABELS.map((label) => `<span hidden data-xr-label="${label}">${label}</span>`).join("")}
  </form>
</dialog>
`;

/** One entry of the target list: its key (an anchor id, or a bookmark's name) and what it shows. */
export interface CrossReferenceTarget {
  key: string;
  label: string;
  /** Heading level (1-9), shown as indentation. */
  level?: number;
}

export interface CrossReferenceDialogHost {
  targets(type: CrossReferenceType): CrossReferenceTarget[];
  /** Insert the reference where the dialog was opened (then after the last one inserted). */
  insert(type: CrossReferenceType, key: string, insert: CrossReferenceInsert, hyperlink: boolean, position: boolean): { ok: boolean; code?: string };
  /** The dialog closed: put the caret back in the document. */
  closed(): void;
}

/** The Cross-reference dialog over its markup (`CROSSREF_DIALOG_HTML`, already in the page). */
export class CrossReferenceDialog {
  /** What the last insert showed, per reference type: the dialog opens on it again, as Word's. */
  private lastInsert: Partial<Record<CrossReferenceType, CrossReferenceInsert>> = {};

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: CrossReferenceDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-xr-title]")!);
    const form = dialog.querySelector("form")!;
    this.select("type").addEventListener("change", () => { this.fillInsert(); this.fillList(); this.showError(null); });
    this.select("insert").addEventListener("change", () => {
      this.lastInsert[this.type] = this.insertChoice;
      this.enable();
    });
    this.list.addEventListener("change", () => { this.showError(null); this.enable(); });
    this.list.addEventListener("dblclick", () => { if (this.list.value) this.insert(); });
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter && submitter.value !== "insert") { this.close(); return; }
      this.insert();
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.close(); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  show(): void {
    this.fillInsert();
    this.fillList();
    this.closeLabel(false);
    this.showError(null);
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.select("type").focus();
  }

  private close(): void {
    this.dialog.close("close");
    this.host.closed();
  }

  private insert(): void {
    const key = this.list.value;
    if (!key) return;
    const result = this.host.insert(this.type, key, this.insertChoice,
      this.input("hyperlink").checked, this.input("position").checked && !this.input("position").disabled);
    if (!result.ok) { this.showError("refused"); return; }
    this.showError(null);
    this.closeLabel(true);
    // (a reference to a heading or numbered item may have given it a new anchor: list them again)
    const keep = this.list.selectedIndex;
    this.fillList();
    if (keep >= 0 && keep < this.list.options.length && !this.list.options[keep].disabled) this.list.selectedIndex = keep;
    this.enable();
    this.list.focus();
  }

  private get type(): CrossReferenceType {
    return this.select("type").value as CrossReferenceType;
  }

  private get insertChoice(): CrossReferenceInsert {
    return this.select("insert").value as CrossReferenceInsert;
  }

  private fillInsert(): void {
    const type = this.type;
    const select = this.select("insert");
    select.replaceChildren();
    for (const insert of CROSS_REFERENCE_INSERTS[type]) {
      const option = select.ownerDocument.createElement("option");
      option.value = insert;
      const label = crossReferenceInsertLabel(type, insert);
      // (from the markup, so a translated ribbon has it in its language)
      option.textContent = this.dialog.querySelector<HTMLElement>(`[data-xr-label="${label}"]`)?.textContent ?? label;
      select.appendChild(option);
    }
    select.value = this.lastInsert[type] ?? CROSS_REFERENCE_INSERTS[type][0];
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-xr-for]")))
      span.hidden = span.dataset.xrFor !== type;
    this.enable();
  }

  private fillList(): void {
    const list = this.list;
    const keep = list.value;
    const targets = this.host.targets(this.type);
    list.replaceChildren();
    for (const target of targets) {
      const option = list.ownerDocument.createElement("option");
      option.value = target.key;
      option.textContent = "   ".repeat(Math.max(0, (target.level ?? 1) - 1)) + target.label;
      option.title = target.label;
      list.appendChild(option);
    }
    if (targets.length === 0) {
      const empty = list.ownerDocument.createElement("option");
      empty.disabled = true;
      empty.value = "";
      empty.className = "dxr-xrempty";
      empty.textContent = this.dialog.querySelector<HTMLElement>('[data-xr-text="none"]')?.textContent ?? "(none)";
      list.appendChild(empty);
    }
    list.value = targets.some((t) => t.key === keep) ? keep : "";
    if (!list.value && targets.length > 0) list.selectedIndex = 0;
    this.enable();
  }

  private enable(): void {
    this.button("insertbtn").disabled = !this.list.value;
    this.input("position").disabled = !crossReferenceTakesPosition(this.insertChoice);
  }

  private closeLabel(inserted: boolean): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-xr-close]")))
      span.hidden = (span.dataset.xrClose === "close") !== inserted;
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-xr-msg]")))
      span.hidden = span.dataset.xrMsg !== message;
  }

  private get list(): HTMLSelectElement {
    return this.select("list");
  }

  private select(name: string): HTMLSelectElement {
    return this.dialog.querySelector<HTMLSelectElement>(`[data-xr="${name}"]`)!;
  }

  private button(name: string): HTMLButtonElement {
    return this.dialog.querySelector<HTMLButtonElement>(`[data-xr="${name}"]`)!;
  }

  private input(name: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`[data-xr="${name}"]`)!;
  }
}
