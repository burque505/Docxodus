/**
 * Word's Preview Results: each merge field in the editor shows a recipient's value in place of
 * «Name», as the merge will fill it, without a change to the document (the value is drawn by CSS
 * over the field's own text, which stays as it is). Next Record and Next Record If move to the
 * next row, as in a sheet of labels; Merge Record # shows the row's number. Fields re-rendered by
 * an edit are previewed again.
 */
import { mergeFieldSample, parseMergeFieldCode } from "./ribbon-mergefields-pane.js";

const ATTR = "data-merge-preview";

export const MERGE_PREVIEW_CSS = `
[${ATTR}], [${ATTR}] * { font-size: 0 !important; }
[${ATTR}]::after { content: attr(${ATTR}); font-size: var(--dxr-merge-preview-size, 11pt) !important; }
`;

export class MergePreview {
  private observer: MutationObserver | null = null;
  private frame = 0;
  private columns = new Map<string, number>();
  private rows: string[][] = [];
  private index = 0;

  constructor(private readonly root: HTMLElement) {}

  get active(): boolean {
    return this.observer !== null;
  }

  /** Show row `index` (0-based) of `rows`, whose columns are `columns`. */
  show(columns: string[], rows: string[][], index: number): void {
    this.columns = new Map(columns.map((c, i) => [c.toLowerCase(), i]));
    this.rows = rows;
    this.index = index;
    if (!this.observer) {
      this.observer = new MutationObserver(() => this.schedule());
      this.observer.observe(this.root, { childList: true, subtree: true });
    }
    this.apply();
  }

  /** Back to the fields' own text. */
  hide(): void {
    this.observer?.disconnect();
    this.observer = null;
    cancelAnimationFrame(this.frame);
    for (const el of Array.from(this.root.querySelectorAll<HTMLElement>(`[${ATTR}]`))) {
      el.removeAttribute(ATTR);
      el.style.removeProperty("--dxr-merge-preview-size");
    }
  }

  private schedule(): void {
    cancelAnimationFrame(this.frame);
    this.frame = requestAnimationFrame(() => this.apply());
  }

  private value(name: string, offset: number): string {
    const row = this.rows[this.index + offset];
    const col = this.columns.get(name.trim().toLowerCase());
    return row && col !== undefined ? row[col] ?? "" : "";
  }

  private apply(): void {
    let offset = 0;
    // (the fields in document order; a field inside another one is its outer field's business)
    const fields = Array.from(this.root.querySelectorAll<HTMLElement>("[data-field-code]"))
      .filter((f) => !f.parentElement?.closest("[data-field-code]") && !f.closest("[data-hf-band]"));
    for (const f of fields) {
      const code = (f.dataset.fieldCode ?? "").trim();
      const word = /^\S+/.exec(code)?.[0].toUpperCase() ?? "";
      let shown: string | null = null;
      switch (word) {
        case "NEXT": case "NEXTIF": offset++; shown = ""; break;
        case "SKIPIF": shown = ""; break;
        case "MERGEREC": shown = String(this.index + offset + 1); break;
        case "MERGESEQ": shown = "1"; break;
        default: {
          const choice = parseMergeFieldCode(code);
          if (word === "MERGEFIELD" || (choice.format !== "other" && choice.name)) {
            const v = this.value(choice.name, offset);
            shown = v.trim() ? mergeFieldSample(choice, v) : choice.format === "iftf" ? mergeFieldSample(choice, "false") : "";
          }
        }
      }
      if (shown === null) continue;
      if (!f.hasAttribute(ATTR)) {
        const size = getComputedStyle(f).fontSize;
        if (size && size !== "0px") f.style.setProperty("--dxr-merge-preview-size", size);
      }
      if (f.getAttribute(ATTR) !== shown) f.setAttribute(ATTR, shown);
    }
  }
}
