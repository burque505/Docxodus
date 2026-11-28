/**
 * Cell selection, as Word's: a rectangle of a table's cells, chosen by dragging from one cell into
 * another, Shift+click or Shift+arrows past a cell's edge. A browser selection is one range of the
 * page, which can't hold a column or a block of cells, so the editor keeps the rectangle itself
 * and shades its cells (the browser's own highlight hidden meanwhile).
 *
 * The rectangle is on the table's grid (HTML's table model: rowspan and colspan, from Word's
 * vMerge and gridSpan), and grows until no merged cell crosses its edge, as Word's does. A nested
 * table is a table of its own: its cells belong to the cell of the outer table they sit in.
 */

/** A cell's place on its table's grid. */
export interface GridCell {
  cell: HTMLTableCellElement;
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}

/** A rectangle of grid slots: rows r0..r1, columns c0..c1 (inclusive). */
export interface CellRect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** A table's grid: which cell fills each slot. */
export class TableGrid {
  readonly cells: GridCell[] = [];
  readonly slots: (GridCell | undefined)[][] = [];
  readonly columns: number;
  private readonly byCell = new Map<HTMLTableCellElement, GridCell>();

  constructor(readonly table: HTMLTableElement) {
    // (table.rows and row.cells are the table's own: a nested table's rows are not among them)
    Array.from(table.rows).forEach((tr, r) => {
      this.slots[r] ??= [];
      let c = 0;
      for (const cell of Array.from(tr.cells)) {
        while (this.slots[r][c]) c++;
        const rowSpan = Math.max(1, cell.rowSpan || 1), colSpan = Math.max(1, cell.colSpan || 1);
        const placed: GridCell = { cell, row: r, col: c, rowSpan, colSpan };
        this.cells.push(placed);
        this.byCell.set(cell, placed);
        for (let dr = 0; dr < rowSpan; dr++) {
          const row = (this.slots[r + dr] ??= []);
          for (let dc = 0; dc < colSpan; dc++) row[c + dc] = placed;
        }
        c += colSpan;
      }
    });
    this.columns = Math.max(0, ...this.slots.map((row) => row.length));
  }

  get rows(): number {
    return this.slots.length;
  }

  of(cell: HTMLTableCellElement): GridCell | undefined {
    return this.byCell.get(cell);
  }

  /** The rectangle holding both cells, grown until no merged cell crosses its edge. */
  rectFor(a: HTMLTableCellElement, b: HTMLTableCellElement): CellRect | null {
    const ga = this.of(a), gb = this.of(b);
    if (!ga || !gb) return null;
    const rect: CellRect = {
      r0: Math.min(ga.row, gb.row), c0: Math.min(ga.col, gb.col),
      r1: Math.max(ga.row + ga.rowSpan, gb.row + gb.rowSpan) - 1,
      c1: Math.max(ga.col + ga.colSpan, gb.col + gb.colSpan) - 1,
    };
    for (let grown = true; grown;) {
      grown = false;
      for (const g of this.cellsOf(rect)) {
        const r0 = Math.min(rect.r0, g.row), c0 = Math.min(rect.c0, g.col);
        const r1 = Math.max(rect.r1, g.row + g.rowSpan - 1), c1 = Math.max(rect.c1, g.col + g.colSpan - 1);
        if (r0 !== rect.r0 || c0 !== rect.c0 || r1 !== rect.r1 || c1 !== rect.c1) {
          Object.assign(rect, { r0, c0, r1, c1 });
          grown = true;
        }
      }
    }
    return rect;
  }

  /** The cells in a rectangle, each once, in document order. */
  cellsOf(rect: CellRect): GridCell[] {
    const seen = new Set<GridCell>();
    for (let r = rect.r0; r <= rect.r1; r++)
      for (let c = rect.c0; c <= rect.c1; c++) {
        const g = this.slots[r]?.[c];
        if (g) seen.add(g);
      }
    return this.cells.filter((g) => seen.has(g));
  }

  /** The whole table as a rectangle. */
  all(): CellRect {
    return { r0: 0, c0: 0, r1: Math.max(0, this.rows - 1), c1: Math.max(0, this.columns - 1) };
  }

  /** The cell a step from `cell` in a direction, on the grid (from the slot `at` within a merged
   *  cell, so a step down a column stays in that column); null at the table's edge. */
  step(cell: HTMLTableCellElement, dr: number, dc: number, at?: { row: number; col: number }): GridCell | null {
    const g = this.of(cell);
    if (!g) return null;
    let row = at?.row ?? g.row, col = at?.col ?? g.col;
    if (dr > 0) row = g.row + g.rowSpan; else if (dr < 0) row = g.row - 1;
    if (dc > 0) col = g.col + g.colSpan; else if (dc < 0) col = g.col - 1;
    return this.slots[row]?.[col] ?? null;
  }
}

/** `el`'s own cell in `table`: the cell of that table it is in (through any nested tables). */
export function ownCell(el: Element, table: HTMLTableElement): HTMLTableCellElement | null {
  for (let cell = el.closest<HTMLTableCellElement>("td, th"); cell; cell = cell.parentElement?.closest<HTMLTableCellElement>("td, th") ?? null)
    if (cell.closest("table") === table) return cell;
  return null;
}

/**
 * The innermost table both elements are in cells of, and their cells there; null when there is
 * none or they are in the same cell (a text selection then, inside that cell).
 */
export function commonTableCells(a: Element, b: Element): { table: HTMLTableElement; a: HTMLTableCellElement; b: HTMLTableCellElement } | null {
  for (let cell = a.closest<HTMLTableCellElement>("td, th"); cell; cell = cell.parentElement?.closest<HTMLTableCellElement>("td, th") ?? null) {
    const table = cell.closest<HTMLTableElement>("table");
    if (!table || !table.contains(b)) continue;
    const other = ownCell(b, table);
    if (!other) return null;
    return other === cell ? null : { table, a: cell, b: other };
  }
  return null;
}

export const CELL_SELECTED_CLASS = "docx-cell-selected";
export const CELL_SELECTING_ATTR = "data-docx-cell-selection";

export const CELL_SELECTION_CSS = `
.${CELL_SELECTED_CLASS} { background-image: linear-gradient(rgba(37, 99, 235, 0.28), rgba(37, 99, 235, 0.28)) !important; }
[${CELL_SELECTING_ATTR}] ::selection { background: transparent; }
[${CELL_SELECTING_ATTR}] [contenteditable="true"] { caret-color: transparent; }
.docx-table-handle { position: fixed; z-index: 40; box-sizing: border-box; width: 16px; height: 16px; padding: 0;
  margin: 0; border: 1px solid #6b7280; border-radius: 2px; background: #fff; color: #374151;
  font: 13px/13px system-ui, sans-serif; text-align: center; cursor: pointer; }
.docx-table-handle:hover { background: #e5e7eb; border-color: #374151; }
.docx-table-handle[hidden] { display: none; }
@media print { .docx-table-handle { display: none !important; } }
`;

/**
 * Word's table handle: a small ⊞ at a table's top-left corner, shown while the mouse is over the
 * table or the caret is in it (the innermost table); a click selects the whole table. It is
 * fixed-position in the page and follows the table through scrolling, hidden when the table's
 * corner is outside the editor's view (under a toolbar).
 */
export class TableHandle {
  readonly el: HTMLButtonElement;
  private table: HTMLTableElement | null = null;
  private hideTimer: ReturnType<typeof setTimeout> | undefined;

  constructor(private readonly doc: Document, onPick: (table: HTMLTableElement) => void,
    private readonly view: () => DOMRect | null) {
    ensureCellSelectionStyles(doc);
    const el = this.el = doc.createElement("button");
    el.type = "button";
    el.className = "docx-table-handle";
    el.textContent = "⊞";
    el.title = "Select the table";
    el.setAttribute("aria-label", "Select the table");
    el.hidden = true;
    // (the document keeps its focus and selection)
    el.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); });
    el.addEventListener("click", (e) => {
      e.preventDefault();
      if (this.table?.isConnected) onPick(this.table);
    });
    el.addEventListener("mouseenter", () => clearTimeout(this.hideTimer));
    (doc.body ?? doc.documentElement).appendChild(el);
    doc.addEventListener("scroll", this.place, true);
    doc.defaultView?.addEventListener("resize", this.place);
  }

  /** The table it is shown for, if any. */
  get shownFor(): HTMLTableElement | null {
    return this.el.hidden ? null : this.table;
  }

  show(table: HTMLTableElement): void {
    clearTimeout(this.hideTimer);
    this.table = table;
    this.place();
  }

  /** Hide after a moment (the mouse may be on its way to the handle, outside the table). */
  hideSoon(ms = 400): void {
    clearTimeout(this.hideTimer);
    this.hideTimer = setTimeout(() => this.hide(), ms);
  }

  hide(): void {
    clearTimeout(this.hideTimer);
    this.table = null;
    this.el.hidden = true;
  }

  readonly place = (): void => {
    const table = this.table;
    if (!table) return;
    if (!table.isConnected) { this.hide(); return; }
    const r = table.getBoundingClientRect();
    const left = r.left - 18, top = r.top - 18;
    const view = this.view();
    const inView = !view || (top >= view.top && top + 16 <= view.bottom && left + 16 >= view.left && left <= view.right);
    this.el.hidden = !inView;
    this.el.style.left = `${Math.round(left)}px`;
    this.el.style.top = `${Math.round(top)}px`;
  };

  destroy(): void {
    clearTimeout(this.hideTimer);
    this.doc.removeEventListener("scroll", this.place, true);
    this.doc.defaultView?.removeEventListener("resize", this.place);
    this.el.remove();
  }
}

const styled = new WeakSet<Document>();
export function ensureCellSelectionStyles(doc: Document): void {
  if (styled.has(doc)) return;
  styled.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusCellSelection = "true";
  style.textContent = CELL_SELECTION_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}
