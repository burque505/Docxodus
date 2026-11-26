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
`;

const styled = new WeakSet<Document>();
export function ensureCellSelectionStyles(doc: Document): void {
  if (styled.has(doc)) return;
  styled.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusCellSelection = "true";
  style.textContent = CELL_SELECTION_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}
