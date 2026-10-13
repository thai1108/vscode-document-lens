// The selected cell (or the merged range it belongs to) and how keys move it, apart from drawing: cells are
// addressed by 0-based row and column, so other tools (an inline cell editor, say) can follow or drive it.
import { SHEET_LIMITS } from '../limits.ts';
import type { CellRange, CellRef } from '../model/address.ts';
import type { Axis } from './axis.ts';
import type { MergeIndex } from './merges.ts';

/** What key navigation needs to know about the grid. */
export interface NavigationBounds {
  rows: Axis;
  cols: Axis;
  rowCount: number;
  colCount: number;
  /** Last row and column with content (Ctrl/Cmd + arrows, End). */
  lastRow: number;
  lastCol: number;
  /** Rows Page Up/Down moves by. */
  page: number;
  /** Where moving starts when nothing is selected. */
  origin: CellRef;
}

export class CellSelection {
  private readonly merges: MergeIndex;
  private current: CellRange | undefined;

  constructor(merges: MergeIndex) {
    this.merges = merges;
  }

  /** The selected range: one cell, or a whole merged cell. */
  get range(): CellRange | undefined {
    return this.current;
  }

  /** The active cell: the top-left cell of the selected range. */
  get active(): CellRef | undefined {
    return this.current ? { row: this.current.top, col: this.current.left } : undefined;
  }

  /** Selects a cell (clamped to the grid), or the merged cell containing it; returns the selected range. */
  select(row: number, col: number): CellRange {
    const r = Math.max(0, Math.min(row, SHEET_LIMITS.maxRows - 1));
    const c = Math.max(0, Math.min(col, SHEET_LIMITS.maxCols - 1));
    this.current = this.merges.at(r, c) ?? { top: r, left: c, bottom: r, right: c };
    return this.current;
  }

  /** The cell `key` moves the selection to, or undefined for keys that do not move it. Hidden rows and columns are skipped. */
  target(key: string, jump: boolean, bounds: NavigationBounds): CellRef | undefined {
    const current = this.current ?? { top: bounds.origin.row, left: bounds.origin.col, bottom: bounds.origin.row, right: bounds.origin.col };
    switch (key) {
      case 'ArrowUp':
        return { row: jump ? 0 : step(bounds.rows, current.top, -1, bounds.rowCount), col: current.left };
      case 'ArrowDown':
        return { row: jump ? bounds.lastRow : step(bounds.rows, current.bottom, 1, bounds.rowCount), col: current.left };
      case 'ArrowLeft':
        return { row: current.top, col: jump ? 0 : step(bounds.cols, current.left, -1, bounds.colCount) };
      case 'ArrowRight':
        return { row: current.top, col: jump ? bounds.lastCol : step(bounds.cols, current.right, 1, bounds.colCount) };
      case 'PageUp':
        return { row: Math.max(0, current.top - bounds.page), col: current.left };
      case 'PageDown':
        return { row: Math.min(bounds.rowCount - 1, current.top + bounds.page), col: current.left };
      case 'Home':
        return jump ? { row: 0, col: 0 } : { row: current.top, col: 0 };
      case 'End':
        return jump ? { row: bounds.lastRow, col: bounds.lastCol } : { row: current.top, col: bounds.lastCol };
      default:
        return undefined;
    }
  }
}

/** The next visible entry from `from` in `direction`, or `from` at the edge. */
function step(axis: Axis, from: number, direction: 1 | -1, count: number): number {
  for (let index = from + direction; index >= 0 && index < count; index += direction) {
    if (axis.size(index) > 0) {
      return index;
    }
  }
  return from;
}
