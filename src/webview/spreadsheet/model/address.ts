// A1-style cell references. Rows and columns are 0-based everywhere in the viewer; only these helpers speak A1.
import { SHEET_LIMITS } from '../limits.ts';

export interface CellRef {
  row: number;
  col: number;
}

/** Inclusive on all sides. */
export interface CellRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export function columnName(col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

export function formatCellRef(row: number, col: number): string {
  return `${columnName(col)}${row + 1}`;
}

export function formatRange(range: CellRange): string {
  const start = formatCellRef(range.top, range.left);
  return range.top === range.bottom && range.left === range.right ? start : `${start}:${formatCellRef(range.bottom, range.right)}`;
}

/**
 * Parses `B12` or `$B$12` (case-insensitive) into a 0-based reference inside Excel's grid. Written for the hot path
 * of reading sheet XML: no regular expression and no allocation.
 */
export function parseCellRef(text: string, from = 0, to = text.length): CellRef | undefined {
  let i = from;
  if (text.charCodeAt(i) === 36) {
    i++;
  }
  let col = 0;
  const lettersStart = i;
  for (; i < to; i++) {
    const code = text.charCodeAt(i) | 0x20;
    if (code < 97 || code > 122) {
      break;
    }
    col = col * 26 + (code - 96);
    if (col > SHEET_LIMITS.maxCols) {
      return undefined;
    }
  }
  if (i === lettersStart) {
    return undefined;
  }
  if (text.charCodeAt(i) === 36) {
    i++;
  }
  let row = 0;
  const digitsStart = i;
  for (; i < to; i++) {
    const code = text.charCodeAt(i);
    if (code < 48 || code > 57) {
      return undefined;
    }
    row = row * 10 + (code - 48);
    if (row > SHEET_LIMITS.maxRows) {
      return undefined;
    }
  }
  if (i === digitsStart || row === 0) {
    return undefined;
  }
  return { row: row - 1, col: col - 1 };
}

/** `A1:C3` or a single cell `B2`; the corners may come in any order. */
export function parseRange(text: string): CellRange | undefined {
  const trimmed = text.trim();
  const colon = trimmed.indexOf(':');
  const start = parseCellRef(trimmed, 0, colon < 0 ? trimmed.length : colon);
  const end = colon < 0 ? start : parseCellRef(trimmed, colon + 1);
  if (!start || !end) {
    return undefined;
  }
  return {
    top: Math.min(start.row, end.row),
    left: Math.min(start.col, end.col),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.col, end.col),
  };
}

export function rangeContains(range: CellRange, row: number, col: number): boolean {
  return row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
}
