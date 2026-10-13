// Shared formulas: Excel writes the formula text once, on the first cell of the block, and the other cells refer to
// it. Their own text (for the cell bar) is the first cell's formula with relative references moved by the offset.
import { SHEET_LIMITS } from '../limits.ts';
import { columnName } from './address.ts';

// A cell reference, or a whole-column or whole-row range. The lookarounds leave function names (LOG10), defined
// names and sheet names (AB1!A1 shifts only A1) alone; quoted strings and quoted sheet names are skipped by the caller.
const REFERENCE = /(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(!])|(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(!])|(?<![A-Za-z0-9_.:])(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_(!])/g;
const REF_ERROR = '#REF!';

/** A formula written for one cell as it reads in the cell `rows` down and `cols` right of it. */
export function shiftFormula(formula: string, rows: number, cols: number): string {
  if (rows === 0 && cols === 0) {
    return formula;
  }
  // Even pieces are outside quotes, odd ones are "strings" or 'sheet names' kept as they are.
  return formula
    .split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/)
    .map((piece, i) => (i % 2 ? piece : piece.replace(REFERENCE, (match, ...groups: (string | undefined)[]) => shiftMatch(match, groups, rows, cols))))
    .join('');
}

function shiftMatch(match: string, groups: (string | undefined)[], rows: number, cols: number): string {
  const [colAbs = '', colText, rowAbs = '', rowText, startColAbs = '', startCol, endColAbs = '', endCol, startRowAbs = '', startRow, endRowAbs = '', endRow] = groups;
  let parts: [string, string | undefined, string, string | undefined] | undefined;
  if (colText !== undefined && rowText !== undefined) {
    parts = [colAbs, shiftColumn(colText, colAbs, cols), rowAbs, shiftRow(rowText, rowAbs, rows)];
  } else if (startCol !== undefined && endCol !== undefined) {
    parts = [startColAbs, shiftColumn(startCol, startColAbs, cols), `:${endColAbs}`, shiftColumn(endCol, endColAbs, cols)];
  } else if (startRow !== undefined && endRow !== undefined) {
    parts = [startRowAbs, shiftRow(startRow, startRowAbs, rows), `:${endRowAbs}`, shiftRow(endRow, endRowAbs, rows)];
  }
  if (!parts || parts[1] === undefined || parts[3] === undefined) {
    return match;
  }
  // A reference moved off the grid becomes #REF! as a whole, as in Excel.
  return parts[1] === REF_ERROR || parts[3] === REF_ERROR ? REF_ERROR : parts.join('');
}

function shiftColumn(letters: string, absolute: string, offset: number): string | undefined {
  let col = 0;
  for (const ch of letters.toUpperCase()) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  if (col > SHEET_LIMITS.maxCols) {
    // Past column XFD: part of a defined name, not a reference.
    return undefined;
  }
  if (absolute) {
    return letters;
  }
  const shifted = col - 1 + offset;
  return shifted < 0 || shifted >= SHEET_LIMITS.maxCols ? REF_ERROR : columnName(shifted);
}

function shiftRow(digits: string, absolute: string, offset: number): string | undefined {
  const row = Number(digits);
  if (row < 1 || row > SHEET_LIMITS.maxRows) {
    return undefined;
  }
  if (absolute) {
    return digits;
  }
  const shifted = row + offset;
  return shifted < 1 || shifted > SHEET_LIMITS.maxRows ? REF_ERROR : String(shifted);
}
