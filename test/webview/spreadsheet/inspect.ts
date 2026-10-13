// Reads one cell of a loaded sheet the way the grid shows it, for the reader tests.
import assert from 'node:assert/strict';
import { cellDisplay, cellRawText } from '../../../src/webview/spreadsheet/model/display.ts';
import { defaultFormat, type CellFormat, type RichRun, type Sheet, type Workbook } from '../../../src/webview/spreadsheet/model/types.ts';

export interface InspectedCell {
  text: string;
  /** What the cell bar shows. */
  raw: string;
  format: CellFormat;
  align: string;
  color?: string;
  runs?: readonly RichRun[];
}

/** What a cell shows, its cell bar text and its format, by A1 address. */
export function inspect(workbook: Workbook, sheet: Sheet, ref: string): InspectedCell {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  assert.ok(match, ref);
  const col = [...(match[1] ?? '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  const index = sheet.cells.find(row, col);
  assert.ok(index >= 0, `${ref} is stored`);
  const format = workbook.formats[sheet.cells.style(index)] ?? defaultFormat();
  const display = cellDisplay(workbook, sheet, index, format);
  return {
    text: display.text,
    raw: cellRawText(workbook, sheet, index, row, col, format),
    format,
    align: display.align,
    ...(display.color ? { color: display.color } : {}),
    ...(display.runs ? { runs: display.runs } : {}),
  };
}
