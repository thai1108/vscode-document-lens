// The format-neutral sheet model: building a sheet without any file format (as a CSV reader would), reading cells
// back, and moving the selection over it.
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { CellRange } from '../../../src/webview/spreadsheet/model/address.ts';
import { CellKind, NO_FORMULA } from '../../../src/webview/spreadsheet/model/cells.ts';
import { readCell } from '../../../src/webview/spreadsheet/model/display.ts';
import { RangeIndex } from '../../../src/webview/spreadsheet/model/rangeIndex.ts';
import { emptySheet, SheetBuilder } from '../../../src/webview/spreadsheet/model/sheet.ts';
import { defaultFormat, EMPTY_STRINGS, type Sheet, type Workbook } from '../../../src/webview/spreadsheet/model/types.ts';
import { Axis } from '../../../src/webview/spreadsheet/view/axis.ts';
import { CellSelection, type NavigationBounds } from '../../../src/webview/spreadsheet/view/selection.ts';

const FORMATS = [defaultFormat(), { ...defaultFormat(), fill: '#ffff00' }, { ...defaultFormat(), numberFormat: '0.0%' }];
const LAYOUT = { rows: { defaultSize: 20, sizes: new Map<number, number>() }, cols: { defaultSize: 64, sizes: new Map<number, number>() } };

function workbookOf(sheet: Sheet): Workbook {
  return { sheets: [{ name: sheet.name, hidden: false, kind: sheet.kind }], activeSheet: 0, date1904: false, formats: FORMATS, sharedStrings: EMPTY_STRINGS, notices: [], loadSheet: () => Promise.resolve(sheet) };
}

/** Rows of text the way a CSV reader would hand them over: numbers as numbers, the rest as text. */
function sheetFromRows(rows: string[][], maxCells?: number): Sheet {
  const builder = new SheetBuilder('data.csv', 'worksheet', FORMATS, maxCells === undefined ? {} : { maxCells });
  rows.forEach((fields, row) =>
    fields.forEach((field, col) => {
      const number = Number(field);
      if (field !== '' && Number.isFinite(number)) {
        builder.add(row, col, CellKind.Number, number, 0);
      } else if (field !== '') {
        builder.add(row, col, CellKind.SheetString, builder.addString(field), 0);
      }
    }),
  );
  return builder.build(LAYOUT);
}

describe('SheetBuilder', () => {
  it('builds a sheet from plain rows, as a CSV reader would', () => {
    const sheet = sheetFromRows([
      ['name', 'qty'],
      ['tea', '3'],
      ['', '0.5'],
    ]);
    const workbook = workbookOf(sheet);
    assert.deepEqual(sheet.used, { rows: 3, cols: 2 });
    assert.deepEqual(readCell(workbook, sheet, 1, 0), { row: 1, col: 0, kind: CellKind.SheetString, value: 'tea', text: 'tea', raw: 'tea', style: 0 });
    assert.deepEqual(readCell(workbook, sheet, 2, 1), { row: 2, col: 1, kind: CellKind.Number, value: 0.5, text: '0.5', raw: '0.5', style: 0 });
    assert.equal(readCell(workbook, sheet, 2, 0), undefined);
    assert.equal(sheet.sourcePart, undefined);
    assert.deepEqual(sheet.notices, []);
  });

  it('keeps blanks only when their format shows, and counts merges in the used range', () => {
    const builder = new SheetBuilder('S', 'worksheet', FORMATS);
    builder.add(0, 0, CellKind.Blank, 0, 0);
    builder.add(1, 1, CellKind.Blank, 0, 1);
    builder.addMerge({ top: 4, left: 0, bottom: 5, right: 6 });
    builder.addMerge({ top: 9, left: 9, bottom: 9, right: 9 });
    const sheet = builder.build(LAYOUT);
    assert.equal(sheet.cells.size, 1);
    assert.equal(readCell(workbookOf(sheet), sheet, 1, 1)?.kind, CellKind.Blank);
    assert.deepEqual(sheet.used, { rows: 6, cols: 7 });
    assert.equal(sheet.merges.length, 1, 'a one-cell merge is no merge');
  });

  it('gives formulas back as written, shared ones moved to each cell', () => {
    const builder = new SheetBuilder('S', 'worksheet', FORMATS);
    builder.defineSharedFormula(0, 'A1*2', 0, 1);
    builder.add(0, 1, CellKind.Number, 2, 2, builder.addFormula('A1*2'));
    builder.add(3, 1, CellKind.Number, 8, 2, builder.sharedFormula(0));
    builder.add(4, 1, CellKind.Number, 1, 0, builder.sharedFormula(7));
    const sheet = builder.build(LAYOUT);
    const workbook = workbookOf(sheet);
    assert.deepEqual(readCell(workbook, sheet, 3, 1), { row: 3, col: 1, kind: CellKind.Number, value: 8, text: '800.0%', raw: '=A4*2', formula: 'A4*2', style: 2 });
    assert.equal(readCell(workbook, sheet, 4, 1)?.formula, undefined, 'an unknown shared formula is no formula');
    assert.equal(builder.addFormula(''), NO_FORMULA);
  });

  it('cuts the sheet off at the cell cap, says so, and ignores later cells', () => {
    const sheet = sheetFromRows(
      Array.from({ length: 5 }, () => ['a', 'b']),
      5,
    );
    assert.equal(sheet.cells.size, 5);
    assert.deepEqual(sheet.notices, ['Only the first 2 rows of “data.csv” are shown: it has more than 5 cells.']);
  });

  it('makes empty sheets with a notice, and notes drawings', () => {
    assert.deepEqual(emptySheet('Chart', 'chartsheet', 20, 64, 'no charts').notices, ['no charts']);
    const builder = new SheetBuilder('Pictures', 'worksheet', FORMATS);
    builder.noteDrawings();
    assert.deepEqual(builder.build({ ...LAYOUT, sourcePart: 'xl/worksheets/sheet3.xml' }).notices, ['Pictures, charts and shapes in “Pictures” are not shown.']);
  });
});

describe('CellSelection', () => {
  // Row 2 is hidden; A4:B5 is merged.
  const rows = new Axis({ defaultSize: 20, sizes: new Map([[2, 0]]) }, 1_048_576);
  const cols = new Axis({ defaultSize: 64, sizes: new Map() }, 16_384);
  const bounds: NavigationBounds = { rows, cols, rowCount: 100, colCount: 30, lastRow: 49, lastCol: 9, page: 10, origin: { row: 1, col: 0 } };
  const selection = () => new CellSelection(new RangeIndex((range: CellRange) => range, [{ top: 3, left: 0, bottom: 4, right: 1 }]));

  it('selects cells by row and column, a merged cell as a whole', () => {
    const cells = selection();
    assert.equal(cells.range, undefined);
    assert.deepEqual(cells.select(4, 1), { top: 3, left: 0, bottom: 4, right: 1 });
    assert.deepEqual(cells.active, { row: 3, col: 0 });
    assert.deepEqual(cells.select(-5, 1e9), { top: 0, left: 16_383, bottom: 0, right: 16_383 });
  });

  it('moves with the keys, skipping hidden rows and stepping out of merged cells', () => {
    const cells = selection();
    assert.deepEqual(cells.target('ArrowDown', false, bounds), { row: 3, col: 0 }, 'from the origin, over the hidden row 2');
    cells.select(3, 0);
    assert.deepEqual(cells.target('ArrowDown', false, bounds), { row: 5, col: 0 });
    assert.deepEqual(cells.target('ArrowRight', false, bounds), { row: 3, col: 2 });
    assert.deepEqual(cells.target('ArrowUp', false, bounds), { row: 1, col: 0 });
    assert.deepEqual(cells.target('ArrowDown', true, bounds), { row: 49, col: 0 });
    assert.deepEqual(cells.target('PageDown', false, bounds), { row: 13, col: 0 });
    assert.deepEqual(cells.target('End', true, bounds), { row: 49, col: 9 });
    assert.deepEqual(cells.target('Home', false, bounds), { row: 3, col: 0 });
    assert.equal(cells.target('a', false, bounds), undefined);
    cells.select(0, 0);
    assert.deepEqual(cells.target('ArrowUp', false, bounds), { row: 0, col: 0 }, 'stays at the edge');
  });
});
