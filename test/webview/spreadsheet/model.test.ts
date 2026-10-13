// The spreadsheet model: A1 references, cell storage, colors, sizes and shared formulas.
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { columnName, formatCellRef, formatRange, parseCellRef, parseRange } from '../../../src/webview/spreadsheet/model/address.ts';
import { CellKind, CellStore, CellStoreBuilder, NO_FORMULA } from '../../../src/webview/spreadsheet/model/cells.ts';
import { applyTint, defaultPalette, mixColors, parseHexColor, resolveColor, DEFAULT_THEME } from '../../../src/webview/spreadsheet/model/colors.ts';
import { shiftFormula } from '../../../src/webview/spreadsheet/model/formula.ts';
import { baseColumnWidthPx, columnWidthPx, defaultRowHeightPt, maxDigitWidth, pointsToPx } from '../../../src/webview/spreadsheet/model/sizes.ts';
import { DEFAULT_FONT } from '../../../src/webview/spreadsheet/model/types.ts';

describe('A1 references', () => {
  it('names columns', () => {
    assert.equal(columnName(0), 'A');
    assert.equal(columnName(25), 'Z');
    assert.equal(columnName(26), 'AA');
    assert.equal(columnName(701), 'ZZ');
    assert.equal(columnName(16_383), 'XFD');
    assert.equal(formatCellRef(99_999, 49), 'AX100000');
  });

  it('parses cell references, absolute or not, in any case', () => {
    assert.deepEqual(parseCellRef('A1'), { row: 0, col: 0 });
    assert.deepEqual(parseCellRef('$AB$12'), { row: 11, col: 27 });
    assert.deepEqual(parseCellRef('xfd1048576'), { row: 1_048_575, col: 16_383 });
  });

  it('refuses references outside the grid and junk', () => {
    for (const text of ['', 'A', '1', 'A0', 'XFE1', 'A1048577', 'A1B', '1A', 'A-1', 'Ä1', 'A 1']) {
      assert.equal(parseCellRef(text), undefined, text);
    }
  });

  it('parses ranges in any corner order and formats them back', () => {
    assert.deepEqual(parseRange('C3:A1'), { top: 0, left: 0, bottom: 2, right: 2 });
    assert.deepEqual(parseRange(' B2 '), { top: 1, left: 1, bottom: 1, right: 1 });
    assert.equal(parseRange('A1:'), undefined);
    assert.equal(formatRange({ top: 0, left: 0, bottom: 0, right: 3 }), 'A1:D1');
    assert.equal(formatRange({ top: 4, left: 1, bottom: 4, right: 1 }), 'B5');
  });
});

describe('CellStore', () => {
  it('stores cells by row and finds them by binary search', () => {
    const builder = new CellStoreBuilder(100);
    builder.add(0, 0, CellKind.Number, 1, 0);
    builder.add(0, 5, CellKind.SharedString, 3, 2, 7);
    builder.add(3, 2, CellKind.Boolean, 1, 1);
    const cells = builder.build();
    assert.equal(cells.size, 3);
    assert.equal(cells.rowCount, 4);
    assert.equal(cells.find(0, 5), 1);
    assert.equal(cells.kind(1), CellKind.SharedString);
    assert.equal(cells.value(1), 3);
    assert.equal(cells.style(1), 2);
    assert.equal(cells.formula(1), 7);
    assert.equal(cells.formula(0), NO_FORMULA);
    assert.equal(cells.find(0, 3), -1);
    assert.equal(cells.find(2, 2), -1);
    assert.equal(cells.find(9, 0), -1);
    assert.equal(cells.lowerBound(0, 1), 1);
    assert.equal(cells.rowStartIndex(3), 2);
    assert.equal(cells.rowEndIndex(3), 3);
    assert.equal(cells.rowStartIndex(1), cells.rowEndIndex(1));
  });

  it('sorts cells written out of order, the last write to a cell winning', () => {
    const builder = new CellStoreBuilder(100);
    builder.add(2, 1, CellKind.Number, 21, 0);
    builder.add(0, 3, CellKind.Number, 3, 0);
    builder.add(2, 0, CellKind.Number, 20, 0);
    builder.add(0, 3, CellKind.Number, 30, 0);
    const cells = builder.build();
    assert.equal(cells.size, 3);
    assert.deepEqual([0, 1, 2].map((i) => cells.value(i)), [30, 20, 21]);
    assert.equal(cells.find(2, 0), 1);
  });

  it('grows past its first capacity and stops at the cap', () => {
    const builder = new CellStoreBuilder(5000, 16);
    for (let i = 0; i < 6000; i++) {
      if (!builder.add(Math.floor(i / 10), i % 10, CellKind.Number, i, 0)) {
        assert.equal(i, 5000);
        break;
      }
    }
    const cells = builder.build();
    assert.equal(cells.size, 5000);
    assert.equal(cells.value(cells.find(499, 9)), 4999);
    assert.equal(CellStore.empty().size, 0);
  });
});

describe('colors', () => {
  it('reads RGB and ARGB hex, ignoring alpha', () => {
    assert.equal(parseHexColor('FF1F4E79'), '#1f4e79');
    assert.equal(parseHexColor('00ff0000'), '#ff0000');
    assert.equal(parseHexColor('abc'), undefined);
    assert.equal(parseHexColor('red'), undefined);
  });

  it('resolves theme and palette colors with tints', () => {
    const context = { palette: defaultPalette(), theme: DEFAULT_THEME };
    assert.equal(defaultPalette().length, 64);
    assert.equal(resolveColor({ indexed: 10 }, context), '#ff0000');
    assert.equal(resolveColor({ indexed: 64 }, context), undefined, 'system colors are automatic');
    assert.equal(resolveColor({ theme: 1 }, context), '#000000');
    assert.equal(resolveColor({ theme: 4, tint: 0.6 }, context), applyTint('#4472c4', 0.6));
    assert.equal(applyTint('#000000', 0.5), '#808080');
    assert.equal(applyTint('#ffffff', -0.5), '#808080');
    assert.equal(mixColors('#000000', '#ffffff', 0.25), '#bfbfbf');
  });
});

describe('Excel units', () => {
  it('derives the widest digit from the default font', () => {
    assert.equal(maxDigitWidth(DEFAULT_FONT), 7);
    assert.equal(maxDigitWidth({ ...DEFAULT_FONT, name: 'Arial', size: 10 }), 7);
    assert.equal(maxDigitWidth({ ...DEFAULT_FONT, name: '游ゴシック', size: 11 }), 8);
    assert.equal(maxDigitWidth({ ...DEFAULT_FONT, name: 'ＭＳ Ｐゴシック', size: 11 }), 7);
  });

  it('converts column widths and row heights to pixels like Excel', () => {
    assert.equal(columnWidthPx(9.140625, 7), 64, 'Excel shows 8.43 characters as 64 pixels');
    assert.equal(columnWidthPx(24.7109375, 7), 173);
    assert.equal(baseColumnWidthPx(8, 7), 64);
    assert.equal(baseColumnWidthPx(8, 8), 72);
    assert.equal(pointsToPx(15), 20);
    assert.equal(defaultRowHeightPt(DEFAULT_FONT), 15);
    assert.equal(defaultRowHeightPt({ ...DEFAULT_FONT, name: '游ゴシック' }), 18.75);
  });
});

describe('shared formulas', () => {
  it('moves relative references and keeps absolute ones', () => {
    assert.equal(shiftFormula('B3/$B$8', 2, 0), 'B5/$B$8');
    assert.equal(shiftFormula('SUM(A1:B2)*$C1+D$1', 1, 1), 'SUM(B2:C3)*$C2+E$1');
    assert.equal(shiftFormula('SUM(A:A)+SUM(1:1)', 1, 2), 'SUM(C:C)+SUM(2:2)');
  });

  it('leaves strings, sheet names, functions and names alone', () => {
    assert.equal(shiftFormula('"A1"&A1', 1, 0), '"A1"&A2');
    assert.equal(shiftFormula("'Q1 A1'!A1+Data!B2", 1, 0), "'Q1 A1'!A2+Data!B3");
    assert.equal(shiftFormula('LOG10(A1)+ATAN2(B1,C1)', 1, 0), 'LOG10(A2)+ATAN2(B2,C2)');
    assert.equal(shiftFormula('TAXES2024*XFE1*A1', 1, 0), 'TAXES2024*XFE1*A2');
    assert.equal(shiftFormula('AB1!A1', 1, 0), 'AB1!A2');
  });

  it('turns references moved off the grid into #REF!', () => {
    assert.equal(shiftFormula('A1', -1, 0), '#REF!');
    assert.equal(shiftFormula('SUM(A1:B2)', 0, -1), 'SUM(#REF!:A2)');
    assert.equal(shiftFormula('A1+B2', 0, 0), 'A1+B2');
  });
});
