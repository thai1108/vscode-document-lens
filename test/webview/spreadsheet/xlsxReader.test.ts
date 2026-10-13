// Reading .xlsx workbooks into the viewer's model: the styled sample, then hand-written variants of the format.
import assert from 'node:assert/strict';
import { before, describe, it } from 'node:test';
import { CellKind } from '../../../src/webview/spreadsheet/model/cells.ts';
import { DEFAULT_THEME, defaultPalette } from '../../../src/webview/spreadsheet/model/colors.ts';
import { readCell } from '../../../src/webview/spreadsheet/model/display.ts';
import { DEFAULT_FONT, defaultFormat, type Sheet, type Workbook } from '../../../src/webview/spreadsheet/model/types.ts';
import { openWorkbook } from '../../../src/webview/spreadsheet/viewer.ts';
import { readWorksheet } from '../../../src/webview/spreadsheet/xlsx/sheet.ts';
import { DATA_ROWS, GITHUB_URL, sampleXlsx } from '../../fixtures/xlsx.ts';
import { cell, xlsxPackage } from '../../support/xlsxWriter.ts';
import { inspect } from './inspect.ts';

async function* chunked(text: string, size = 7): AsyncGenerator<string> {
  for (let i = 0; i < text.length; i += size) {
    yield text.slice(i, i + size);
  }
}

const worksheetContext = {
  name: 'Test',
  kind: 'worksheet' as const,
  formats: [defaultFormat(), { ...defaultFormat(), fill: '#ff0000' }],
  sharedStringCount: 0,
  date1904: false,
  defaultFont: DEFAULT_FONT,
  colors: { palette: defaultPalette(), theme: DEFAULT_THEME },
  relationships: new Map(),
};

describe('sample.xlsx', () => {
  let workbook: Workbook;
  let summary: Sheet;
  before(async () => {
    workbook = await openWorkbook(new Uint8Array(sampleXlsx()));
    summary = await workbook.loadSheet(0);
  });

  it('lists the sheets with hidden ones marked, and starts on the active one', () => {
    assert.deepEqual(workbook.sheets, [
      { name: 'Summary', hidden: false, kind: 'worksheet' },
      { name: 'Data', hidden: false, kind: 'worksheet' },
      { name: 'Hidden notes', hidden: true, kind: 'worksheet' },
      { name: 'メモ', hidden: false, kind: 'worksheet' },
    ]);
    assert.equal(workbook.activeSheet, 0);
    assert.equal(workbook.date1904, false);
    assert.deepEqual(workbook.notices, []);
  });

  it('reads column widths, row heights, hidden rows and columns, merges and frozen panes', () => {
    assert.equal(summary.cols.defaultSize, 64);
    assert.equal(summary.cols.sizes.get(0), 173);
    assert.equal(summary.cols.sizes.get(5), 0, 'column F is hidden');
    assert.equal(summary.rows.defaultSize, 20);
    assert.equal(summary.rows.sizes.get(0), 40);
    assert.equal(summary.rows.sizes.get(3), 60);
    assert.equal(summary.rows.sizes.get(10), 0, 'row 11 is hidden');
    assert.deepEqual(summary.merges, [{ top: 0, left: 0, bottom: 0, right: 3 }]);
    assert.deepEqual(summary.frozen, { rows: 2, cols: 1, top: 0, left: 0 });
    assert.deepEqual(summary.topLeft, { row: 2, col: 1 });
    assert.deepEqual(summary.activeCell, { row: 2, col: 1 });
    assert.deepEqual(summary.used, { rows: 15, cols: 6 });
    assert.equal(summary.showGridLines, true);
  });

  it('formats numbers, dates, percentages, booleans and errors', () => {
    assert.deepEqual([inspect(workbook, summary, 'B3').text, inspect(workbook, summary, 'B3').raw], ['1,234.50', '1234.5']);
    assert.deepEqual(inspect(workbook, summary, 'B7'), { ...inspect(workbook, summary, 'B7'), text: '-120.00', color: '#ff0000', align: 'right' });
    assert.deepEqual([inspect(workbook, summary, 'C3').text, inspect(workbook, summary, 'C3').raw], ['2024-01-15', '2024-01-15']);
    assert.equal(inspect(workbook, summary, 'D4').text, '39.97%');
    assert.equal(inspect(workbook, summary, 'A13').text, '令和6年1月15日');
    assert.deepEqual([inspect(workbook, summary, 'B13').text, inspect(workbook, summary, 'B13').align], ['TRUE', 'center']);
    assert.equal(inspect(workbook, summary, 'C13').text, '#DIV/0!');
  });

  it('shows formulas in the cell bar, shared formulas moved to their cell', () => {
    assert.equal(inspect(workbook, summary, 'B8').raw, '=SUM(B3:B7)');
    assert.equal(inspect(workbook, summary, 'D3').raw, '=B3/$B$8');
    assert.equal(inspect(workbook, summary, 'D7').raw, '=B7/$B$8');
  });

  it('says where each sheet came from, and reads single cells as plain objects', () => {
    assert.equal(summary.sourcePart, 'xl/worksheets/sheet1.xml');
    assert.deepEqual(readCell(workbook, summary, 7, 1), {
      row: 7,
      col: 1,
      kind: CellKind.Number,
      value: 2452.55,
      text: '2,452.55',
      raw: '=SUM(B3:B7)',
      formula: 'SUM(B3:B7)',
      style: 7,
    });
  });

  it('reads shared strings: Japanese text, phonetic guides left out, rich text runs', () => {
    assert.equal(inspect(workbook, summary, 'A1').text, 'Quarterly report 四半期報告');
    assert.equal(inspect(workbook, summary, 'E3').text, '東京本社');
    const rich = inspect(workbook, summary, 'A10');
    assert.equal(rich.text, 'Plain, bold red, italic');
    assert.deepEqual(rich.runs, [
      { text: 'Plain, ' },
      { text: 'bold red', font: { bold: true, size: 11, color: '#ff0000', name: 'Calibri' } },
      { text: ', italic', font: { italic: true, size: 11, name: 'Calibri' } },
    ]);
  });

  it('resolves fills, fonts, borders and alignment, theme colors from the file', () => {
    const title = inspect(workbook, summary, 'A1').format;
    assert.equal(title.fill, '#1f4e79');
    assert.deepEqual([title.font.size, title.font.bold, title.font.color, title.horizontal, title.vertical], [16, true, '#ffffff', 'center', 'center']);
    assert.deepEqual(inspect(workbook, summary, 'A2').format.borders, { bottom: { style: 'medium', color: '#000000' } });
    assert.deepEqual(inspect(workbook, summary, 'B8').format.borders, { top: { style: 'thin', color: '#000000' }, bottom: { style: 'double', color: '#000000' } });
    const note = inspect(workbook, summary, 'E4').format;
    assert.deepEqual([note.wrap, note.vertical, note.font.name], [true, 'top', '游ゴシック']);
    assert.equal(inspect(workbook, summary, 'A15').format.font.color, '#c0504d', 'accent2 of the file theme');
    assert.equal(inspect(workbook, summary, 'E9').format.font.underline, 'single');
  });

  it('reads hyperlinks: web links by URL, places in the workbook by location', () => {
    assert.deepEqual(summary.links, [
      { range: { top: 8, left: 4, bottom: 8, right: 4 }, url: `${GITHUB_URL}` },
      { range: { top: 13, left: 0, bottom: 13, right: 0 }, location: 'Data!A1' },
    ]);
  });

  it('reads the other sheets: inline strings, string formulas, no gridlines', async () => {
    const data = await workbook.loadSheet(1);
    assert.equal(data.cells.size, (DATA_ROWS + 1) * 6 + 4);
    assert.deepEqual(data.frozen, { rows: 1, cols: 0, top: 0, left: 0 });
    assert.equal(inspect(workbook, data, 'B301').text, 'Product 300');
    assert.deepEqual([inspect(workbook, data, 'G2').text, inspect(workbook, data, 'G2').raw], ['1-Product 1', '=A2&"-"&B2']);
    assert.equal(inspect(workbook, data, 'A1').format.fill, '#b9cde5', 'theme accent1 with a 60% tint');
    const memo = await workbook.loadSheet(3);
    assert.equal(memo.showGridLines, false);
    assert.equal(inspect(workbook, memo, 'A1').text, '日本語のメモ');
  });
});

describe('.xlsx variants', () => {
  it('uses the 1904 date system when the workbook says so', async () => {
    const workbook = await openWorkbook(
      new Uint8Array(
        xlsxPackage({
          date1904: true,
          styles: '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>',
          sheets: [{ name: 'S', body: `<sheetData><row r="1">${cell('A1', 43844, { style: 1 })}</row></sheetData>` }],
        }),
      ),
    );
    assert.equal(workbook.date1904, true);
    assert.equal(inspect(workbook, await workbook.loadSheet(0), 'A1').raw, '2024-01-15');
  });

  it('reads cells without r attributes, out of order, with prefixes, ISO dates and inline rich text', async () => {
    const body =
      '<x:sheetData xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<x:row r="3"><x:c><x:v>1</x:v></x:c><x:c><x:v>2</x:v></x:c></x:row>' +
      '<x:row r="1"><x:c r="B1" t="d"><x:v>2024-01-15T12:00:00</x:v></x:c><x:c r="A1" t="inlineStr"><x:is><x:r><x:rPr><x:b/></x:rPr><x:t>bold</x:t></x:r><x:r><x:t> plain</x:t></x:r></x:is></x:c></x:row>' +
      '<x:row><x:c r="C4" t="s"><x:v>99</x:v></x:c><x:c r="D4" s="999"><x:v>5</x:v></x:c><x:c r="E4"><x:v/></x:c></x:row>' +
      '</x:sheetData>';
    const workbook = await openWorkbook(new Uint8Array(xlsxPackage({ sheets: [{ name: 'S', body }] })));
    const sheet = await workbook.loadSheet(0);
    assert.equal(inspect(workbook, sheet, 'A3').raw, '1');
    assert.equal(inspect(workbook, sheet, 'B3').raw, '2');
    assert.equal(inspect(workbook, sheet, 'B1').raw, '45306.5', 'an ISO date as a serial date');
    assert.deepEqual(inspect(workbook, sheet, 'A1').runs, [{ text: 'bold', font: { bold: true } }, { text: ' plain' }]);
    assert.equal(sheet.cells.find(3, 2), -1, 'a shared string index past the table is no value');
    assert.equal(inspect(workbook, sheet, 'D4').format, workbook.formats[0], 'an unknown style is the default style');
    assert.equal(sheet.cells.find(3, 4), -1, 'an empty value without formatting is not stored');
  });

  it('shows chart sheets and missing parts as empty sheets with a notice, and drawings with a notice', async () => {
    const workbook = await openWorkbook(
      new Uint8Array(
        xlsxPackage({
          sheets: [
            { name: 'Chart', chart: true, body: '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' },
            { name: 'Pictures', body: '<sheetData/><drawing r:id="rId1"/>' },
          ],
        }),
      ),
    );
    assert.equal(workbook.sheets[0]?.kind, 'chartsheet');
    assert.deepEqual((await workbook.loadSheet(0)).notices, ['“Chart” is a chart sheet; charts are not shown.']);
    assert.deepEqual((await workbook.loadSheet(1)).notices, ['Pictures, charts and shapes in “Pictures” are not shown.']);
  });

  it('refuses packages that are not workbooks', async () => {
    const empty = xlsxPackage({ sheets: [] });
    await assert.rejects(openWorkbook(new Uint8Array(empty)), /lists no sheets/);
    await assert.rejects(openWorkbook(new Uint8Array(Buffer.from('plain text'))), /Not a spreadsheet file/);
  });

  it('keeps only http(s) links and drops scripts and file links', async () => {
    const workbook = await openWorkbook(
      new Uint8Array(
        xlsxPackage({
          sheets: [
            {
              name: 'S',
              body: '<sheetData/><hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="A2" r:id="rId2"/><hyperlink ref="A3" r:id="rId3" location="top"/><hyperlink ref="A4:B5" location="\'Other sheet\'!C3"/></hyperlinks>',
              relationships: [
                { id: 'rId1', type: 'hyperlink', target: 'javascript:alert(1)', external: true },
                { id: 'rId2', type: 'hyperlink', target: 'file:///etc/passwd', external: true },
                { id: 'rId3', type: 'hyperlink', target: 'https://example.com/page', external: true },
              ],
            },
          ],
        }),
      ),
    );
    assert.deepEqual((await workbook.loadSheet(0)).links, [
      { range: { top: 2, left: 0, bottom: 2, right: 0 }, url: 'https://example.com/page#top' },
      { range: { top: 3, left: 0, bottom: 4, right: 1 }, location: "'Other sheet'!C3" },
    ]);
  });
});

describe('readWorksheet', () => {
  it('cuts a sheet off at the cell cap with a notice', async () => {
    const rows = Array.from({ length: 10 }, (_value, r) => `<row r="${r + 1}">${cell(`A${r + 1}`, r)}${cell(`B${r + 1}`, r)}</row>`).join('');
    const sheet = await readWorksheet(chunked(`<worksheet><sheetData>${rows}</sheetData></worksheet>`), { ...worksheetContext, maxCells: 7 });
    assert.equal(sheet.cells.size, 7);
    assert.deepEqual(sheet.notices, ['Only the first 3 rows of “Test” are shown: it has more than 7 cells.']);
  });

  it('keeps formatted blanks (they show) and counts them as used', async () => {
    const sheet = await readWorksheet(chunked('<worksheet><sheetData><row r="5"><c r="C5" s="1"/><c r="D5"/></row></sheetData></worksheet>'), worksheetContext);
    assert.equal(sheet.cells.size, 1);
    assert.deepEqual(sheet.used, { rows: 5, cols: 3 });
  });

  it('stops when the load is abandoned', async () => {
    const controller = new AbortController();
    controller.abort();
    await assert.rejects(readWorksheet(chunked('<worksheet><sheetData/></worksheet>'), worksheetContext, controller.signal));
  });
});
