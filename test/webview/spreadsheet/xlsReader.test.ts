// Reading .xls (BIFF8) workbooks: values from SheetJS, formatting from the record scan in xls/biff.ts.
import assert from 'node:assert/strict';
import { before, describe, it } from 'node:test';
import type { Sheet, Workbook } from '../../../src/webview/spreadsheet/model/types.ts';
import { openWorkbook } from '../../../src/webview/spreadsheet/viewer.ts';
import { readBiffFormatting } from '../../../src/webview/spreadsheet/xls/biff.ts';
import { sampleXls, XLS_LINK } from '../../fixtures/xls.ts';
import { biffStream, compoundFile, type BiffWorkbookSpec } from '../../support/biffWriter.ts';
import { inspect } from './inspect.ts';

const SIMPLE: BiffWorkbookSpec = {
  fonts: [{ name: 'Arial', size: 10 }],
  xfs: [{ font: 0 }],
  sheets: [{ name: 'S', cells: [{ row: 0, col: 0, xf: 0, value: 1 }] }],
};

/** A BIFF stream with a record inserted right after the workbook's BOF record. */
function withRecordAfterBof(stream: Buffer, type: number, body: Buffer): Buffer {
  const header = Buffer.alloc(4);
  header.writeUInt16LE(type, 0);
  header.writeUInt16LE(body.length, 2);
  const bofEnd = 4 + stream.readUInt16LE(2);
  return Buffer.concat([stream.subarray(0, bofEnd), header, body, stream.subarray(bofEnd)]);
}

describe('sample.xls', () => {
  let workbook: Workbook;
  let sales: Sheet;
  before(async () => {
    workbook = await openWorkbook(new Uint8Array(sampleXls()));
    sales = await workbook.loadSheet(0);
  });

  it('lists the sheets with hidden ones marked, in the 1904 date system', () => {
    assert.deepEqual(workbook.sheets, [
      { name: 'Sales', hidden: false, kind: 'worksheet' },
      { name: 'Hidden', hidden: true, kind: 'worksheet' },
      { name: 'Second', hidden: false, kind: 'worksheet' },
    ]);
    assert.equal(workbook.activeSheet, 0);
    assert.equal(workbook.date1904, true);
  });

  it('reads column widths, row heights, hidden columns, merges and frozen rows', () => {
    assert.equal(sales.cols.defaultSize, 64);
    assert.equal(sales.cols.sizes.get(0), 168, '24 characters of Arial 10');
    assert.equal(sales.cols.sizes.get(3), 0);
    assert.equal(sales.rows.sizes.get(0), 32);
    assert.equal(Math.round(sales.rows.sizes.get(8) ?? 0), 53);
    assert.deepEqual(sales.merges, [{ top: 0, left: 0, bottom: 0, right: 2 }]);
    assert.deepEqual(sales.frozen, { rows: 2, cols: 0 });
    assert.deepEqual(sales.topLeft, { row: 2, col: 0 });
  });

  it('formats values with the file’s number formats and 1904 dates', () => {
    assert.deepEqual([inspect(workbook, sales, 'B3').text, inspect(workbook, sales, 'B3').raw], ['1,500.00', '1500']);
    assert.deepEqual([inspect(workbook, sales, 'C3').text, inspect(workbook, sales, 'C3').raw], ['2024/01/15', '2024-01-15']);
    assert.deepEqual([inspect(workbook, sales, 'B6').text, inspect(workbook, sales, 'B6').raw], ['5,230.50', '=SUM(B3:B5)']);
    assert.equal(inspect(workbook, sales, 'A11').text, '12.50%');
    assert.equal(inspect(workbook, sales, 'B10').text, 'TRUE');
    assert.equal(inspect(workbook, sales, 'C10').text, '#DIV/0!');
    assert.equal(inspect(workbook, sales, 'A3').text, 'Tokyo 東京');
  });

  it('reads fonts, fills, borders and alignment from the XF records', () => {
    const title = inspect(workbook, sales, 'A1').format;
    assert.deepEqual([title.fill, title.font.name, title.font.size, title.font.bold, title.font.color, title.horizontal, title.vertical], ['#000080', 'Arial', 14, true, '#ffffff', 'center', 'center']);
    assert.deepEqual(inspect(workbook, sales, 'A2').format.borders, { bottom: { style: 'medium', color: '#000000' } });
    assert.equal(inspect(workbook, sales, 'A2').format.fill, '#c0c0c0');
    assert.deepEqual(inspect(workbook, sales, 'B3').format.borders.left, { style: 'thin', color: '#808080' });
    const note = inspect(workbook, sales, 'A9').format;
    assert.deepEqual([note.font.name, note.wrap, note.vertical], ['ＭＳ Ｐゴシック', true, 'top']);
    const accent = inspect(workbook, sales, 'A10').format.font;
    assert.deepEqual([accent.italic, accent.color], [true, '#ff0000']);
    const link = inspect(workbook, sales, 'A8').format.font;
    assert.deepEqual([link.underline, link.color], ['single', '#0000ff']);
  });

  it('reads hyperlinks and the gridline setting', async () => {
    assert.deepEqual(sales.links, [{ range: { top: 7, left: 0, bottom: 7, right: 0 }, url: XLS_LINK }]);
    assert.equal((await workbook.loadSheet(2)).showGridLines, false);
  });
});

describe('BIFF8 formatting records', () => {
  it('applies a custom palette and skips the missing font index 4', () => {
    const formatting = readBiffFormatting(
      biffStream({
        fonts: ['A', 'B', 'C', 'D', 'E'].map((name) => ({ name, size: 10, color: 8 })),
        palette: ['123456'],
        xfs: [{ font: 4, fill: { pattern: 1, foreground: 8 } }],
        sheets: [{ name: 'S', cells: [{ row: 0, col: 0, xf: 0, value: 1 }] }],
      }),
    );
    const format = formatting?.formats[15];
    assert.deepEqual([format?.font.name, format?.font.color, format?.fill], ['E', '#123456', '#123456']);
  });

  it('looks cell formats up by position, whatever order the records come in', () => {
    const formatting = readBiffFormatting(
      biffStream({
        fonts: [{ name: 'Arial', size: 10 }],
        xfs: [{ font: 0 }, { font: 0, wrap: true }],
        sheets: [
          {
            name: 'S',
            cells: [
              { row: 5, col: 1, xf: 1, value: 'x' },
              { row: 0, col: 2, xf: 0, value: null },
            ],
          },
        ],
      }),
    );
    const sheet = formatting?.sheets[0];
    assert.equal(sheet?.formatAt(5, 1), 16);
    assert.equal(sheet?.formatAt(0, 2), 15);
    assert.equal(sheet?.formatAt(1, 1), undefined);
  });

  it('notices pictures and charts, but not cell notes', () => {
    const objects = (types: number[]) => readBiffFormatting(biffStream({ ...SIMPLE, sheets: [{ name: 'S', cells: [], objects: types }] }))?.sheets[0]?.hasDrawings;
    assert.equal(objects([]), false);
    assert.equal(objects([0x19]), false);
    assert.equal(objects([0x19, 0x08]), true);
    assert.equal(objects([0x05]), true);
  });

  it('leaves older BIFF versions to SheetJS alone', () => {
    const stream = biffStream(SIMPLE);
    stream.writeUInt16LE(0x0500, 4);
    assert.equal(readBiffFormatting(stream), undefined);
  });
});

describe('.xls files that cannot be shown', () => {
  it('says that an encrypted workbook is password-protected', async () => {
    const encrypted = compoundFile(withRecordAfterBof(biffStream(SIMPLE), 0x002f, Buffer.from([0, 0, 0x34, 0x12, 0x78, 0x56])));
    await assert.rejects(openWorkbook(new Uint8Array(encrypted)), /This \.xls file is password-protected/);
  });

  it('says that another Office document is not a workbook', async () => {
    await assert.rejects(openWorkbook(new Uint8Array(compoundFile(Buffer.alloc(5000), 'WordDocument'))), /not an Excel workbook/);
  });

  it('fails cleanly (no hang, no crash) on damaged files', { timeout: 60_000 }, async () => {
    const original = sampleXls();
    let seed = 7;
    const random = () => {
      seed = (seed * 1_103_515_245 + 12_345) % 2 ** 31;
      return seed / 2 ** 31;
    };
    let opened = 0;
    for (let round = 0; round < 150; round++) {
      const damaged = Buffer.from(original);
      for (let i = 0; i < 8; i++) {
        damaged[Math.floor(random() * damaged.length)] = Math.floor(random() * 256);
      }
      try {
        const workbook = await openWorkbook(new Uint8Array(damaged));
        for (let sheet = 0; sheet < workbook.sheets.length; sheet++) {
          await workbook.loadSheet(sheet);
        }
        opened++;
      } catch (error) {
        assert.ok(error instanceof Error, `round ${round} threw ${String(error)}`);
      }
    }
    assert.ok(opened > 0, 'small damage usually still opens');
  });
});
