// Crafted and odd files from the review of PR #15: each must stay fast and bounded, and show what Excel would.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as path from 'node:path';
import { describe, it } from 'node:test';
import { DEFAULT_THEME, defaultPalette } from '../../../src/webview/spreadsheet/model/colors.ts';
import { columnWidthPx, maxDigitWidth } from '../../../src/webview/spreadsheet/model/sizes.ts';
import { formatNumber, MAX_FORMAT_CODE_LENGTH, sectionColor } from '../../../src/webview/spreadsheet/model/numberFormat.ts';
import { DEFAULT_FONT, defaultFormat } from '../../../src/webview/spreadsheet/model/types.ts';
import { fontFamily } from '../../../src/webview/spreadsheet/view/styleRules.ts';
import { openWorkbook } from '../../../src/webview/spreadsheet/viewer.ts';
import { readBiffFormatting } from '../../../src/webview/spreadsheet/xls/biff.ts';
import { openXls } from '../../../src/webview/spreadsheet/xls/workbook.ts';
import { readWorksheet } from '../../../src/webview/spreadsheet/xlsx/sheet.ts';
import { XmlScanner } from '../../../src/webview/spreadsheet/xlsx/xml.ts';
import { biffStream, biffWorkbook, compoundFile } from '../../support/biffWriter.ts';
import { cell, xlsxPackage } from '../../support/xlsxWriter.ts';

const ROOT = path.resolve(import.meta.dirname, '../../..');

async function* chunks(text: string, size = 65_536): AsyncGenerator<string> {
  for (let i = 0; i < text.length; i += size) {
    yield text.slice(i, i + size);
  }
}

const context = {
  name: 'S',
  kind: 'worksheet' as const,
  formats: [defaultFormat()],
  sharedStringCount: 0,
  date1904: false,
  defaultFont: DEFAULT_FONT,
  colors: { palette: defaultPalette(), theme: DEFAULT_THEME },
  relationships: new Map(),
};

function record(type: number, body: Uint8Array): Buffer {
  const header = Buffer.alloc(4);
  header.writeUInt16LE(type, 0);
  header.writeUInt16LE(body.length, 2);
  return Buffer.concat([header, body]);
}

/** A BIFF8 stream whose `entries` BOUNDSHEET records all point at one sheet of `cells` blank cells. */
function repeatedSheetStream(entries: number, cells: number): Buffer {
  const bof = (kind: number) => {
    const body = Buffer.alloc(16);
    body.writeUInt16LE(0x0600, 0);
    body.writeUInt16LE(kind, 2);
    return record(0x0809, body);
  };
  const boundSheet = (position: number) => {
    const body = Buffer.from([0, 0, 0, 0, 0, 0, 2, 0, 0x53, 0x31]);
    body.writeUInt32LE(position, 0);
    return record(0x0085, body);
  };
  const globalsLength = bof(5).length + entries * boundSheet(0).length + 4;
  const globals = Buffer.concat([bof(5), ...Array.from({ length: entries }, () => boundSheet(globalsLength)), record(0x000a, Buffer.alloc(0))]);
  const perRecord = 250;
  const blanks: Buffer[] = [];
  for (let row = 0; row * perRecord < cells; row++) {
    const body = Buffer.alloc(6 + perRecord * 2);
    body.writeUInt16LE(row, 0);
    for (let i = 0; i < perRecord; i++) {
      body.writeUInt16LE(15, 4 + i * 2);
    }
    body.writeUInt16LE(perRecord - 1, 4 + perRecord * 2);
    blanks.push(record(0x00be, body));
  }
  return Buffer.concat([globals, bof(0x10), ...blanks, record(0x000a, Buffer.alloc(0))]);
}

describe('number format codes from the file', () => {
  it('replaces codes longer than Excel allows with General, in both readers', async () => {
    const long = '0;' + '['.repeat(40_000);
    const xlsx = await openWorkbook(
      new Uint8Array(
        xlsxPackage({
          styles: `<numFmts><numFmt numFmtId="200" formatCode="${long}"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="200"/></cellXfs>`,
          sheets: [{ name: 'S', body: `<sheetData><row r="1">${cell('A1', -1, { style: 1 })}</row></sheetData>` }],
        }),
      ),
    );
    assert.equal(xlsx.formats[1]?.numberFormat, 'General');
    const xls = readBiffFormatting(
      biffStream({ fonts: [{ name: 'Arial', size: 10 }], formats: [[164, '0;' + '['.repeat(300)]], xfs: [{ font: 0, format: 164 }], sheets: [{ name: 'S', cells: [] }] }),
    );
    assert.equal(xls?.formats[15]?.numberFormat, 'General');
    assert.equal(formatNumber(-1, long, false).text, '-1', 'an over-long code given directly is General too');
  });

  it('formats with the longest allowed pathological code quickly, working the code out once', () => {
    const code = ('0;' + '['.repeat(MAX_FORMAT_CODE_LENGTH)).slice(0, MAX_FORMAT_CODE_LENGTH);
    const started = performance.now();
    for (let i = 0; i < 10_000; i++) {
      formatNumber(-i, code, false);
      sectionColor(code, -i);
    }
    assert.ok(performance.now() - started < 5000);
  });

  it('reads era years without a day, month or year in the code', () => {
    assert.equal(formatNumber(45306, '[$-411]ee"年"', false).text, '06年');
    assert.equal(formatNumber(45306, '[$-411]e"年"', false).text, '6年');
    assert.equal(formatNumber(123456, '[$-411]0.00E+00', false).text, '1.23E+05');
  });
});

describe('.xls sheets listed many times', () => {
  it('reads one sheet once, whatever the number of BOUNDSHEET records pointing at it', () => {
    const stream = repeatedSheetStream(60, 250_000);
    const started = performance.now();
    const formatting = readBiffFormatting(stream);
    assert.equal(formatting?.sheets.length, 1, 'the duplicates are one sheet');
    assert.ok(performance.now() - started < 1000, 'nothing but the workbook globals is read up front');
    assert.equal(formatting?.readSheet(0).formatAt(999, 249), 15);
  });

  it('opens such a file without SheetJS objects for its blank cells', async () => {
    const started = performance.now();
    const workbook = openXls(new Uint8Array(compoundFile(repeatedSheetStream(60, 250_000))));
    assert.equal(workbook.sheets.length, 1);
    const sheet = await workbook.loadSheet(0);
    assert.equal(sheet.cells.size, 0, 'blanks with the default format show nothing and are not kept');
    assert.ok(performance.now() - started < 10_000);
  });

  it('still shows blank cells whose format has a fill or border', async () => {
    const file = biffWorkbook({
      fonts: [{ name: 'Arial', size: 10 }],
      xfs: [{ font: 0 }, { font: 0, borders: { top: [1, 8], bottom: [1, 8], left: [1, 8], right: [1, 8] } }, { font: 0, fill: { pattern: 1, foreground: 13 } }],
      sheets: [
        {
          name: 'S',
          cells: [
            { row: 0, col: 0, xf: 0, value: 'label' },
            { row: 1, col: 1, xf: 1, value: null },
            { row: 2, col: 2, xf: 2, value: null },
            { row: 3, col: 3, xf: 0, value: null },
          ],
        },
      ],
    });
    const sheet = await openXls(new Uint8Array(file)).loadSheet(0);
    assert.equal(sheet.cells.size, 3);
    assert.ok(sheet.cells.find(1, 1) >= 0 && sheet.cells.find(2, 2) >= 0);
    assert.equal(sheet.cells.find(3, 3), -1);
  });
});

describe('.xlsx sheets with crafted ranges', () => {
  it('reads thousands of full-width <col> elements quickly, the last one winning', async () => {
    const cols = '<col min="1" max="16384" width="9" customWidth="1"/>'.repeat(16_000) + '<col min="2" max="2" width="20" customWidth="1"/>';
    const started = performance.now();
    const sheet = await readWorksheet(chunks(`<worksheet><cols>${cols}</cols><sheetData/></worksheet>`), context);
    assert.ok(performance.now() - started < 5000);
    assert.equal(sheet.cols.sizes.size, 16_384);
    const digitWidth = maxDigitWidth(DEFAULT_FONT);
    assert.equal(sheet.cols.sizes.get(1), columnWidthPx(20, digitWidth));
    assert.equal(sheet.cols.sizes.get(16_383), columnWidthPx(9, digitWidth));
  });

  it('keeps one of many overlapping or repeated merged cells, as Excel repairs it', async () => {
    const merges = '<mergeCell ref="A1:J40"/>'.repeat(20_000) + '<mergeCell ref="B2:C3"/><mergeCell ref="K1:L1"/>';
    const sheet = await readWorksheet(chunks(`<worksheet><sheetData/><mergeCells>${merges}</mergeCells></worksheet>`), context);
    assert.deepEqual(sheet.merges, [
      { top: 0, left: 0, bottom: 39, right: 9 },
      { top: 0, left: 10, bottom: 0, right: 11 },
    ]);
  });

  it('handles merges, links and a dimension over the whole grid', async () => {
    const whole = 'A1:XFD1048576';
    const started = performance.now();
    const workbook = await openWorkbook(
      new Uint8Array(
        xlsxPackage({
          sheets: [
            {
              name: 'S',
              body: `<dimension ref="${whole}"/><sheetData><row r="1">${cell('A1', 'top left')}</row></sheetData><mergeCells><mergeCell ref="${whole}"/><mergeCell ref="B2:C3"/></mergeCells><hyperlinks><hyperlink ref="${whole}" r:id="rId1"/></hyperlinks>`,
              relationships: [{ id: 'rId1', type: 'hyperlink', target: 'https://example.com/', external: true }],
            },
          ],
        }),
      ),
    );
    const sheet = await workbook.loadSheet(0);
    assert.ok(performance.now() - started < 5000);
    assert.equal(sheet.merges.length, 1);
    assert.deepEqual(sheet.used, { rows: 1_048_576, cols: 16_384 });
    assert.equal(sheet.links[0]?.url, 'https://example.com/');
  });

  it('counts frozen rows from the view’s top-left cell and hides the rows scrolled out above them', async () => {
    const body = '<sheetViews><sheetView workbookViewId="0" topLeftCell="A5"><pane ySplit="2" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData/>';
    const sheet = await readWorksheet(chunks(`<worksheet>${body}</worksheet>`), context);
    assert.deepEqual(sheet.frozen, { rows: 2, cols: 0, top: 4, left: 0 });
    assert.deepEqual(sheet.topLeft, { row: 6, col: 0 });
    const xls = await openXls(
      new Uint8Array(
        biffWorkbook({
          fonts: [{ name: 'Arial', size: 10 }],
          xfs: [{ font: 0 }],
          sheets: [{ name: 'S', cells: [{ row: 0, col: 0, xf: 0, value: 1 }], frozen: { rows: 2, cols: 1 }, viewTopLeft: { row: 4, col: 3 } }],
        }),
      ),
    ).loadSheet(0);
    assert.deepEqual(xls.frozen, { rows: 2, cols: 1, top: 4, left: 3 });
    assert.deepEqual(xls.topLeft, { row: 6, col: 4 });
  });
});

describe('XML tokens split over many chunks', () => {
  it('finishes one long tag in time proportional to its length', () => {
    const tag = `<c ${'a'.repeat(900 * 1024)}/>`;
    const opened: string[] = [];
    const scanner = new XmlScanner({ open: (element) => opened.push(element.name) });
    const started = performance.now();
    for (let i = 0; i < tag.length; i += 4096) {
      scanner.write(tag.slice(i, i + 4096));
    }
    scanner.end();
    assert.deepEqual(opened, ['c']);
    assert.ok(performance.now() - started < 1000);
  });

  it('refuses a single tag over 1 MB', () => {
    const scanner = new XmlScanner({});
    assert.throws(() => {
      for (let i = 0; i < 40; i++) {
        scanner.write(i === 0 ? `<c ${'a'.repeat(65_536)}` : 'a'.repeat(65_536));
      }
    }, /too large/);
  });
});

describe('font names from the file', () => {
  it('cannot end the CSS value they are put in', () => {
    assert.equal(fontFamily('Calibri"; } body { background: url(x) } .a { x: "'), '"Calibri  body  background: url(x)  .a  x:", Calibri, Carlito, "Segoe UI", sans-serif');
    assert.equal(fontFamily('ＭＳ Ｐゴシック\u0000\\'), '"ＭＳ Ｐゴシック", "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic", Meiryo, sans-serif');
    assert.equal(fontFamily('";{}<>'), 'Calibri, Carlito, "Segoe UI", sans-serif');
  });
});

describe('the SheetJS dependency', () => {
  it('stays pinned by checksum in pnpm-lock.yaml (pnpm records none for URL tarballs by itself)', () => {
    const manifest = JSON.parse(readFileSync(path.join(ROOT, 'package.json'), 'utf8')) as { dependencies: Record<string, string> };
    const url = manifest.dependencies['xlsx'];
    assert.match(url ?? '', /^https:\/\/cdn\.sheetjs\.com\/xlsx-\d+\.\d+\.\d+\/xlsx-\d+\.\d+\.\d+\.tgz$/);
    const lock = readFileSync(path.join(ROOT, 'pnpm-lock.yaml'), 'utf8');
    const entry = lock.split('\n').find((line) => line.startsWith(`  xlsx@${url}:`));
    assert.ok(entry, 'the lockfile lists the tarball');
    const resolution = lock.split('\n')[lock.split('\n').indexOf(entry) + 1] ?? '';
    assert.match(resolution, new RegExp(`resolution: \\{integrity: sha512-[A-Za-z0-9+/]{86}==, tarball: ${url?.replaceAll('.', '\\.')}\\}`));
  });
});
