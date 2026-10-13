// sample.xls: a real Excel 97-2003 workbook (BIFF8 in a compound file) written by test/support/biffWriter.ts, with
// the 1904 date system so the .xls path covers it (sample.xlsx uses 1900).
import { biffWorkbook, type BiffCell } from '../support/biffWriter.ts';
import type { Sample } from '../support/samples.ts';

export const XLS_LINK = 'https://github.com/thai1108/vscode-document-lens/issues';

/** Cell formats (index into xfs). */
const XF = { normal: 0, title: 1, header: 2, amount: 3, date: 4, note: 5, link: 6, accent: 7, percent: 8 };
const THIN_GRAY: [number, number] = [1, 23];

/** 2024-01-15 and following months in the 1904 date system (1900 serial minus 1462). */
const DATES_1904 = [45306, 45337, 45366].map((serial) => serial - 1462);
const AMOUNTS = [1500, 2750.5, 980];

export function sampleXls(): Buffer {
  const cells: BiffCell[] = [
    { row: 0, col: 0, xf: XF.title, value: 'Sales 売上' },
    { row: 0, col: 1, xf: XF.title, value: null },
    { row: 0, col: 2, xf: XF.title, value: null },
    { row: 1, col: 0, xf: XF.header, value: 'Region' },
    { row: 1, col: 1, xf: XF.header, value: 'Amount' },
    { row: 1, col: 2, xf: XF.header, value: 'Date' },
    ...['Tokyo 東京', 'Osaka 大阪', 'Nagoya 名古屋'].flatMap((region, i): BiffCell[] => [
      { row: i + 2, col: 0, xf: XF.normal, value: region },
      { row: i + 2, col: 1, xf: XF.amount, value: AMOUNTS[i] ?? 0 },
      { row: i + 2, col: 2, xf: XF.date, value: DATES_1904[i] ?? 0 },
    ]),
    { row: 2, col: 3, xf: XF.normal, value: 'secret' },
    { row: 5, col: 0, xf: XF.header, value: 'Total' },
    { row: 5, col: 1, xf: XF.amount, value: AMOUNTS.reduce((sum, amount) => sum + amount, 0), sum: [2, 4, 1] },
    { row: 7, col: 0, xf: XF.link, value: 'Document Lens issues' },
    { row: 8, col: 0, xf: XF.note, value: '備考: 長いメモは折り返して表示されます。This note wraps inside its cell.' },
    { row: 9, col: 0, xf: XF.accent, value: 'Italic red' },
    { row: 9, col: 1, xf: XF.normal, value: true },
    { row: 9, col: 2, xf: XF.normal, value: { error: 0x07 } },
    { row: 10, col: 0, xf: XF.percent, value: 0.125 },
  ];
  return biffWorkbook({
    date1904: true,
    fonts: [
      { name: 'Arial', size: 10 },
      { name: 'Arial', size: 10, bold: true },
      { name: 'Arial', size: 14, bold: true, color: 9 },
      { name: 'Arial', size: 10, underline: true, color: 12 },
      { name: 'ＭＳ Ｐゴシック', size: 11 },
      { name: 'Arial', size: 10, italic: true, color: 10 },
    ],
    formats: [
      [164, '#,##0.00'],
      [165, 'yyyy/mm/dd'],
    ],
    xfs: [
      { font: 0 },
      { font: 2, horizontal: 2, vertical: 1, fill: { pattern: 1, foreground: 18 } },
      { font: 1, horizontal: 2, fill: { pattern: 1, foreground: 22 }, borders: { bottom: [2, 8] } },
      { font: 0, format: 164, borders: { top: THIN_GRAY, right: THIN_GRAY, bottom: THIN_GRAY, left: THIN_GRAY } },
      { font: 0, format: 165, borders: { top: THIN_GRAY, right: THIN_GRAY, bottom: THIN_GRAY, left: THIN_GRAY } },
      { font: 4, wrap: true, vertical: 0 },
      { font: 3 },
      { font: 5 },
      { font: 0, format: 10 },
    ],
    sheets: [
      {
        name: 'Sales',
        cells,
        merges: [[0, 0, 0, 2]],
        columns: [
          { first: 0, last: 0, width: 24 },
          { first: 1, last: 1, width: 14 },
          { first: 2, last: 2, width: 12 },
          { first: 3, last: 3, width: 9, hidden: true },
        ],
        rows: [
          { row: 0, height: 24 },
          { row: 8, height: 40 },
        ],
        frozen: { rows: 2, cols: 0 },
        links: [{ row: 7, col: 0, url: XLS_LINK }],
      },
      { name: 'Hidden', hidden: true, cells: [{ row: 0, col: 0, xf: XF.normal, value: 'hidden sheet' }] },
      {
        name: 'Second',
        showGridLines: false,
        cells: [
          { row: 0, col: 0, xf: XF.normal, value: 'Second sheet' },
          { row: 1, col: 0, xf: XF.normal, value: 3.5 },
        ],
      },
    ],
  });
}

export const samples: Sample[] = [{ fileName: 'sample.xls', build: sampleXls }];
