// .csv samples in the encodings and dialects Excel and other tools write. Every value here is asserted in
// test/csv.test.ts, test/webview/spreadsheet/csv.dom.test.ts or test/e2e/csv.spec.ts.
import type { Sample } from '../support/samples.ts';
import { encodeShiftJis } from '../support/shiftJis.ts';

/** sample.csv as rows of field values. */
export const SAMPLE_CSV_ROWS = [
  ['ID', 'Item', 'Price', 'Note', 'Code'],
  ['1', 'Coffee beans', '1,234.50', 'Fair trade, organic', '007'],
  ['2', 'Green tea 緑茶', '980', 'Line one\nline two', '1.50'],
  ['3', 'Rice crackers', '312', 'Says "crunchy"', ''],
  ['4', 'Short row'],
  ['5', 'Long row', '45.8', '', 'E', 'extra field'],
];

/** UTF-8 with a byte order mark and CRLF line breaks, as Excel's "CSV UTF-8" writes; quotes only where needed. */
export function sampleCsv(): Uint8Array {
  const line = (fields: string[]) => fields.map((field) => (/[",\r\n]/.test(field) ? `"${field.replaceAll('"', '""')}"` : field)).join(',');
  const text = SAMPLE_CSV_ROWS.map(line).join('\r\n') + '\r\n';
  return new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(text)]);
}

export const SHIFT_JIS_ROWS = [
  ['商品コード', '商品名', '価格', '備考'],
  ['A-001', '緑茶（ペットボトル）', '150', '全角スペース　入り'],
  ['A-002', 'ほうじ茶', '130', '①②③ と ﾊﾝｶｸ'],
];

/** Shift_JIS without a byte order mark, as Excel's plain "CSV" writes on Japanese Windows. */
export function shiftJisCsv(): Uint8Array {
  return encodeShiftJis(SHIFT_JIS_ROWS.map((row) => row.join(',')).join('\r\n') + '\r\n');
}

export const SEMICOLON_ROWS = [
  ['Artikel', 'Menge', 'Preis'],
  ['Kaffee', '2', '12,50'],
  ['Tee', '10', '3,95'],
];

/** Semicolons and decimal commas with LF line breaks and no final line break, as written in many European locales. */
export function semicolonCsv(): Uint8Array {
  return new TextEncoder().encode(SEMICOLON_ROWS.map((row) => row.join(';')).join('\n'));
}

export const samples: Sample[] = [
  { fileName: 'sample.csv', build: sampleCsv },
  { fileName: 'shift-jis.csv', build: shiftJisCsv },
  { fileName: 'semicolon.csv', build: semicolonCsv },
];
