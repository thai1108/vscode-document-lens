// sample.tsv: UTF-16 LE with a byte order mark and tabs, as Excel's "Unicode Text" writes; asserted in
// test/csv.test.ts and test/e2e/csv.spec.ts.
import type { Sample } from '../support/samples.ts';

export const SAMPLE_TSV_ROWS = [
  ['Region', 'Sales', 'Comment'],
  ['東京', '1200', 'Tab\tinside quotes'],
  ['大阪', '980', 'Emoji 🎉 and comma, here'],
];

export function utf16(text: string, endianness: 'le' | 'be', bom = true): Uint8Array {
  const bytes = Buffer.from(text, 'utf16le');
  if (endianness === 'be') {
    bytes.swap16();
  }
  const mark = endianness === 'le' ? [0xff, 0xfe] : [0xfe, 0xff];
  return new Uint8Array([...(bom ? mark : []), ...bytes]);
}

export function sampleTsv(): Uint8Array {
  const line = (fields: string[]) => fields.map((field) => (/[\t"\r\n]/.test(field) ? `"${field.replaceAll('"', '""')}"` : field)).join('\t');
  return utf16(SAMPLE_TSV_ROWS.map(line).join('\r\n') + '\r\n', 'le');
}

export const samples: Sample[] = [{ fileName: 'sample.tsv', build: sampleTsv }];
