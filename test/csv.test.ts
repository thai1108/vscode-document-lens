// The delimited text core shared by the host and the webview (src/csv): encodings, the tokenizer and the dialect.
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CSV_DELIMITERS, detectDelimiter, LineEndings } from '../src/csv/dialect.ts';
import { decodeText, encodingLabel, sniffEncoding } from '../src/csv/encoding.ts';
import { fieldValue, readRecords, stringUnits, type Units } from '../src/csv/tokenizer.ts';
import { checkText } from '../src/preflight.ts';
import { SAMPLE_CSV_ROWS, SEMICOLON_ROWS, SHIFT_JIS_ROWS, sampleCsv, semicolonCsv, shiftJisCsv } from './fixtures/csv.ts';
import { SAMPLE_TSV_ROWS, sampleTsv, utf16 } from './fixtures/tsv.ts';
import { encodeShiftJis } from './support/shiftJis.ts';
import { textPdf } from './support/pdfWriter.ts';
import { writeZip } from './support/zipWriter.ts';
import { compoundFileHeaderStub } from './support/compoundFile.ts';

interface Parsed {
  rows: string[][];
  /** [start, end, open] of every field, and [end, next] of every record. */
  fields: [number, number, boolean][];
  records: [number, number][];
}

function parse(text: string, delimiter = ','): Parsed {
  return parseUnits(stringUnits(text), delimiter, (start, end) => fieldValue(text, start, end));
}

function parseUnits(units: Units, delimiter: string, value: (start: number, end: number) => string): Parsed {
  const parsed: Parsed = { rows: [], fields: [], records: [] };
  let row: string[] = [];
  readRecords(units, delimiter.charCodeAt(0), {
    field(start, end, open) {
      row.push(value(start, end));
      parsed.fields.push([start, end, open]);
    },
    record(end, next) {
      parsed.rows.push(row);
      parsed.records.push([end, next]);
      row = [];
      return true;
    },
  });
  return parsed;
}

function byteUnits(bytes: Uint8Array): Units {
  return { length: bytes.length, at: (index) => bytes[index] ?? 0 };
}

describe('readRecords and fieldValue', () => {
  it('reads RFC 4180 quoting: delimiters, line breaks and doubled quotes inside quotes', () => {
    assert.deepEqual(parse('a,"b,c","d\r\ne","say ""hi"""\r\n').rows, [['a', 'b,c', 'd\r\ne', 'say "hi"']]);
  });

  it('accepts CRLF, LF and CR line breaks, and a last line without one', () => {
    assert.deepEqual(parse('a\r\nb\nc\rd').rows, [['a'], ['b'], ['c'], ['d']]);
    assert.deepEqual(parse('a\r\nb\nc\rd').records, [
      [1, 3],
      [4, 5],
      [6, 7],
      [8, 8],
    ]);
  });

  it('keeps rows of any length, empty fields and empty lines', () => {
    assert.deepEqual(parse('a,b,c\nd\n\ne,,\n').rows, [['a', 'b', 'c'], ['d'], [''], ['e', '', '']]);
    assert.deepEqual(parse(',').rows, [['', '']]);
    assert.deepEqual(parse('').rows, []);
    assert.deepEqual(parse('\n').rows, [['']]);
  });

  it('reads past the RFC the way Excel does: text after a closing quote, quotes inside unquoted fields', () => {
    assert.deepEqual(parse('"ab"cd,e"f"g,"h""').rows, [['abcd', 'e"f"g', 'h"']]);
  });

  it('runs a quote that is never closed to the end of the text and says so', () => {
    const parsed = parse('a,"open\nb,c\n');
    assert.deepEqual(parsed.rows, [['a', 'open\nb,c\n']]);
    assert.deepEqual(parsed.fields.at(-1), [2, 12, true]);
    assert.equal(parse('"closed"').fields[0]?.[2], false);
  });

  it('stops after a budget of records and goes on from where it stopped', () => {
    const text = 'a,b\n"c\nd",e\nf\n';
    const rows: string[][] = [];
    let row: string[] = [];
    const sink = {
      field: (start: number, end: number) => void row.push(fieldValue(text, start, end)),
      record: () => {
        rows.push(row);
        row = [];
        return true;
      },
    };
    let position = 0;
    const stops: number[] = [];
    while (position < text.length) {
      position = readRecords(stringUnits(text), 44, sink, position, 1);
      stops.push(position);
    }
    assert.deepEqual(stops, [4, 12, 14]);
    assert.deepEqual(rows, parse(text).rows);
  });

  it('splits encoded bytes exactly where it splits the decoded text (UTF-8, Shift_JIS)', () => {
    const text = '名前,"説明, ""引用"" 付き"\r\n東京,"改行\nあり"\r\nｶﾅ,表\r\n';
    const decoded = parse(text);
    for (const bytes of [new TextEncoder().encode(text), encodeShiftJis(text)]) {
      const encoding = bytes.length === new TextEncoder().encode(text).length ? 'utf-8' : 'shift_jis';
      const decoder = new TextDecoder(encoding);
      const fromBytes = parseUnits(byteUnits(bytes), ',', (start, end) => {
        const raw = decoder.decode(bytes.subarray(start, end));
        return fieldValue(raw, 0, raw.length);
      });
      assert.deepEqual(fromBytes.rows, decoded.rows, encoding);
      assert.equal(fromBytes.records.length, decoded.records.length);
    }
  });

  it('caps a value without building all of a huge field', () => {
    const huge = `"${'x'.repeat(20_000_000)}`;
    const started = performance.now();
    const parsed = parseUnits(stringUnits(huge), ',', (start, end) => fieldValue(huge, start, end, 100));
    assert.ok(performance.now() - started < 2000, 'linear in the length of the text');
    assert.deepEqual(parsed.fields, [[0, huge.length, true]]);
    assert.equal(parsed.rows[0]?.[0]?.length, 101);
    assert.equal(fieldValue('"ab""cd""ef"', 0, 12, 3), 'ab"c');
    assert.equal(fieldValue('abcdef', 0, 6, 2), 'abc');
    assert.equal(fieldValue('"ab"cdef', 0, 8, 3), 'abcd');
  });

  it('reads a line of 100,000 delimiters and a million empty lines in linear time', () => {
    for (const text of [','.repeat(100_000), '\n'.repeat(1_000_000)]) {
      const started = performance.now();
      let fields = 0;
      readRecords(stringUnits(text), 44, { field: () => void fields++, record: () => true });
      assert.ok(performance.now() - started < 2000);
      assert.ok(fields >= 100_000);
    }
  });
});

describe('decodeText and sniffEncoding', () => {
  it('reads a byte order mark: UTF-8, UTF-16 LE and BE', () => {
    assert.deepEqual(decodeText(sampleCsv()).encoding, { name: 'utf-8', bom: 3 });
    assert.equal(decodeText(sampleCsv()).text.startsWith('ID,'), true, 'without the mark');
    const tsv = decodeText(sampleTsv());
    assert.deepEqual(tsv.encoding, { name: 'utf-16le', bom: 2 });
    assert.ok(tsv.text.includes('東京') && tsv.text.includes('🎉'));
    assert.deepEqual(decodeText(utf16('a,b\r\n', 'be')).encoding, { name: 'utf-16be', bom: 2 });
  });

  it('recognizes UTF-16 without a byte order mark by its zero bytes', () => {
    assert.deepEqual(decodeText(utf16('a,b\r\n東京,1\r\n', 'le', false)), { encoding: { name: 'utf-16le', bom: 0 }, text: 'a,b\r\n東京,1\r\n', lossy: false });
    assert.deepEqual(decodeText(utf16('a,b\r\n東京,1\r\n', 'be', false)).encoding, { name: 'utf-16be', bom: 0 });
  });

  it('tells UTF-8 from Shift_JIS', () => {
    assert.deepEqual(decodeText(new TextEncoder().encode('名前,値\n')).encoding, { name: 'utf-8', bom: 0 });
    const shiftJis = decodeText(shiftJisCsv());
    assert.deepEqual(shiftJis.encoding, { name: 'shift_jis', bom: 0 });
    assert.equal(shiftJis.text, SHIFT_JIS_ROWS.map((row) => row.join(',')).join('\r\n') + '\r\n');
    assert.equal(shiftJis.lossy, false);
  });

  it('reads a file valid in neither as UTF-8 and says some bytes were replaced', () => {
    const decoded = decodeText(new Uint8Array([0x61, 0x2c, 0xff, 0xfe, 0xfd, 0x0a]));
    assert.deepEqual(decoded.encoding, { name: 'utf-8', bom: 0 });
    assert.equal(decoded.lossy, true);
    assert.ok(decoded.text.startsWith('a,�'));
    const loneSurrogate = decodeText(new Uint8Array([0xff, 0xfe, 0x61, 0x00, 0x00, 0xd8, 0x62, 0x00]));
    assert.deepEqual([loneSurrogate.encoding.name, loneSurrogate.lossy], ['utf-16le', true]);
  });

  it('refuses binary data, but not a text file with a stray zero byte', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0]);
    assert.equal(sniffEncoding(png), 'binary');
    assert.throws(() => decodeText(png), /binary data/);
    const stray = new TextEncoder().encode(`${'a,b\n'.repeat(1000)}\0`);
    assert.equal(sniffEncoding(stray), undefined);
    assert.equal(sniffEncoding(new Uint8Array()), undefined);
    assert.equal(sniffEncoding(new Uint8Array([0x61])), undefined);
  });

  it('labels encodings for the toolbar', () => {
    assert.deepEqual(
      [{ name: 'utf-8', bom: 0 }, { name: 'utf-8', bom: 3 }, { name: 'utf-16le', bom: 2 }, { name: 'utf-16be', bom: 0 }, { name: 'shift_jis', bom: 0 }].map((encoding) =>
        encodingLabel(encoding as Parameters<typeof encodingLabel>[0]),
      ),
      ['UTF-8', 'UTF-8 with BOM', 'UTF-16 LE', 'UTF-16 BE (no BOM)', 'Shift_JIS'],
    );
  });
});

describe('detectDelimiter', () => {
  const text = (bytes: Uint8Array) => decodeText(bytes).text;

  it('detects commas, semicolons and tabs', () => {
    assert.equal(detectDelimiter(text(sampleCsv())), ',');
    assert.equal(detectDelimiter(text(semicolonCsv())), ';');
    assert.equal(detectDelimiter(text(sampleTsv())), '\t');
    assert.deepEqual(CSV_DELIMITERS, [',', ';', '\t']);
  });

  it('prefers semicolons over commas when the fields have decimal commas', () => {
    assert.equal(detectDelimiter('1,5;2,5\n3,5;4,5\n'), ';');
    assert.equal(detectDelimiter('a;b\n1,5;2\n'), ';');
  });

  it('ignores delimiters inside quotes and follows Excel’s sep= line', () => {
    assert.equal(detectDelimiter('"a;b;c",d\n"e;f;g",h\n'), ',');
    assert.equal(detectDelimiter('sep=;\na,b;c\nd,e;f\n'), ';');
  });

  it('falls back to commas for a single column or an empty file', () => {
    assert.equal(detectDelimiter('one\ntwo\n'), ',');
    assert.equal(detectDelimiter(''), ',');
  });

  it('looks at a sample, not the whole of a huge file', () => {
    const huge = 'a;b;c\n' + 'x'.repeat(50_000_000);
    const started = performance.now();
    assert.equal(detectDelimiter(huge), ';');
    assert.ok(performance.now() - started < 500);
  });
});

describe('LineEndings', () => {
  it('counts the line breaks by kind and names the most common', () => {
    const endings = new LineEndings();
    const source = 'a\r\nb\nc\nd\re';
    for (const [end, next] of parse(source).records) {
      endings.add(source, end, next);
    }
    assert.deepEqual([endings.crlf, endings.lf, endings.cr, endings.dominant, endings.mixed], [1, 2, 1, '\n', true]);
    assert.equal(new LineEndings().dominant, '\r\n');
  });
});

describe('the samples', () => {
  it('read back as the rows they were written from', () => {
    const rowsOf = (bytes: Uint8Array, delimiter?: string) => {
      const decoded = decodeText(bytes);
      return parse(decoded.text, delimiter ?? detectDelimiter(decoded.text)).rows;
    };
    assert.deepEqual(rowsOf(sampleCsv()), SAMPLE_CSV_ROWS);
    assert.deepEqual(rowsOf(shiftJisCsv()), SHIFT_JIS_ROWS);
    assert.deepEqual(rowsOf(semicolonCsv()), SEMICOLON_ROWS);
    assert.deepEqual(rowsOf(sampleTsv(), '\t'), SAMPLE_TSV_ROWS);
  });
});

describe('checkText', () => {
  it('accepts text in every encoding the viewer reads', () => {
    for (const bytes of [sampleCsv(), shiftJisCsv(), sampleTsv(), utf16('a\tb', 'be', false), new Uint8Array()]) {
      checkText(bytes, '.csv');
    }
  });

  it('names workbooks, PDFs and other binary files saved under a text extension', () => {
    assert.throws(() => checkText(writeZip([{ name: 'a.txt', data: 'x' }]), '.csv'), /This \.csv file is a ZIP package, probably an \.xlsx workbook/);
    assert.throws(() => checkText(compoundFileHeaderStub(), '.tsv'), /This \.tsv file is an Office binary file, probably an \.xls workbook/);
    assert.throws(() => checkText(textPdf([['x']]), '.csv'), /This \.csv file is a PDF document/);
    assert.throws(() => checkText(new Uint8Array(4096).fill(7).map((byte, i) => (i % 3 ? byte : 0)), '.csv'), /contains binary data, not text/);
  });
});
