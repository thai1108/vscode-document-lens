// A .csv or .tsv file as a one-sheet workbook. Every field is shown as the text it holds (nothing is converted to
// numbers or dates, so `007` and `1.50` stay as written); fields that look like numbers are right-aligned.
import { decodeText, type TextEncoding } from '../../../csv/encoding.ts';
import { detectDelimiter, LineEndings, type LineEnding } from '../../../csv/dialect.ts';
import { fieldValue, readRecords, stringUnits, type RecordSink } from '../../../csv/tokenizer.ts';
import { SHEET_LIMITS } from '../limits.ts';
import { CellKind } from '../model/cells.ts';
import { SheetBuilder } from '../model/sheet.ts';
import { DEFAULT_FONT, defaultFormat, EMPTY_STRINGS, type CellFormat, type Sheet, type Workbook } from '../model/types.ts';

export interface CsvDetails {
  readonly delimiter: string;
  readonly encoding: TextEncoding;
  /** The most common line break (known once the sheet is read); `mixedLineEndings` when there is more than one kind. */
  lineEnding: LineEnding;
  mixedLineEndings: boolean;
  /** The first row is shown as a header: bold and frozen. */
  readonly headerRow: boolean;
}

export interface CsvOptions {
  /** .tsv files are always tab-separated; .csv files get their delimiter detected. */
  tabSeparated: boolean;
  /** Show the first row as a header; detected when not given. */
  headerRow?: boolean;
}

export const CSV_STYLE = { normal: 0, header: 1, number: 2 } as const;

const ROW_HEIGHT = 20;
const MIN_COLUMN_WIDTH = 64;
const MAX_COLUMN_WIDTH = 400;
/** Rows looked at to size the columns. */
const WIDTH_SAMPLE_ROWS = 1000;
/** Records read between checks of the clock, and the longest stretch of reading before the page gets a turn. */
const BATCH_RECORDS = 2000;
const MAX_BUSY_MS = 30;

/** Plain numbers, with thousands separators and a decimal point or comma, an exponent or a percent sign. */
const NUMBER = /^[+-]?(?:\d{1,3}(?:[,.' ]\d{3})+|\d+)?(?:[.,]\d+)?(?:[eE][+-]?\d+)?%?$/;

function csvFormats(): CellFormat[] {
  const normal = defaultFormat(DEFAULT_FONT);
  return [
    normal,
    { ...normal, font: { ...DEFAULT_FONT, bold: true }, fill: '#f2f2f2', borders: { bottom: { style: 'thin', color: '#a6a6a6' } } },
    { ...normal, horizontal: 'right' },
  ];
}

/** Opens the file: decodes it and works out its delimiter; the cells are read when the sheet is shown. */
export function openCsv(bytes: Uint8Array, name: string, options: CsvOptions): Workbook & { readonly csv: CsvDetails } {
  const decoded = decodeText(bytes);
  const { text } = decoded;
  const delimiter = options.tabSeparated ? '\t' : detectDelimiter(text);
  const headerRow = options.headerRow ?? looksLikeHeader(text, delimiter);
  const formats = csvFormats();
  const expected = decoded.encoding.name !== 'utf-8' ? 'UTF-16' : decoded.encoding.bom ? 'UTF-8' : 'UTF-8 or Shift_JIS';
  const notices = decoded.lossy ? [`Some bytes of this file are not valid ${expected} text and are shown as \uFFFD.`] : [];
  const workbook = {
    sheets: [{ name, hidden: false, kind: 'worksheet' as const }],
    activeSheet: 0,
    date1904: false,
    formats,
    sharedStrings: EMPTY_STRINGS,
    notices,
    csv: { delimiter, encoding: decoded.encoding, headerRow, lineEnding: '\r\n' as LineEnding, mixedLineEndings: false },
    async loadSheet(index: number, signal?: AbortSignal): Promise<Sheet> {
      if (index !== 0) {
        throw new Error(`No sheet ${index}`);
      }
      const lineEndings = new LineEndings();
      const sheet = await readSheet(text, name, delimiter, headerRow, formats, lineEndings, signal);
      workbook.csv.lineEnding = lineEndings.dominant;
      workbook.csv.mixedLineEndings = lineEndings.mixed;
      return sheet;
    },
  };
  return workbook;
}

async function readSheet(
  text: string,
  name: string,
  delimiter: string,
  headerRow: boolean,
  formats: readonly CellFormat[],
  lineEndings: LineEndings,
  signal?: AbortSignal,
): Promise<Sheet> {
  const sheet = new SheetBuilder(name, 'worksheet', formats);
  const widths: number[] = [];
  let row = 0;
  let col = 0;
  let stopped = false;
  let rowsCut = false;
  let colsCut = false;
  let longFields = false;
  let openQuoteRow: number | undefined;

  const sink: RecordSink = {
    field(start, end, open) {
      if (open) {
        openQuoteRow = row;
      }
      if (col >= SHEET_LIMITS.maxCols) {
        colsCut = true;
        return;
      }
      let value = fieldValue(text, start, end, SHEET_LIMITS.maxStringLength);
      if (value.length > SHEET_LIMITS.maxStringLength) {
        longFields = true;
        value = value.slice(0, SHEET_LIMITS.maxStringLength);
      }
      const header = headerRow && row === 0;
      if (value === '') {
        if (header) {
          sheet.add(row, col, CellKind.Blank, 0, CSV_STYLE.header);
        }
      } else {
        const style = header ? CSV_STYLE.header : isNumber(value) ? CSV_STYLE.number : CSV_STYLE.normal;
        if (!sheet.add(row, col, CellKind.SheetString, sheet.addString(value), style)) {
          stopped = true;
        }
        if (row < WIDTH_SAMPLE_ROWS) {
          widths[col] = Math.max(widths[col] ?? 0, textWidth(value, header));
        }
      }
      col++;
    },
    record(end, next) {
      lineEndings.add(text, end, next);
      row++;
      col = 0;
      if (row >= SHEET_LIMITS.maxRows && next < text.length) {
        rowsCut = true;
        stopped = true;
      }
      return !stopped;
    },
  };

  const units = stringUnits(text);
  const code = delimiter.charCodeAt(0);
  let position = 0;
  let busySince = Date.now();
  while (position < text.length && !stopped) {
    position = readRecords(units, code, sink, position, BATCH_RECORDS);
    if (Date.now() - busySince > MAX_BUSY_MS) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      signal?.throwIfAborted();
      busySince = Date.now();
    }
  }

  if (rowsCut) {
    sheet.notice(`Only the first ${SHEET_LIMITS.maxRows.toLocaleString('en-US')} rows are shown.`);
  }
  if (colsCut) {
    sheet.notice(`Fields after column ${SHEET_LIMITS.maxCols.toLocaleString('en-US')} are not shown.`);
  }
  if (longFields) {
    sheet.notice(`Fields longer than ${SHEET_LIMITS.maxStringLength.toLocaleString('en-US')} characters are cut off.`);
  }
  if (openQuoteRow !== undefined) {
    sheet.notice(`A quote opened in row ${(openQuoteRow + 1).toLocaleString('en-US')} is never closed: the rest of the file is in that cell.`);
  }
  const cols = new Map<number, number>();
  widths.forEach((width, index) => {
    if (width > MIN_COLUMN_WIDTH) {
      cols.set(index, Math.min(MAX_COLUMN_WIDTH, width));
    }
  });
  return sheet.build({
    rows: { defaultSize: ROW_HEIGHT, sizes: new Map() },
    cols: { defaultSize: MIN_COLUMN_WIDTH, sizes: cols },
    activeCell: { row: 0, col: 0 },
    ...(headerRow ? { frozen: { rows: 1, cols: 0, top: 0, left: 0 } } : {}),
  });
}

/** Width in pixels the text needs in an 11pt cell: 7px per character, twice that for wide (CJK) characters. */
function textWidth(text: string, bold: boolean): number {
  let units = 0;
  const line = text.length > 200 ? text.slice(0, 200) : text;
  for (let i = 0; i < line.length; i++) {
    const code = line.charCodeAt(i);
    if (code === 0x0a || code === 0x0d) {
      break;
    }
    units += isWide(code) ? 2 : 1;
  }
  return Math.ceil(units * (bold ? 7.5 : 7) + 12);
}

/**
 * Whether the first row reads like column names: at least two rows, and every field of the first one non-empty,
 * not a number and different from the others.
 */
function looksLikeHeader(text: string, delimiter: string): boolean {
  const rows: string[][] = [];
  let fields: string[] = [];
  readRecords(
    stringUnits(text),
    delimiter.charCodeAt(0),
    {
      field(start, end) {
        if (fields.length < 1000) {
          fields.push(fieldValue(text, start, end, 256));
        }
      },
      record() {
        rows.push(fields);
        fields = [];
        return true;
      },
    },
    0,
    2,
  );
  const first = rows[0];
  if (!first || rows.length < 2) {
    return false;
  }
  return first.every((value) => value.trim() !== '' && !isNumber(value.trim())) && new Set(first).size === first.length;
}

function isNumber(text: string): boolean {
  return NUMBER.test(text) && /\d/.test(text);
}

/** East Asian wide characters: Hangul Jamo, CJK, kana, Hangul syllables, full-width forms. */
function isWide(code: number): boolean {
  return (
    (code >= 0x1100 && code <= 0x115f) ||
    (code >= 0x2e80 && code <= 0xa4cf) ||
    (code >= 0xac00 && code <= 0xd7a3) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xfe30 && code <= 0xfe4f) ||
    (code >= 0xff00 && code <= 0xff60) ||
    (code >= 0xffe0 && code <= 0xffe6)
  );
}
