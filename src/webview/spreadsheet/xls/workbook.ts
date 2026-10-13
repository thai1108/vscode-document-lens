// An Excel 97-2003 .xls workbook: SheetJS reads the values, formulas, merges and links; biff.ts adds the formatting
// SheetJS leaves out. SheetJS reads the whole file at once (the format allows no less), sheets are converted to the
// viewer's model when first shown.
import { read as readWorkbook, CFB } from 'xlsx';
import { externalUrl } from '../../../protocol.ts';
import { SHEET_LIMITS } from '../limits.ts';
import { CellKind, NO_FORMULA } from '../model/cells.ts';
import { emptySheet, SheetBuilder } from '../model/sheet.ts';
import { baseColumnWidthPx, columnWidthPx, defaultRowHeightPt, maxDigitWidth, pointsToPx } from '../model/sizes.ts';
import { DEFAULT_FONT, defaultFormat, EMPTY_STRINGS, type CellFormat, type Font, type Hyperlink, type Sheet, type SheetEntry, type Workbook } from '../model/types.ts';
import { readBiffFormatting, type BiffSheet, type BiffWorkbook } from './biff.ts';

interface SheetJsCell {
  t: string;
  v?: unknown;
  f?: string;
  l?: { Target?: string; Tooltip?: string };
}

interface SheetJsSheet {
  '!data'?: (SheetJsCell[] | undefined)[];
  '!merges'?: { s: { r: number; c: number }; e: { r: number; c: number } }[];
  '!cols'?: ({ wpx?: number; hidden?: boolean } | undefined)[];
  '!rows'?: ({ hpx?: number; hidden?: boolean } | undefined)[];
  '!type'?: string;
}

interface SheetJsWorkbook {
  SheetNames: string[];
  Sheets: Record<string, SheetJsSheet | undefined>;
  Workbook?: { Sheets?: { Hidden?: number }[]; WBProps?: { date1904?: boolean } };
  cfb?: unknown;
}

const ERRORS: Record<number, string> = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA' };
/** The cell XF every BIFF writer puts first after the 15 style XFs. */
const DEFAULT_CELL_FORMAT = 15;

export function openXls(bytes: Uint8Array): Workbook {
  const book = readWithSheetJs(bytes);
  const formatting = readFormatting(book);
  const formats: readonly CellFormat[] = formatting?.formats ?? [defaultFormat()];
  const defaultFont = formatting?.defaultFont ?? DEFAULT_FONT;
  const date1904 = Boolean(book.Workbook?.WBProps?.date1904);
  const names = book.SheetNames.slice(0, SHEET_LIMITS.maxSheets);
  const formattingByName = new Map((formatting?.sheets ?? []).map((sheet) => [sheet.name, sheet]));
  const sheets: SheetEntry[] = names.map((name, i) => {
    const own = formattingByName.get(name);
    const chart = book.Sheets[name]?.['!type'] === 'chart';
    return {
      name,
      hidden: own?.hidden ?? Boolean(book.Workbook?.Sheets?.[i]?.Hidden),
      kind: chart ? 'chartsheet' : (own?.kind ?? 'worksheet'),
    };
  });
  const notices = book.SheetNames.length > names.length ? [`Only the first ${SHEET_LIMITS.maxSheets} sheets are shown.`] : [];
  const activeName = formatting?.sheets[formatting.activeSheet]?.name;
  const active = activeName === undefined ? -1 : names.indexOf(activeName);
  const firstVisible = Math.max(0, sheets.findIndex((sheet) => !sheet.hidden));
  return {
    sheets,
    activeSheet: active >= 0 && !sheets[active]?.hidden ? active : firstVisible,
    date1904,
    formats,
    sharedStrings: EMPTY_STRINGS,
    notices,
    loadSheet(index: number): Promise<Sheet> {
      const entry = sheets[index];
      if (!entry) {
        return Promise.reject(new Error(`No sheet ${index}`));
      }
      const own = formattingByName.get(entry.name);
      return Promise.resolve(convertSheet(entry, book.Sheets[entry.name] ?? {}, own, formats, defaultFont));
    },
  };
}

function readWithSheetJs(bytes: Uint8Array): SheetJsWorkbook {
  // SheetJS reports records it has no parser for through console.error; for a viewer that is noise, not an error.
  const consoleError = console.error;
  console.error = () => undefined;
  try {
    return readWorkbook(bytes, {
      type: 'array',
      dense: true,
      cellStyles: true,
      cellNF: false,
      cellFormula: true,
      cellHTML: false,
      cellText: false,
      cellDates: false,
      bookFiles: true,
    }) as unknown as SheetJsWorkbook;
  } catch (error) {
    throw xlsError(error);
  } finally {
    console.error = consoleError;
  }
}

function xlsError(error: unknown): Error {
  const message = error instanceof Error ? error.message : String(error);
  if (/password|encrypt/i.test(message)) {
    return new Error('This .xls file is password-protected; Document Lens cannot open encrypted files');
  }
  if (/Cannot find Workbook stream/i.test(message)) {
    return new Error('This file is not an Excel workbook (no Workbook stream); it may be another Office document with an .xls extension');
  }
  return new Error(`Not a readable .xls workbook: ${message}`, { cause: error });
}

function readFormatting(book: SheetJsWorkbook): BiffWorkbook | undefined {
  const find = CFB.find as (container: unknown, path: string) => { content?: ArrayLike<number> } | null;
  const stream = book.cfb ? (find(book.cfb, '/Workbook') ?? find(book.cfb, '/Book')) : null;
  if (!stream?.content) {
    return undefined;
  }
  const content = stream.content instanceof Uint8Array ? stream.content : Uint8Array.from(stream.content);
  try {
    return readBiffFormatting(content);
  } catch (error) {
    if ((error as Error).message === 'password-protected') {
      throw xlsError(error);
    }
    // The values are readable without the formatting.
    return undefined;
  }
}

function convertSheet(entry: SheetEntry, source: SheetJsSheet, formatting: BiffSheet | undefined, formats: readonly CellFormat[], defaultFont: Font): Sheet {
  const digitWidth = maxDigitWidth(defaultFont);
  const defaultRow = pointsToPx(formatting?.defaultRowHeight || defaultRowHeightPt(defaultFont));
  const defaultCol =
    formatting?.standardWidth !== undefined
      ? columnWidthPx(formatting.standardWidth, digitWidth)
      : baseColumnWidthPx(formatting?.defaultColumnChars ?? 8, digitWidth);
  if (entry.kind === 'chartsheet') {
    return emptySheet(entry.name, entry.kind, defaultRow, defaultCol, `“${entry.name}” is a chart sheet; charts are not shown.`);
  }

  const sheet = new SheetBuilder(entry.name, entry.kind, formats);
  const fallbackFormat = DEFAULT_CELL_FORMAT < formats.length ? DEFAULT_CELL_FORMAT : 0;
  const rows = source['!data'] ?? [];
  rows: for (let r = 0; r < Math.min(rows.length, SHEET_LIMITS.maxRows); r++) {
    const row = rows[r];
    if (!row) {
      continue;
    }
    for (let c = 0; c < Math.min(row.length, SHEET_LIMITS.maxCols); c++) {
      const cell = row[c];
      if (!cell) {
        continue;
      }
      const ownFormat = formatting?.formatAt(r, c);
      const style = ownFormat !== undefined && ownFormat < formats.length ? ownFormat : fallbackFormat;
      let kind: CellKind = CellKind.Blank;
      let value = 0;
      switch (cell.t) {
        case 'n':
          if (typeof cell.v === 'number' && Number.isFinite(cell.v)) {
            kind = CellKind.Number;
            value = cell.v;
          }
          break;
        case 's':
          kind = CellKind.SheetString;
          value = sheet.addString(String(cell.v ?? ''));
          break;
        case 'b':
          kind = CellKind.Boolean;
          value = cell.v ? 1 : 0;
          break;
        case 'e':
          kind = CellKind.Error;
          value = sheet.addString(ERRORS[Number(cell.v)] ?? '#VALUE!');
          break;
      }
      const formula = typeof cell.f === 'string' ? sheet.addFormula(cell.f) : NO_FORMULA;
      if (!sheet.add(r, c, kind, value, style, formula)) {
        break rows;
      }
      const link = cell.l?.Target ? hyperlink(cell.l.Target, cell.l.Tooltip, r, c) : undefined;
      if (link) {
        sheet.addLink(link);
      }
    }
  }
  for (const { s, e } of source['!merges'] ?? []) {
    sheet.addMerge({ top: s.r, left: s.c, bottom: e.r, right: e.c });
  }
  if (formatting?.hasDrawings) {
    sheet.noteDrawings();
  }
  return sheet.build({
    rows: { defaultSize: formatting?.rowsHiddenByDefault ? 0 : defaultRow, sizes: rowSizes(source, formatting) },
    cols: { defaultSize: defaultCol, sizes: columnSizes(source, formatting, digitWidth) },
    showGridLines: formatting?.showGridLines ?? true,
    ...(formatting ? { frozen: formatting.frozen } : {}),
    ...(formatting?.topLeft ? { topLeft: formatting.topLeft } : {}),
    ...(formatting?.activeCell ? { activeCell: formatting.activeCell } : {}),
  });
}

function hyperlink(target: string, tooltip: string | undefined, row: number, col: number): Hyperlink | undefined {
  const range = { top: row, left: col, bottom: row, right: col };
  const link: Hyperlink = { range };
  if (target.startsWith('#')) {
    link.location = target.slice(1);
  } else {
    const url = externalUrl(target);
    if (!url) {
      return undefined;
    }
    link.url = url;
  }
  if (tooltip) {
    link.tooltip = tooltip.slice(0, 1024);
  }
  return link;
}

function rowSizes(source: SheetJsSheet, formatting: BiffSheet | undefined): Map<number, number> {
  const sizes = new Map<number, number>();
  if (formatting) {
    for (const [row, points] of formatting.rowHeights) {
      sizes.set(row, pointsToPx(points));
    }
    return sizes;
  }
  (source['!rows'] ?? []).forEach((row, i) => {
    if (row?.hidden) {
      sizes.set(i, 0);
    } else if (row?.hpx) {
      sizes.set(i, row.hpx);
    }
  });
  return sizes;
}

function columnSizes(source: SheetJsSheet, formatting: BiffSheet | undefined, digitWidth: number): Map<number, number> {
  const sizes = new Map<number, number>();
  if (formatting) {
    for (const column of formatting.columns) {
      const size = column.hidden ? 0 : columnWidthPx(column.width, digitWidth);
      for (let col = column.first; col <= Math.min(column.last, SHEET_LIMITS.maxCols - 1); col++) {
        sizes.set(col, size);
      }
    }
    return sizes;
  }
  (source['!cols'] ?? []).forEach((col, i) => {
    if (col?.hidden) {
      sizes.set(i, 0);
    } else if (col?.wpx) {
      sizes.set(i, col.wpx);
    }
  });
  return sizes;
}
