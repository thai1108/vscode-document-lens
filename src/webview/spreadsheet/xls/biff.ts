// Formatting of an Excel 97-2003 (BIFF8) workbook stream, which SheetJS reads but does not hand out: fonts, cell
// formats (XF records), the color palette, which format each cell uses, row heights, column widths, gridlines and
// frozen panes. Values, formulas, merges and links still come from SheetJS (see workbook.ts). Only fixed-layout
// record fields are read, every offset is bounds-checked, and the scan always moves forward.
import { SHEET_LIMITS } from '../limits.ts';
import type { CellRef } from '../model/address.ts';
import { defaultPalette, mixColors } from '../model/colors.ts';
import { builtinFormat, formatCode } from '../model/numberFormat.ts';
import { DEFAULT_FONT, defaultFormat, frozenPanes, NO_FROZEN_PANES, type BorderEdge, type BorderStyle, type Borders, type CellFormat, type Font, type FrozenPanes, type HorizontalAlignment, type SheetKind, type VerticalAlignment } from '../model/types.ts';

const RECORD = {
  BOF: 0x0809,
  EOF: 0x000a,
  FILEPASS: 0x002f,
  FONT: 0x0031,
  FORMAT: 0x041e,
  XF: 0x00e0,
  PALETTE: 0x0092,
  WINDOW1: 0x003d,
  BOUNDSHEET: 0x0085,
  NUMBER: 0x0203,
  RK: 0x027e,
  LABELSST: 0x00fd,
  LABEL: 0x0204,
  BLANK: 0x0201,
  BOOLERR: 0x0205,
  FORMULA: 0x0006,
  MULRK: 0x00bd,
  MULBLANK: 0x00be,
  ROW: 0x0208,
  COLINFO: 0x007d,
  DEFCOLWIDTH: 0x0055,
  STANDARDWIDTH: 0x0099,
  DEFAULTROWHEIGHT: 0x0225,
  WINDOW2: 0x023e,
  PANE: 0x0041,
  SELECTION: 0x001d,
  OBJ: 0x005d,
} as const;

const SINGLE_CELL_RECORDS = new Set<number>([RECORD.NUMBER, RECORD.RK, RECORD.LABELSST, RECORD.LABEL, RECORD.BLANK, RECORD.BOOLERR, RECORD.FORMULA]);
const BIFF8 = 0x0600;
const GLOBALS = 0x0005;
const AUTO_COLOR = 0x7fff;

const HORIZONTAL: HorizontalAlignment[] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const VERTICAL: VerticalAlignment[] = ['top', 'center', 'bottom', 'justify', 'distributed'];
const BORDER_STYLES: (BorderStyle | undefined)[] = [
  undefined,
  'thin',
  'medium',
  'dashed',
  'dotted',
  'thick',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
];
/** Share of the foreground color in each BIFF fill pattern (0 none, 1 solid, 2 50% gray, ...). */
const PATTERN_DENSITY = [0, 1, 0.5, 0.75, 0.25, 0.5, 0.5, 0.5, 0.5, 0.5, 0.75, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.125, 0.0625];

export interface ColumnInfo {
  first: number;
  last: number;
  /** Characters of the default font's widest digit, padding included. */
  width: number;
  hidden: boolean;
}

/** A sheet as the workbook lists it (BOUNDSHEET); its own records are read by readSheet() when it is shown. */
export interface BiffSheetEntry {
  name: string;
  hidden: boolean;
  kind: SheetKind;
}

/** The formatting of one sheet. */
export interface BiffSheet {
  /** Format (XF index) of the cell at (row, col), or undefined. */
  formatAt(row: number, col: number): number | undefined;
  /** Calls `visit` for each blank cell (BLANK/MULBLANK) with its format; SheetJS is not asked to make objects of them. */
  forEachBlank(visit: (row: number, col: number, format: number) => void): void;
  /** Row heights in points (0 for hidden rows). */
  rowHeights: Map<number, number>;
  columns: ColumnInfo[];
  /** DEFCOLWIDTH: characters, padding not included. */
  defaultColumnChars?: number;
  /** STANDARDWIDTH: characters, padding included; wins over DEFCOLWIDTH. */
  standardWidth?: number;
  defaultRowHeight?: number;
  rowsHiddenByDefault: boolean;
  showGridLines: boolean;
  frozen: FrozenPanes;
  topLeft?: CellRef;
  activeCell?: CellRef;
  /** Pictures, shapes, controls or charts, none of which the preview shows. */
  hasDrawings: boolean;
}

export interface BiffWorkbook {
  formats: CellFormat[];
  defaultFont: Font;
  sheets: BiffSheetEntry[];
  /** Index of the sheet Excel shows first. */
  activeSheet: number;
  /** Reads the formatting of sheet `index`; nothing is kept between calls. */
  readSheet(index: number): BiffSheet;
}

interface XfRecord {
  font: number;
  numberFormat: number;
  horizontal: number;
  vertical: number;
  wrap: boolean;
  indent: number;
  borderStyles: [number, number, number, number];
  borderColors: [number, number, number, number];
  pattern: number;
  foreground: number;
  background: number;
}

class Reader {
  readonly bytes: Uint8Array;
  readonly view: DataView;

  constructor(bytes: Uint8Array) {
    this.bytes = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  u8(offset: number): number {
    return offset < this.bytes.length ? (this.bytes[offset] ?? 0) : 0;
  }

  u16(offset: number): number {
    return offset + 2 <= this.bytes.length ? this.view.getUint16(offset, true) : 0;
  }

  u32(offset: number): number {
    return offset + 4 <= this.bytes.length ? this.view.getUint32(offset, true) : 0;
  }

  /** Records from `offset` on: [type, data start, data length]; stops at the end of the stream. */
  *records(offset: number): Generator<[number, number, number]> {
    let position = offset;
    while (position + 4 <= this.bytes.length) {
      const type = this.u16(position);
      const length = this.u16(position + 2);
      const start = position + 4;
      if (start + length > this.bytes.length) {
        return;
      }
      yield [type, start, length];
      position = start + length;
    }
  }

  /** An 8-bit or UTF-16 string of `count` characters at `offset` (the flag byte just before it says which). */
  string(offset: number, count: number, wide: boolean, end: number): string {
    const length = Math.max(0, Math.min(count * (wide ? 2 : 1), end - offset));
    const bytes = this.bytes.subarray(offset, offset + length);
    if (wide) {
      return new TextDecoder('utf-16le').decode(bytes);
    }
    // "Compressed" BIFF8 strings hold the low byte of each UTF-16 code unit, i.e. Latin-1.
    let text = '';
    for (const byte of bytes) {
      text += String.fromCharCode(byte);
    }
    return text;
  }
}

/** Reads the formatting of a BIFF8 `Workbook` stream; undefined for older BIFF versions. */
export function readBiffFormatting(stream: Uint8Array): BiffWorkbook | undefined {
  const reader = new Reader(stream);
  if (reader.u16(0) !== RECORD.BOF || reader.u16(4) !== BIFF8 || reader.u16(6) !== GLOBALS) {
    return undefined;
  }
  const fonts: Font[] = [];
  const fontColors: number[] = [];
  const numberFormats = new Map<number, string>();
  const xfs: XfRecord[] = [];
  const palette = defaultPalette();
  // By stream position: several BOUNDSHEET records pointing at the same sheet data are one sheet (the last one's name
  // wins, as in SheetJS), or a crafted file could make the same records be read again for every entry.
  const boundSheets = new Map<number, BiffSheetEntry & { position: number }>();
  let activeSheet = 0;

  for (const [type, start, length] of reader.records(0)) {
    const end = start + length;
    if (type === RECORD.EOF) {
      break;
    }
    switch (type) {
      case RECORD.FILEPASS:
        throw new Error('password-protected');
      case RECORD.WINDOW1:
        if (length >= 12) {
          activeSheet = reader.u16(start + 10);
        }
        break;
      case RECORD.FONT:
        if (length >= 16 && fonts.length < SHEET_LIMITS.maxStyles) {
          const flags = reader.u16(start + 2);
          const underline = reader.u8(start + 10);
          const script = reader.u16(start + 8);
          const font: Font = {
            name: reader.string(start + 16, reader.u8(start + 14), (reader.u8(start + 15) & 1) === 1, end).slice(0, 64) || DEFAULT_FONT.name,
            size: Math.min(409, Math.max(1, reader.u16(start) / 20)),
            bold: reader.u16(start + 6) >= 600,
            italic: (flags & 0x0002) !== 0,
            underline: underline === 0 ? 'none' : underline === 2 || underline === 0x22 ? 'double' : 'single',
            strike: (flags & 0x0008) !== 0,
          };
          if (script === 1 || script === 2) {
            font.vertAlign = script === 1 ? 'superscript' : 'subscript';
          }
          fonts.push(font);
          fontColors.push(reader.u16(start + 4));
        }
        break;
      case RECORD.FORMAT:
        if (length >= 5) {
          numberFormats.set(reader.u16(start), formatCode(reader.string(start + 5, reader.u16(start + 2), (reader.u8(start + 4) & 1) === 1, end)));
        }
        break;
      case RECORD.XF:
        if (length >= 20 && xfs.length < SHEET_LIMITS.maxStyles) {
          const alignment = reader.u8(start + 6);
          const borders1 = reader.u32(start + 10);
          const borders2 = reader.u32(start + 14);
          const colors = reader.u16(start + 18);
          xfs.push({
            font: reader.u16(start),
            numberFormat: reader.u16(start + 2),
            horizontal: alignment & 0x07,
            vertical: (alignment >> 4) & 0x07,
            wrap: (alignment & 0x08) !== 0,
            indent: reader.u8(start + 8) & 0x0f,
            borderStyles: [(borders1 >> 8) & 0x0f, (borders1 >> 4) & 0x0f, (borders1 >> 12) & 0x0f, borders1 & 0x0f],
            borderColors: [borders2 & 0x7f, (borders1 >>> 23) & 0x7f, (borders2 >>> 7) & 0x7f, (borders1 >>> 16) & 0x7f],
            pattern: (borders2 >>> 26) & 0x3f,
            foreground: colors & 0x7f,
            background: (colors >> 7) & 0x7f,
          });
        }
        break;
      case RECORD.PALETTE: {
        const count = Math.min(56, reader.u16(start), Math.floor((length - 2) / 4));
        for (let i = 0; i < count; i++) {
          const at = start + 2 + i * 4;
          palette[8 + i] = `#${[reader.u8(at), reader.u8(at + 1), reader.u8(at + 2)].map((value) => value.toString(16).padStart(2, '0')).join('')}`;
        }
        break;
      }
      case RECORD.BOUNDSHEET: {
        const position = reader.u32(start);
        if (length >= 8 && (boundSheets.size < SHEET_LIMITS.maxSheets || boundSheets.has(position))) {
          const kind = reader.u8(start + 5);
          boundSheets.set(position, {
            position,
            hidden: (reader.u8(start + 4) & 0x03) !== 0,
            kind: kind === 0 ? 'worksheet' : kind === 2 ? 'chartsheet' : 'macrosheet',
            name: reader.string(start + 8, reader.u8(start + 6), (reader.u8(start + 7) & 1) === 1, end),
          });
        }
        break;
      }
    }
  }

  const color = (index: number): string | undefined => (index === AUTO_COLOR || index >= 64 ? undefined : palette[index]);
  const resolvedFonts = fonts.map((font, i) => {
    const fontColor = color(fontColors[i] ?? AUTO_COLOR);
    return fontColor && fontColor !== '#000000' ? { ...font, color: fontColor } : font;
  });
  const defaultFont = resolvedFonts[0] ?? DEFAULT_FONT;
  // Font index 4 does not exist in BIFF: the fifth FONT record is index 5.
  const fontAt = (index: number) => resolvedFonts[index < 4 ? index : index - 1] ?? defaultFont;
  const formats = xfs.map((xf): CellFormat => {
    const borders: Borders = {};
    (['top', 'right', 'bottom', 'left'] as const).forEach((side, i) => {
      const style = BORDER_STYLES[xf.borderStyles[i] ?? 0];
      if (style) {
        const edge: BorderEdge = { style, color: color(xf.borderColors[i] ?? 64) ?? '#000000' };
        borders[side] = edge;
      }
    });
    return {
      font: fontAt(xf.font),
      fill: fillColor(xf, color),
      borders,
      horizontal: HORIZONTAL[xf.horizontal] ?? 'general',
      vertical: VERTICAL[xf.vertical] ?? 'bottom',
      wrap: xf.wrap,
      indent: xf.indent,
      numberFormat: numberFormats.get(xf.numberFormat) ?? builtinFormat(xf.numberFormat),
    };
  });
  if (!formats.length) {
    formats.push(defaultFormat(defaultFont));
  }
  const entries = [...boundSheets.values()];
  return {
    formats,
    defaultFont,
    sheets: entries.map(({ name, hidden, kind }) => ({ name, hidden, kind })),
    activeSheet,
    readSheet: (index) => readSheetStream(reader, entries[index]?.position ?? -1),
  };
}

function fillColor(xf: XfRecord, color: (index: number) => string | undefined): string | undefined {
  const density = PATTERN_DENSITY[xf.pattern] ?? 0;
  if (!density) {
    return undefined;
  }
  const foreground = color(xf.foreground) ?? '#000000';
  return density === 1 ? foreground : mixColors(foreground, color(xf.background) ?? '#ffffff', density);
}

function readSheetStream(reader: Reader, position: number): BiffSheet {
  const cells = new CellFormats();
  const rowHeights = new Map<number, number>();
  const columns: ColumnInfo[] = [];
  const sheet: BiffSheet = {
    formatAt: () => undefined,
    forEachBlank: () => undefined,
    rowHeights,
    columns,
    rowsHiddenByDefault: false,
    showGridLines: true,
    frozen: NO_FROZEN_PANES,
    hasDrawings: false,
  };
  if (position < 0 || reader.u16(position) !== RECORD.BOF) {
    return sheet;
  }
  let frozen = false;
  let viewTopLeft: CellRef | undefined;
  let depth = 0;
  for (const [type, start, length] of reader.records(position)) {
    if (type === RECORD.BOF) {
      depth++;
      // A substream inside the sheet's is an embedded chart.
      sheet.hasDrawings ||= depth > 1;
      continue;
    }
    if (type === RECORD.EOF) {
      depth--;
      if (depth <= 0) {
        break;
      }
      continue;
    }
    // Records of embedded charts (their own BOF/EOF inside the sheet) are not the sheet's.
    if (depth !== 1) {
      continue;
    }
    if (SINGLE_CELL_RECORDS.has(type)) {
      if (length >= 6) {
        cells.add(reader.u16(start), reader.u16(start + 2), reader.u16(start + 4), type === RECORD.BLANK);
      }
      continue;
    }
    switch (type) {
      case RECORD.OBJ:
        // The common object data (ftCmo) comes first; type 0x19 is a cell note, the others are drawings and controls.
        if (length >= 6 && reader.u16(start) === 0x15 && reader.u16(start + 4) !== 0x19) {
          sheet.hasDrawings = true;
        }
        break;
      case RECORD.MULRK:
      case RECORD.MULBLANK: {
        const step = type === RECORD.MULRK ? 6 : 2;
        const row = reader.u16(start);
        const first = reader.u16(start + 2);
        const count = Math.floor((length - 6) / step);
        for (let i = 0; i < count; i++) {
          cells.add(row, first + i, reader.u16(start + 4 + i * step), type === RECORD.MULBLANK);
        }
        break;
      }
      case RECORD.ROW:
        if (length >= 16 && rowHeights.size < SHEET_LIMITS.maxRowEntries) {
          const flags = reader.u16(start + 12);
          const height = (reader.u16(start + 6) & 0x7fff) / 20;
          if (flags & 0x20) {
            rowHeights.set(reader.u16(start), 0);
          } else if (height > 0) {
            rowHeights.set(reader.u16(start), height);
          }
        }
        break;
      case RECORD.COLINFO:
        if (length >= 10 && columns.length < SHEET_LIMITS.maxColumnRanges) {
          columns.push({ first: reader.u16(start), last: reader.u16(start + 2), width: reader.u16(start + 4) / 256, hidden: (reader.u16(start + 8) & 0x01) !== 0 });
        }
        break;
      case RECORD.DEFCOLWIDTH:
        if (length >= 2) {
          sheet.defaultColumnChars = reader.u16(start);
        }
        break;
      case RECORD.STANDARDWIDTH:
        if (length >= 2) {
          sheet.standardWidth = reader.u16(start) / 256;
        }
        break;
      case RECORD.DEFAULTROWHEIGHT:
        if (length >= 4) {
          sheet.rowsHiddenByDefault = (reader.u16(start) & 0x02) !== 0;
          sheet.defaultRowHeight = reader.u16(start + 2) / 20;
        }
        break;
      case RECORD.WINDOW2:
        if (length >= 6) {
          const flags = reader.u16(start);
          sheet.showGridLines = (flags & 0x0002) !== 0;
          frozen = (flags & 0x0008) !== 0;
          viewTopLeft = { row: reader.u16(start + 2), col: reader.u16(start + 4) };
          if (viewTopLeft.row || viewTopLeft.col) {
            sheet.topLeft = viewTopLeft;
          }
        }
        break;
      case RECORD.PANE:
        // Frozen rows and columns are counted from the view's top-left cell; the pane's is where scrolling starts.
        if (length >= 8 && frozen) {
          sheet.frozen = frozenPanes(reader.u16(start + 2), reader.u16(start), viewTopLeft, SHEET_LIMITS.maxRows, SHEET_LIMITS.maxCols);
          sheet.topLeft = { row: reader.u16(start + 4), col: reader.u16(start + 6) };
        }
        break;
      case RECORD.SELECTION:
        if (length >= 5 && !sheet.activeCell) {
          sheet.activeCell = { row: reader.u16(start + 1), col: reader.u16(start + 3) };
        }
        break;
    }
  }
  const lookup = cells.build();
  sheet.formatAt = lookup.formatAt;
  sheet.forEachBlank = lookup.forEachBlank;
  return sheet;
}

/** (row, col) → XF index and whether the cell is blank, collected in record order and looked up by binary search. */
class CellFormats {
  private keys: number[] = [];
  private formats: number[] = [];
  private blanks: boolean[] = [];
  private sorted = true;

  add(row: number, col: number, format: number, blank: boolean): void {
    if (this.keys.length >= SHEET_LIMITS.maxCells) {
      return;
    }
    const key = row * SHEET_LIMITS.maxCols + col;
    const last = this.keys[this.keys.length - 1];
    if (last !== undefined && key <= last) {
      this.sorted = false;
    }
    this.keys.push(key);
    this.formats.push(format);
    this.blanks.push(blank);
  }

  build(): Pick<BiffSheet, 'formatAt' | 'forEachBlank'> {
    let keys = Float64Array.from(this.keys);
    let formats = Uint16Array.from(this.formats);
    let blanks = Uint8Array.from(this.blanks, Number);
    if (!this.sorted) {
      const order = Array.from(keys.keys()).sort((a, b) => (keys[a] ?? 0) - (keys[b] ?? 0) || a - b);
      keys = Float64Array.from(order, (i) => keys[i] ?? 0);
      formats = Uint16Array.from(order, (i) => formats[i] ?? 0);
      blanks = Uint8Array.from(order, (i) => blanks[i] ?? 0);
    }
    this.keys = [];
    this.formats = [];
    this.blanks = [];
    return {
      formatAt(row, col) {
        const key = row * SHEET_LIMITS.maxCols + col;
        let low = 0;
        let high = keys.length;
        // The last of equal keys wins, as when a record is written twice.
        while (low < high) {
          const mid = (low + high) >>> 1;
          if ((keys[mid] ?? 0) <= key) {
            low = mid + 1;
          } else {
            high = mid;
          }
        }
        return low > 0 && keys[low - 1] === key ? formats[low - 1] : undefined;
      },
      forEachBlank(visit) {
        for (let i = 0; i < keys.length; i++) {
          const key = keys[i] ?? 0;
          // Only the last record for a cell counts.
          if (blanks[i] && keys[i + 1] !== key) {
            visit(Math.floor(key / SHEET_LIMITS.maxCols), key % SHEET_LIMITS.maxCols, formats[i] ?? 0);
          }
        }
      },
    };
  }
}
