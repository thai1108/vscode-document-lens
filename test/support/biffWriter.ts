// Writes real Excel 97-2003 (.xls) workbooks for tests: a BIFF8 Workbook stream in an OLE compound file (the
// container comes from SheetJS's CFB writer). The record sequence follows what Excel and SheetJS write; only the
// records the viewer reads or Excel needs are produced. Specification: [MS-XLS].
import { CFB } from 'xlsx';

export interface BiffFont {
  name: string;
  /** Points. */
  size: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Palette index; 0x7fff is automatic. */
  color?: number;
}

export type BiffBorder = [style: number, color: number];

export interface BiffXf {
  /** Index into BiffWorkbookSpec.fonts (the writer skips BIFF's missing font 4). */
  font: number;
  /** Built-in format id or the id of a custom BiffWorkbookSpec.formats entry. */
  format?: number;
  /** 0 general, 1 left, 2 center, 3 right. */
  horizontal?: number;
  /** 0 top, 1 center, 2 bottom. */
  vertical?: number;
  wrap?: boolean;
  borders?: { top?: BiffBorder; right?: BiffBorder; bottom?: BiffBorder; left?: BiffBorder };
  /** Fill pattern (1 solid) with foreground and background palette indexes. */
  fill?: { pattern: number; foreground: number; background?: number };
}

export type BiffValue = number | string | boolean | { error: number } | null;

export interface BiffCell {
  row: number;
  col: number;
  /** Index into BiffWorkbookSpec.xfs; cell formats start after the 15 style formats the writer adds. */
  xf: number;
  value: BiffValue;
  /** SUM over a range, the one formula shape the writer encodes: [firstRow, lastRow, col]. */
  sum?: [number, number, number];
}

export interface BiffSheetSpec {
  name: string;
  hidden?: boolean;
  cells: BiffCell[];
  /** [firstRow, lastRow, firstCol, lastCol]. */
  merges?: [number, number, number, number][];
  /** Width in characters (padding included). */
  columns?: { first: number; last: number; width: number; hidden?: boolean }[];
  /** Height in points. */
  rows?: { row: number; height: number; hidden?: boolean }[];
  frozen?: { rows: number; cols: number };
  /** Top-left cell of the view (WINDOW2): with frozen panes, the first frozen row and column. */
  viewTopLeft?: { row: number; col: number };
  links?: { row: number; col: number; url: string }[];
  showGridLines?: boolean;
  /** OBJ records by object type (0x08 picture, 0x19 cell note, ...), without the drawing data Excel adds. */
  objects?: number[];
}

export interface BiffWorkbookSpec {
  fonts: BiffFont[];
  /** Custom number formats: id (164 and up) and format code. */
  formats?: [number, string][];
  xfs: BiffXf[];
  sheets: BiffSheetSpec[];
  date1904?: boolean;
  /** Replaces palette colors from index 8 on (RRGGBB). */
  palette?: string[];
}

/** Index of the first cell format: XF 0-14 are style formats, as in every Excel file. */
export const FIRST_CELL_XF = 15;

class Records {
  private readonly parts: Buffer[] = [];
  length = 0;

  add(type: number, body: Buffer = Buffer.alloc(0)): void {
    const header = Buffer.alloc(4);
    header.writeUInt16LE(type, 0);
    header.writeUInt16LE(body.length, 2);
    this.parts.push(header, body);
    this.length += 4 + body.length;
  }

  bytes(): Buffer {
    return Buffer.concat(this.parts);
  }
}

const u8 = (value: number) => Buffer.from([value & 0xff]);
function u16(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 2);
  values.forEach((value, i) => buffer.writeUInt16LE(value & 0xffff, i * 2));
  return buffer;
}
function u32(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 4);
  values.forEach((value, i) => buffer.writeUInt32LE(value >>> 0, i * 4));
  return buffer;
}
function f64(value: number): Buffer {
  const buffer = Buffer.alloc(8);
  buffer.writeDoubleLE(value, 0);
  return buffer;
}

/** ShortXLUnicodeString (8-bit length) or XLUnicodeString (16-bit length), always stored as UTF-16. */
function unicodeString(text: string, wideLength: boolean): Buffer {
  return Buffer.concat([wideLength ? u16(text.length) : u8(text.length), u8(1), Buffer.from(text, 'utf16le')]);
}

function bof(type: number): Buffer {
  return Buffer.concat([u16(0x0600, type, 0x7262, 0x07cd), u32(0x000100c9, 0x00000706)]);
}

/** Excel's font index skips 4. */
function fontIndex(font: number): number {
  return font < 4 ? font : font + 1;
}

function fontRecord(font: BiffFont): Buffer {
  const flags = (font.italic ? 0x02 : 0) | (font.strike ? 0x08 : 0);
  return Buffer.concat([
    u16(Math.round(font.size * 20), flags, font.color ?? 0x7fff, font.bold ? 700 : 400, 0),
    Buffer.from([font.underline ? 1 : 0, 2, 0, 0]),
    unicodeString(font.name, false),
  ]);
}

function xfRecord(xf: BiffXf, style: boolean): Buffer {
  const borders = xf.borders ?? {};
  const [top = 0, topColor = 64] = borders.top ?? [];
  const [right = 0, rightColor = 64] = borders.right ?? [];
  const [bottom = 0, bottomColor = 64] = borders.bottom ?? [];
  const [left = 0, leftColor = 64] = borders.left ?? [];
  const alignment = (xf.horizontal ?? 0) | (xf.wrap ? 0x08 : 0) | ((xf.vertical ?? 2) << 4);
  // Style XFs have fStyle set and parent 0xFFF; cell XFs use style 0 and say every attribute group is their own.
  const flags = style ? 0xfff5 : 0x0001;
  const used = style ? 0 : 0xfc;
  const borders1 = left | (right << 4) | (top << 8) | (bottom << 12) | (leftColor << 16) | (rightColor << 23);
  const pattern = xf.fill?.pattern ?? 0;
  const borders2 = topColor | (bottomColor << 7) | (pattern << 26);
  const colors = (xf.fill?.foreground ?? 64) | ((xf.fill?.background ?? 65) << 7);
  return Buffer.concat([u16(fontIndex(xf.font), xf.format ?? 0, flags), Buffer.from([alignment, 0, 0, used]), u32(borders1, borders2), u16(colors)]);
}

const HLINK_CLSID = Buffer.from('d0c9ea79f9bace118c8200aa004ba90b', 'hex');
const URL_MONIKER_CLSID = Buffer.from('e0c9ea79f9bace118c8200aa004ba90b', 'hex');

function hyperlinkRecord(row: number, col: number, url: string): Buffer {
  const target = Buffer.from(`${url}\0`, 'utf16le');
  return Buffer.concat([u16(row, row, col, col), HLINK_CLSID, u32(2, 0x03), URL_MONIKER_CLSID, u32(target.length), target]);
}

function cellRecords(records: Records, cell: BiffCell, sst: Map<string, number>): void {
  const head = u16(cell.row, cell.col, FIRST_CELL_XF + cell.xf);
  const value = cell.value;
  if (cell.sum) {
    const [first, last, col] = cell.sum;
    // ptgArea (relative rows and columns) then ptgFuncVar SUM with one argument.
    const area = Buffer.concat([u8(0x25), u16(first, last, col | 0xc000, col | 0xc000)]);
    const sum = Buffer.concat([u8(0x22), u8(1), u16(4)]);
    const rgce = Buffer.concat([area, sum]);
    records.add(0x0006, Buffer.concat([head, f64(typeof value === 'number' ? value : 0), u16(0), u32(0), u16(rgce.length), rgce]));
    return;
  }
  if (value === null) {
    records.add(0x0201, head);
  } else if (typeof value === 'string') {
    records.add(0x00fd, Buffer.concat([head, u32(sst.get(value) ?? 0)]));
  } else if (typeof value === 'boolean') {
    records.add(0x0205, Buffer.concat([head, Buffer.from([value ? 1 : 0, 0])]));
  } else if (typeof value === 'object') {
    records.add(0x0205, Buffer.concat([head, Buffer.from([value.error, 1])]));
  } else if (Number.isInteger(value) && Math.abs(value) < 2 ** 29) {
    // RK: a 30-bit integer shifted left by two, with the integer flag.
    records.add(0x027e, Buffer.concat([head, u32(((value << 2) | 0x02) >>> 0)]));
  } else {
    records.add(0x0203, Buffer.concat([head, f64(value)]));
  }
}

function sheetStream(spec: BiffSheetSpec, sst: Map<string, number>, selected: boolean): Buffer {
  const records = new Records();
  records.add(0x0809, bof(0x0010));
  records.add(0x000d, u16(1));
  records.add(0x000c, u16(100));
  records.add(0x000f, u16(1));
  records.add(0x0011, u16(0));
  records.add(0x0010, f64(0.001));
  records.add(0x005f, u16(1));
  records.add(0x002a, u16(0));
  records.add(0x002b, u16(0));
  records.add(0x0082, u16(1));
  records.add(0x0080, u16(0, 0, 0, 0));
  records.add(0x0225, u16(0, 300));
  records.add(0x0081, u16(0x04c1));
  records.add(0x0083, u16(0));
  records.add(0x0084, u16(0));
  records.add(0x0055, u16(8));
  for (const column of spec.columns ?? []) {
    records.add(0x007d, u16(column.first, column.last, Math.round(column.width * 256), FIRST_CELL_XF, column.hidden ? 0x01 : 0x00, 0));
  }
  const rows = spec.cells.map((cell) => cell.row);
  const cols = spec.cells.map((cell) => cell.col);
  const [firstRow, lastRow] = rows.length ? [Math.min(...rows), Math.max(...rows) + 1] : [0, 0];
  const [firstCol, lastCol] = cols.length ? [Math.min(...cols), Math.max(...cols) + 1] : [0, 0];
  records.add(0x0200, Buffer.concat([u32(firstRow, lastRow), u16(firstCol, lastCol, 0)]));

  // Rows in blocks of 32: the ROW records of a block, then its cells.
  const heights = new Map((spec.rows ?? []).map((row) => [row.row, row]));
  const rowNumbers = [...new Set([...rows, ...heights.keys()])].sort((a, b) => a - b);
  for (let block = 0; block < rowNumbers.length; block += 32) {
    const blockRows = rowNumbers.slice(block, block + 32);
    for (const row of blockRows) {
      const own = spec.cells.filter((cell) => cell.row === row);
      const custom = heights.get(row);
      const flags = 0x0100 | (custom?.hidden ? 0x20 : 0) | (custom ? 0x40 : 0);
      const first = own.length ? Math.min(...own.map((cell) => cell.col)) : 0;
      const last = own.length ? Math.max(...own.map((cell) => cell.col)) + 1 : 0;
      records.add(0x0208, u16(row, first, last, Math.round((custom?.height ?? 15) * 20), 0, 0, flags, FIRST_CELL_XF));
    }
    for (const row of blockRows) {
      for (const cell of spec.cells.filter((candidate) => candidate.row === row).sort((a, b) => a.col - b.col)) {
        cellRecords(records, cell, sst);
      }
    }
  }

  (spec.objects ?? []).forEach((type, i) => {
    // ftCmo (object type and id), then ftEnd.
    records.add(0x005d, Buffer.concat([u16(0x15, 0x12, type, i + 1, 0x6011), Buffer.alloc(12), u16(0, 0)]));
  });
  const frozen = spec.frozen;
  const freeze = Boolean(frozen && (frozen.rows || frozen.cols));
  const window = 0x04b4 | (spec.showGridLines === false ? 0 : 0x02) | (freeze ? 0x0108 : 0) | (selected ? 0x0600 : 0);
  records.add(0x023e, Buffer.concat([u16(window, spec.viewTopLeft?.row ?? 0, spec.viewTopLeft?.col ?? 0), u32(64), u16(0, 0), u32(0)]));
  if (freeze && frozen) {
    const pane = frozen.rows && frozen.cols ? 0 : frozen.rows ? 2 : 1;
    records.add(0x0041, Buffer.concat([u16(frozen.cols, frozen.rows, (spec.viewTopLeft?.row ?? 0) + frozen.rows, (spec.viewTopLeft?.col ?? 0) + frozen.cols), Buffer.from([pane, 0])]));
  }
  records.add(0x001d, Buffer.concat([u8(3), u16(0, 0, 0, 1, 0, 0), Buffer.from([0, 0])]));
  if (spec.merges?.length) {
    records.add(0x00e5, Buffer.concat([u16(spec.merges.length), ...spec.merges.map(([r1, r2, c1, c2]) => u16(r1, r2, c1, c2))]));
  }
  for (const link of spec.links ?? []) {
    records.add(0x01b8, hyperlinkRecord(link.row, link.col, link.url));
  }
  records.add(0x000a);
  return records.bytes();
}

/** The Workbook stream: globals, then one substream per sheet. */
export function biffStream(spec: BiffWorkbookSpec): Buffer {
  const texts = spec.sheets.flatMap((sheet) => sheet.cells.map((cell) => cell.value).filter((value): value is string => typeof value === 'string'));
  const strings = [...new Set(texts)];
  const sst = new Map(strings.map((text, i) => [text, i]));
  const sheets = spec.sheets.map((sheet, i) => sheetStream(sheet, sst, i === 0));

  const head = new Records();
  head.add(0x0809, bof(0x0005));
  head.add(0x00e1, u16(0x04b0));
  head.add(0x00c1, u16(0));
  head.add(0x00e2);
  const user = unicodeString('Document Lens', true);
  head.add(0x005c, Buffer.concat([user, Buffer.alloc(112 - user.length, 0x20)]));
  head.add(0x0042, u16(0x04b0));
  head.add(0x0161, u16(0));
  head.add(0x01c0);
  head.add(0x013d, u16(...spec.sheets.map((_sheet, i) => i + 1)));
  head.add(0x009c, u16(0x11));
  head.add(0x0019, u16(0));
  head.add(0x0012, u16(0));
  head.add(0x0013, u16(0));
  head.add(0x01af, u16(0));
  head.add(0x01bc, u16(0));
  head.add(0x003d, u16(0, 0, 0x7260, 0x44c0, 0x38, 0, 0, 1, 0x01f4));
  head.add(0x0040, u16(0));
  head.add(0x008d, u16(0));
  head.add(0x0022, u16(spec.date1904 ? 1 : 0));
  head.add(0x000e, u16(1));
  head.add(0x01b7, u16(0));
  head.add(0x00da, u16(0));
  for (const font of spec.fonts) {
    head.add(0x0031, fontRecord(font));
  }
  for (const [id, code] of spec.formats ?? []) {
    head.add(0x041e, Buffer.concat([u16(id), unicodeString(code, true)]));
  }
  for (let i = 0; i < FIRST_CELL_XF; i++) {
    head.add(0x00e0, xfRecord({ font: 0 }, true));
  }
  for (const xf of spec.xfs) {
    head.add(0x00e0, xfRecord(xf, false));
  }
  head.add(0x0293, Buffer.concat([u16(0x8000), Buffer.from([0, 0xff])]));
  if (spec.palette) {
    head.add(0x0092, Buffer.concat([u16(spec.palette.length), ...spec.palette.map((hex) => Buffer.from([...Buffer.from(hex, 'hex'), 0]))]));
  }
  head.add(0x0160, u16(0));

  const tail = new Records();
  tail.add(0x008c, u16(1, 81));
  const sstBody = Buffer.concat([u32(texts.length, strings.length), ...strings.map((text) => unicodeString(text, true))]);
  if (sstBody.length > 8224) {
    throw new Error('biffWriter: the shared string table needs CONTINUE records, which this writer does not split');
  }
  tail.add(0x00fc, sstBody);
  tail.add(0x000a);

  const boundSheetSizes = spec.sheets.map((sheet) => 4 + 8 + 2 * sheet.name.length);
  let position = head.length + boundSheetSizes.reduce((sum, size) => sum + size, 0) + tail.length;
  const bound = new Records();
  spec.sheets.forEach((sheet, i) => {
    bound.add(0x0085, Buffer.concat([u32(position), Buffer.from([sheet.hidden ? 1 : 0, 0]), unicodeString(sheet.name, false)]));
    position += sheets[i]?.length ?? 0;
  });
  return Buffer.concat([head.bytes(), bound.bytes(), tail.bytes(), ...sheets]);
}

/** A complete .xls file. */
export function biffWorkbook(spec: BiffWorkbookSpec): Buffer {
  return compoundFile(biffStream(spec));
}

export function compoundFile(workbook: Buffer, streamName = 'Workbook'): Buffer {
  const container = CFB.utils.cfb_new();
  CFB.utils.cfb_add(container, streamName, workbook);
  return Buffer.from(CFB.write(container, { type: 'buffer', fileType: 'cfb' }) as Uint8Array);
}
