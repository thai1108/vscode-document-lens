// The record layer shared by the PowerPoint binary format ([MS-PPT]) and the drawings inside it ([MS-ODRAW]): every
// record starts with an 8-byte header (version and instance packed in 16 bits, type, data length). Lengths come from
// the file, so each record must end inside its container and the number of records read is capped.
import { damaged, PptError } from './errors.ts';

/** Record types used by the viewer. */
export const RT = {
  Document: 0x03e8,
  DocumentAtom: 0x03e9,
  Slide: 0x03ee,
  SlidePersistAtom: 0x03f3,
  SlideShowSlideInfoAtom: 0x03f9,
  DrawingGroup: 0x040b,
  Drawing: 0x040c,
  PlaceholderAtom: 0x0bc3,
  OutlineTextRefAtom: 0x0f9e,
  TextHeaderAtom: 0x0f9f,
  TextCharsAtom: 0x0fa0,
  StyleTextPropAtom: 0x0fa1,
  TextBytesAtom: 0x0fa8,
  SlideListWithText: 0x0ff0,
  UserEditAtom: 0x0ff5,
  CurrentUserAtom: 0x0ff6,
  PersistDirectoryAtom: 0x1772,
  CryptSession10Container: 0x2f14,
  // [MS-ODRAW]
  DggContainer: 0xf000,
  BStoreContainer: 0xf001,
  DgContainer: 0xf002,
  SpgrContainer: 0xf003,
  SpContainer: 0xf004,
  FBSE: 0xf007,
  FSP: 0xf00a,
  FOPT: 0xf00b,
  ClientTextbox: 0xf00d,
  ClientAnchor: 0xf010,
  ClientData: 0xf011,
  BlipFirst: 0xf018,
  BlipEmf: 0xf01a,
  BlipWmf: 0xf01b,
  BlipPict: 0xf01c,
  BlipJpeg: 0xf01d,
  BlipPng: 0xf01e,
  BlipDib: 0xf01f,
  BlipTiff: 0xf029,
  BlipJpegCmyk: 0xf02a,
  BlipLast: 0xf117,
  TertiaryFOPT: 0xf122,
} as const;

export interface PptRecord {
  readonly type: number;
  readonly instance: number;
  readonly version: number;
  /** Offset of the record header in its stream. */
  readonly offset: number;
  /** The record's data is [start, end) in its stream. */
  readonly start: number;
  readonly end: number;
}

const HEADER_SIZE = 8;

/**
 * Reads records from one stream, counting them against a budget shared by everything parsed from the file. Other
 * per-item loops over file data (e.g. shape property tables) spend from the same budget.
 */
export class RecordStream {
  readonly bytes: Uint8Array;
  private readonly view: DataView;
  private readonly budget: { records: number };

  constructor(bytes: Uint8Array, budget: { records: number }) {
    this.bytes = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.budget = budget;
  }

  /** The record whose header starts at `offset`; it must end by `limit` (its container's end). */
  at(offset: number, limit = this.bytes.byteLength): PptRecord {
    if (!Number.isInteger(offset) || offset < 0 || offset + HEADER_SIZE > Math.min(limit, this.bytes.byteLength)) {
      throw damaged('a record lies outside its stream');
    }
    this.spend(1);
    const versionAndInstance = this.view.getUint16(offset, true);
    const start = offset + HEADER_SIZE;
    const length = this.view.getUint32(offset + 4, true);
    if (length > Math.min(limit, this.bytes.byteLength) - start) {
      throw damaged('a record runs past the end of its container');
    }
    return {
      version: versionAndInstance & 0xf,
      instance: versionAndInstance >> 4,
      type: this.view.getUint16(offset + 2, true),
      offset,
      start,
      end: start + length,
    };
  }

  /** Counts `units` of work against the budget; a file that exhausts it is refused. */
  spend(units: number): void {
    this.budget.records -= units;
    if (this.budget.records < 0) {
      throw new PptError('too-complex', 'This .ppt file has too many records to preview');
    }
  }

  /** The records inside a container, in order; trailing bytes too short for a header are ignored. */
  children(parent: Pick<PptRecord, 'start' | 'end'>): PptRecord[] {
    const children: PptRecord[] = [];
    for (let offset = parent.start; offset + HEADER_SIZE <= parent.end; ) {
      const child = this.at(offset, parent.end);
      children.push(child);
      offset = child.end;
    }
    return children;
  }

  /** The first child of `parent` with the given type (and instance, when given). */
  child(parent: Pick<PptRecord, 'start' | 'end'>, type: number, instance?: number): PptRecord | undefined {
    return this.children(parent).find((child) => child.type === type && (instance === undefined || child.instance === instance));
  }

  /** Checks that an atom holds at least `length` bytes before its fields are read. */
  expect(record: PptRecord, length: number, what: string): void {
    if (record.end - record.start < length) {
      throw damaged(`${what} is too short`);
    }
  }

  u8(offset: number): number {
    return this.view.getUint8(offset);
  }

  u16(offset: number): number {
    return this.view.getUint16(offset, true);
  }

  u32(offset: number): number {
    return this.view.getUint32(offset, true);
  }

  i16(offset: number): number {
    return this.view.getInt16(offset, true);
  }

  i32(offset: number): number {
    return this.view.getInt32(offset, true);
  }

  data(record: Pick<PptRecord, 'start' | 'end'>): Uint8Array {
    return this.bytes.subarray(record.start, record.end);
  }
}
