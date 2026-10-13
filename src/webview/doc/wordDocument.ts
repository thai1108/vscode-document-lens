// The text of a Word 97–2003 binary document ([MS-DOC]): the File Information Block (FIB) at the start of the
// WordDocument stream says which table stream holds the piece table (CLX), and the pieces map character positions (CPs)
// to cp1252 or UTF-16LE text in the WordDocument stream. Paragraph properties are read only as far as tables need them.
import { CompoundFile } from './compoundFile.ts';

export const DOC_LIMITS = {
  /** Pieces in the piece table; real documents have a few thousand at most. */
  maxPieces: 1_000_000,
  /** Characters decoded across all stories; the rest of a longer document is cut with a notice. */
  maxCharacters: 3_000_000,
};

const WORD_97_IDENT = 0xa5ec;
const OLDER_IDENTS = new Set([0xa5dc, 0xa59b]);
/** Word 97 writes nFib 0xC1; Word 6.0/95 files have nFib 101–104 and a different layout. */
const MIN_NFIB = 0xc0;
const FLAG_ENCRYPTED = 0x0100;
const FLAG_WHICH_TABLE_STREAM = 0x0200;
const FKP_SIZE = 512;
const ENCRYPTED = 'This .doc file is password-protected; Document Lens cannot open encrypted files';

// Indexes of fc/lcb pairs in FibRgFcLcb97.
const PAIR_FOOTNOTE_REFS = 2;
const PAIR_PAPX_BIN_TABLE = 13;
const PAIR_CLX = 33;
const PAIR_ENDNOTE_REFS = 46;

// Paragraph sprms that mark table paragraphs ([MS-DOC] 2.6.2).
const SPRM_IN_TABLE = 0x2416;
const SPRM_ROW_END = 0x2417;
const SPRM_INNER_ROW_END = 0x244c;
const SPRM_TABLE_DEPTH = 0x6649;
const SPRM_DEF_TABLE = 0xd608;
const SPRM_CHANGE_TABS = 0xc615;

function corrupt(detail: string): Error {
  return new Error(`Corrupt .doc file: ${detail}`);
}

export interface Piece {
  cpStart: number;
  cpEnd: number;
  /** Byte offset of the piece's first character in the WordDocument stream. */
  offset: number;
  /** One cp1252 byte per character instead of UTF-16LE. */
  compressed: boolean;
}

export interface ParagraphProperties {
  inTable: boolean;
  /** The 0x07 mark that ends a table row rather than a cell. */
  rowEnd: boolean;
  /** The paragraph mark that ends a row of a nested table. */
  innerRowEnd: boolean;
}

export type StoryKind = 'main' | 'footnotes' | 'headers' | 'comments' | 'endnotes' | 'textBoxes' | 'headerTextBoxes';

export interface StoryRange {
  kind: StoryKind;
  start: number;
  end: number;
}

export interface WordDocument {
  /** The characters from CP 0 on (string index = CP), cut at DOC_LIMITS.maxCharacters. */
  text: string;
  truncated: boolean;
  stories: StoryRange[];
  /** CPs of the footnote and endnote reference marks in the main text. */
  footnoteRefs: readonly number[];
  endnoteRefs: readonly number[];
  /** Properties of the paragraph whose mark is at `cp`; undefined when the file has no usable paragraph table. */
  paragraphAt(cp: number): ParagraphProperties | undefined;
}

// Story lengths in FibRgLw97, in CP order; index 7 (ccpMcr) is unused and always 0.
const STORY_LENGTHS: readonly [number, StoryKind | undefined][] = [
  [3, 'main'],
  [4, 'footnotes'],
  [5, 'headers'],
  [6, undefined],
  [7, 'comments'],
  [8, 'endnotes'],
  [9, 'textBoxes'],
  [10, 'headerTextBoxes'],
];

export function readWordDocument(bytes: Uint8Array, limits = DOC_LIMITS): WordDocument {
  const file = CompoundFile.open(bytes);
  if (!file.has('WordDocument')) {
    throw notWord(file);
  }
  const word = file.readStream('WordDocument');
  const fib = readFib(word);
  const table = file.readStream(fib.tableStream);
  const pieces = readPieces(fcLcb(table, fib.pairs, PAIR_CLX, 'the piece table'), limits);

  let storyEnd = 0;
  const stories: StoryRange[] = [];
  for (const [index, kind] of STORY_LENGTHS) {
    const length = fib.lengths[index] ?? 0;
    if (kind && length > 0) {
      stories.push({ kind, start: storyEnd, end: storyEnd + length });
    }
    storyEnd += length;
  }
  const available = pieces.at(-1)?.cpEnd ?? 0;
  const wanted = Math.min(storyEnd, available);
  const decoded = Math.min(wanted, limits.maxCharacters);
  const text = decodeText(word, pieces, decoded);
  const papx = new ParagraphTable(word, pieces, optionalFcLcb(table, fib.pairs, PAIR_PAPX_BIN_TABLE));

  return {
    text,
    truncated: decoded < wanted,
    stories: stories.map((story) => ({ ...story, end: Math.min(story.end, decoded) })).filter((story) => story.end > story.start),
    footnoteRefs: noteRefs(optionalFcLcb(table, fib.pairs, PAIR_FOOTNOTE_REFS)),
    endnoteRefs: noteRefs(optionalFcLcb(table, fib.pairs, PAIR_ENDNOTE_REFS)),
    paragraphAt: (cp) => papx.at(cp),
  };
}

function notWord(file: CompoundFile): Error {
  if (file.has('EncryptedPackage')) {
    return new Error(ENCRYPTED);
  }
  if (file.has('Workbook') || file.has('Book')) {
    return new Error('This .doc file is actually an Excel workbook; rename it to .xls to open it');
  }
  if (file.has('PowerPoint Document')) {
    return new Error('This .doc file is actually a PowerPoint presentation; rename it to .ppt to open it');
  }
  return new Error('Not a Word document (the file has no WordDocument stream)');
}

interface Fib {
  tableStream: '0Table' | '1Table';
  /** FibRgLw97 values (story lengths at 3–10). */
  lengths: number[];
  /** FibRgFcLcb97 as [fc, lcb] pairs. */
  pairs: [number, number][];
}

function readFib(word: Uint8Array): Fib {
  const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
  if (word.byteLength < 0x22) {
    throw corrupt('the file information block is truncated');
  }
  const ident = view.getUint16(0, true);
  const nFib = view.getUint16(2, true);
  if (ident !== WORD_97_IDENT && !OLDER_IDENTS.has(ident)) {
    throw new Error('Not a Word document (the WordDocument stream has no Word header)');
  }
  if (ident !== WORD_97_IDENT || nFib < MIN_NFIB) {
    throw new Error('This .doc file was saved by Word 6.0/95 or earlier; Document Lens reads Word 97–2003 documents only');
  }
  const flags = view.getUint16(0x0a, true);
  if (flags & FLAG_ENCRYPTED) {
    throw new Error(ENCRYPTED);
  }

  // FibBase is followed by three counted arrays: fibRgW (16-bit), fibRgLw (32-bit) and fibRgFcLcb (fc/lcb pairs).
  let offset = 0x20;
  const count = (size: number) => {
    if (offset + 2 > word.byteLength) {
      throw corrupt('the file information block is truncated');
    }
    const n = view.getUint16(offset, true);
    const start = offset + 2;
    offset = start + n * size;
    if (offset > word.byteLength) {
      throw corrupt('the file information block is truncated');
    }
    return { n, start };
  };
  count(2);
  const rgLw = count(4);
  const rgFcLcb = count(8);
  // FibRgLw97 must reach ccpText (index 3) and FibRgFcLcb97 the CLX; Word 97 writes 22 and 93 entries.
  if (rgLw.n < 4 || rgFcLcb.n <= PAIR_CLX) {
    throw corrupt('the file information block is too short');
  }
  const lengths: number[] = [];
  for (let i = 0; i < rgLw.n; i++) {
    lengths.push(view.getInt32(rgLw.start + i * 4, true));
  }
  if (STORY_LENGTHS.some(([index]) => (lengths[index] ?? 0) < 0)) {
    throw corrupt('a story has a negative length');
  }
  const pairs: [number, number][] = [];
  for (let i = 0; i < rgFcLcb.n; i++) {
    pairs.push([view.getUint32(rgFcLcb.start + i * 8, true), view.getUint32(rgFcLcb.start + i * 8 + 4, true)]);
  }
  return { tableStream: flags & FLAG_WHICH_TABLE_STREAM ? '1Table' : '0Table', lengths, pairs };
}

/** The structure at fc/lcb pair `index` in the table stream; required, so a bad location is corruption. */
function fcLcb(table: Uint8Array, pairs: [number, number][], index: number, what: string): Uint8Array {
  const [fc, lcb] = pairs[index] ?? [0, 0];
  if (lcb === 0 || fc + lcb > table.byteLength) {
    throw corrupt(`${what} is ${lcb === 0 ? 'missing' : 'outside the table stream'}`);
  }
  return table.subarray(fc, fc + lcb);
}

/** An optional structure: absent or misplaced ones are skipped rather than failing the document. */
function optionalFcLcb(table: Uint8Array, pairs: [number, number][], index: number): Uint8Array | undefined {
  const [fc, lcb] = pairs[index] ?? [0, 0];
  return lcb > 0 && fc + lcb <= table.byteLength ? table.subarray(fc, fc + lcb) : undefined;
}

/** The PlcPcd inside the CLX, after any property lists (Prc) that precede it. */
function readPieces(clx: Uint8Array, limits: typeof DOC_LIMITS): Piece[] {
  const view = new DataView(clx.buffer, clx.byteOffset, clx.byteLength);
  let offset = 0;
  while (offset < clx.byteLength && clx[offset] === 0x01) {
    if (offset + 3 > clx.byteLength) {
      throw corrupt('the piece table is truncated');
    }
    offset += 3 + Math.max(0, view.getInt16(offset + 1, true));
  }
  if (clx[offset] !== 0x02 || offset + 5 > clx.byteLength) {
    throw corrupt('the piece table is missing');
  }
  const lcb = view.getUint32(offset + 1, true);
  const start = offset + 5;
  if (start + lcb > clx.byteLength || lcb < 4 || (lcb - 4) % 12 !== 0) {
    throw corrupt('the piece table has an invalid size');
  }
  const count = (lcb - 4) / 12;
  if (count > limits.maxPieces) {
    throw new Error(`This .doc file has too many text pieces to preview (${count}; limit ${limits.maxPieces})`);
  }
  const cp = (i: number) => view.getUint32(start + i * 4, true);
  if (cp(0) !== 0) {
    throw corrupt('the piece table does not start at the first character');
  }
  const pieces: Piece[] = [];
  for (let i = 0; i < count; i++) {
    const cpStart = cp(i);
    const cpEnd = cp(i + 1);
    if (cpEnd < cpStart || cpEnd > 0x7fffffff) {
      throw corrupt('the piece table is out of order');
    }
    const fc = view.getUint32(start + (count + 1) * 4 + i * 8 + 2, true);
    const compressed = (fc & 0x40000000) !== 0;
    const position = fc & 0x3fffffff;
    pieces.push({ cpStart, cpEnd, offset: compressed ? Math.floor(position / 2) : position, compressed });
  }
  return pieces;
}

const cp1252 = new TextDecoder('windows-1252');
const utf16 = new TextDecoder('utf-16le', { ignoreBOM: true });

/** Characters 0 to `end` (string index = CP: both decoders keep one code unit per character). */
function decodeText(word: Uint8Array, pieces: Piece[], end: number): string {
  const chunks: string[] = [];
  for (const piece of pieces) {
    if (piece.cpStart >= end) {
      break;
    }
    const count = Math.min(piece.cpEnd, end) - piece.cpStart;
    const length = piece.compressed ? count : count * 2;
    if (piece.offset + length > word.byteLength) {
      throw corrupt('a text piece lies outside the WordDocument stream');
    }
    const bytes = word.subarray(piece.offset, piece.offset + length);
    chunks.push((piece.compressed ? cp1252 : utf16).decode(bytes));
  }
  return chunks.join('');
}

/** CPs from a PlcfFndRef / PlcfEndRef (n + 1 CPs, then n 2-byte FRDs). */
function noteRefs(plc: Uint8Array | undefined): number[] {
  if (!plc || plc.byteLength < 10 || (plc.byteLength - 4) % 6 !== 0) {
    return [];
  }
  const view = new DataView(plc.buffer, plc.byteOffset, plc.byteLength);
  const refs: number[] = [];
  for (let i = 0; i < (plc.byteLength - 4) / 6; i++) {
    refs.push(view.getUint32(i * 4, true));
  }
  return refs;
}

/**
 * Paragraph properties from the PAPX bin table (PlcBtePapx): its entries point at 512-byte FKP pages in the
 * WordDocument stream, which map stream offsets of paragraph marks to property lists. Pages are parsed on first use.
 */
class ParagraphTable {
  private readonly word: Uint8Array;
  private readonly pieces: Piece[];
  private readonly boundaries: Uint32Array;
  private readonly pages: Uint32Array;
  private readonly cache = new Map<number, FkpPage | null>();

  constructor(word: Uint8Array, pieces: Piece[], plc: Uint8Array | undefined) {
    this.word = word;
    this.pieces = pieces;
    const count = plc && plc.byteLength >= 12 && (plc.byteLength - 4) % 8 === 0 ? (plc.byteLength - 4) / 8 : 0;
    this.boundaries = new Uint32Array(count + 1);
    this.pages = new Uint32Array(count);
    if (plc && count > 0) {
      const view = new DataView(plc.buffer, plc.byteOffset, plc.byteLength);
      for (let i = 0; i <= count; i++) {
        this.boundaries[i] = view.getUint32(i * 4, true);
      }
      for (let i = 0; i < count; i++) {
        this.pages[i] = view.getUint32((count + 1) * 4 + i * 4, true) & 0x3fffff;
      }
    }
  }

  at(cp: number): ParagraphProperties | undefined {
    if (this.pages.length === 0) {
      return undefined;
    }
    const piece = this.pieces[lastAtOrBefore(this.pieces.length, (i) => this.pieces[i]?.cpStart ?? 0, cp)];
    if (!piece || cp >= piece.cpEnd) {
      return undefined;
    }
    const fc = piece.offset + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
    const index = lastAtOrBefore(this.pages.length, (i) => this.boundaries[i] ?? 0, fc);
    const pn = this.pages[index];
    if (pn === undefined || fc >= (this.boundaries[index + 1] ?? 0)) {
      return undefined;
    }
    let page = this.cache.get(pn);
    if (page === undefined) {
      page = FkpPage.parse(this.word, pn);
      this.cache.set(pn, page);
    }
    return page?.at(fc);
  }
}

const DEFAULT_PARAGRAPH: ParagraphProperties = { inTable: false, rowEnd: false, innerRowEnd: false };

/** One PapxFkp: crun + 1 stream offsets, then crun BxPap entries whose bOffset * 2 locates each PapxInFkp. */
class FkpPage {
  private readonly offsets: number[];
  private readonly properties: ParagraphProperties[];

  private constructor(offsets: number[], properties: ParagraphProperties[]) {
    this.offsets = offsets;
    this.properties = properties;
  }

  static parse(word: Uint8Array, pn: number): FkpPage | null {
    const start = pn * FKP_SIZE;
    if (start + FKP_SIZE > word.byteLength) {
      return null;
    }
    const page = word.subarray(start, start + FKP_SIZE);
    const view = new DataView(page.buffer, page.byteOffset, FKP_SIZE);
    const crun = page[FKP_SIZE - 1] ?? 0;
    if (crun === 0 || (crun + 1) * 4 + crun * 13 > FKP_SIZE - 1) {
      return null;
    }
    const offsets: number[] = [];
    for (let i = 0; i <= crun; i++) {
      offsets.push(view.getUint32(i * 4, true));
    }
    const properties: ParagraphProperties[] = [];
    for (let i = 0; i < crun; i++) {
      const bOffset = page[(crun + 1) * 4 + i * 13] ?? 0;
      properties.push(bOffset === 0 ? DEFAULT_PARAGRAPH : papxProperties(page, bOffset * 2));
    }
    return new FkpPage(offsets, properties);
  }

  at(fc: number): ParagraphProperties | undefined {
    const index = lastAtOrBefore(this.properties.length, (i) => this.offsets[i] ?? 0, fc);
    return index >= 0 && fc < (this.offsets[index + 1] ?? 0) ? this.properties[index] : undefined;
  }
}

/** Reads the table flags from a PapxInFkp: cb, then a GrpPrlAndIstd of 2 * cb - 1 bytes (or cb' and 2 * cb' bytes). */
function papxProperties(page: Uint8Array, offset: number): ParagraphProperties {
  const cb = page[offset] ?? 0;
  const start = cb === 0 ? offset + 2 : offset + 1;
  const size = cb === 0 ? (page[offset + 1] ?? 0) * 2 : cb * 2 - 1;
  const end = Math.min(start + size, FKP_SIZE);
  const result = { ...DEFAULT_PARAGRAPH };
  // Skip the 2-byte style index, then walk the sprms.
  for (let at = start + 2; at + 2 <= end; ) {
    const sprm = (page[at] ?? 0) | ((page[at + 1] ?? 0) << 8);
    at += 2;
    const size = operandSize(sprm, page, at);
    if (at + size > end) {
      break;
    }
    const operand = page[at] ?? 0;
    if (sprm === SPRM_IN_TABLE) {
      result.inTable = operand !== 0;
    } else if (sprm === SPRM_ROW_END) {
      result.rowEnd = operand !== 0;
    } else if (sprm === SPRM_INNER_ROW_END) {
      result.innerRowEnd = operand !== 0;
    } else if (sprm === SPRM_TABLE_DEPTH) {
      result.inTable ||= operand > 0;
    }
    at += size;
  }
  return result;
}

/** Operand size of a sprm from its spra bits ([MS-DOC] 2.2.5.1), including the two variable-length exceptions. */
function operandSize(sprm: number, bytes: Uint8Array, at: number): number {
  switch (sprm >> 13) {
    case 0:
    case 1:
      return 1;
    case 2:
    case 4:
    case 5:
      return 2;
    case 3:
      return 4;
    case 7:
      return 3;
    default:
      if (sprm === SPRM_DEF_TABLE) {
        return ((bytes[at] ?? 0) | ((bytes[at + 1] ?? 0) << 8)) + 1;
      }
      if (sprm === SPRM_CHANGE_TABS && bytes[at] === 255) {
        const deleted = bytes[at + 1] ?? 0;
        return 3 + deleted * 4 + (bytes[at + 2 + deleted * 4] ?? 0) * 3;
      }
      return 1 + (bytes[at] ?? 0);
  }
}

/** Binary search over ascending keys: the last index whose key is <= value, or -1. */
function lastAtOrBefore(length: number, key: (index: number) => number, value: number): number {
  let low = 0;
  let high = length - 1;
  let found = -1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    if (key(mid) <= value) {
      found = mid;
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  return found;
}
