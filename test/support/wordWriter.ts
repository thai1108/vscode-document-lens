// Minimal Word 97–2003 (.doc) writer for synthetic test documents ([MS-DOC]): a FIB, the text as one piece per input
// item (cp1252 when every character fits, UTF-16LE otherwise), the piece table, paragraph properties for tables,
// character properties for special characters, tracked deletions and hidden text, fields, footnotes and one section,
// packed in a compound file.
import { writeCompoundFile, type CompoundFileOptions, type CompoundStream } from './compoundFileWriter.ts';

/** One piece of the main text; a row becomes cells ending in 0x07 plus the row-end mark. */
export type DocItem = string | { text: string; unicode?: boolean; deleted?: boolean; hidden?: boolean } | { row: string[] };

export interface WordDocumentInput {
  body: DocItem[];
  /** Footnote texts, referenced in order by the 0x02 characters of the body. */
  footnotes?: string[];
  /** FIB overrides for hostile-file tests. */
  ident?: number;
  nFib?: number;
  encrypted?: boolean;
  tableStream?: '0Table' | '1Table';
  /** Store the pieces back to front in the WordDocument stream, as fast-saved documents do. */
  reversePieceStorage?: boolean;
  /** Leave out the paragraph property bin table. */
  omitParagraphTable?: boolean;
  /** Property lists (Prc grpprls) written in the CLX before the piece table. */
  clxPropertyLists?: number[][];
  /** Edits the streams just before they are packed (e.g. to corrupt the piece table). */
  edit?(streams: { word: Buffer; table: Buffer }): void;
  extraStreams?: CompoundStream[];
  compound?: CompoundFileOptions;
}

const FIB_SIZE = 900;
const TEXT_START = 0x400;
const PAGE = 512;
const PAIR_STSHF = 1;
const PAIR_FOOTNOTE_REFS = 2;
const PAIR_FOOTNOTE_TEXT = 3;
const PAIR_SECTIONS = 6;
const PAIR_CHPX = 12;
const PAIR_PAPX = 13;
const PAIR_FONTS = 15;
const PAIR_FIELDS = 16;
const PAIR_DOP = 31;
const PAIR_CLX = 33;
const PAIR_COUNT = 0x5d;
const SPECIAL = new Set([0x01, 0x02, 0x08, 0x13, 0x14, 0x15]);
const SPRM_SPECIAL = [0x55, 0x08, 0x01];
const SPRM_DELETED = [0x00, 0x08, 0x01];
const SPRM_HIDDEN = [0x3c, 0x08, 0x01];
const FIELD_TYPES: Record<string, number> = { HYPERLINK: 88, PAGE: 33, DATE: 31 };

const cp1252Bytes = (() => {
  const decoded = new TextDecoder('windows-1252').decode(Uint8Array.from({ length: 256 }, (_, i) => i));
  return new Map([...decoded].map((char, byte) => [char, byte]));
})();

interface Piece {
  text: string;
  compressed: boolean;
  cpStart: number;
  offset: number;
  /** Character properties of the whole piece. */
  grpprl: number[];
}

interface ParagraphMark {
  cp: number;
  inTable: boolean;
  rowEnd: boolean;
  cells: number;
}

export function writeWordDocument(input: WordDocumentInput): Buffer {
  // Main text pieces and the table properties of each paragraph mark.
  const pieces: Piece[] = [];
  const marks = new Map<number, ParagraphMark>();
  let cp = 0;
  const addPiece = (text: string, unicode: boolean, table?: number, grpprl: number[] = []) => {
    const compressed = !unicode && [...text].every((char) => cp1252Bytes.has(char));
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      if (code === 0x0d || code === 0x07) {
        const rowEnd = table !== undefined && i === text.length - 1;
        marks.set(cp + i, { cp: cp + i, inTable: table !== undefined || code === 0x07, rowEnd, cells: table ?? 0 });
      }
    }
    pieces.push({ text, compressed, cpStart: cp, offset: 0, grpprl });
    cp += text.length;
  };
  for (const item of input.body) {
    if (typeof item === 'string') {
      addPiece(item, false);
    } else if ('row' in item) {
      addPiece(item.row.map((cell) => `${cell}\x07`).join('') + '\x07', false, item.row.length);
    } else {
      addPiece(item.text, item.unicode ?? false, undefined, [...(item.deleted ? SPRM_DELETED : []), ...(item.hidden ? SPRM_HIDDEN : [])]);
    }
  }
  const ccpText = cp;
  const mainText = pieces.map((piece) => piece.text).join('');
  const footnoteRefs = [...mainText].flatMap((char, i) => (char === '\x02' ? [i] : []));
  const footnotes = input.footnotes ?? [];
  const footnoteStarts: number[] = [];
  let footnoteText = '';
  for (const note of footnotes) {
    footnoteStarts.push(footnoteText.length);
    footnoteText += `\x02${note}\r`;
  }
  let ccpFtn = 0;
  if (footnoteText) {
    // Word ends the footnote story with a guard paragraph, and the whole text with one more paragraph mark.
    footnoteText += '\r';
    ccpFtn = footnoteText.length;
    addPiece(`${footnoteText}\r`, false);
  }

  // WordDocument stream: FIB, text, then the FKP pages.
  const word: number[] = new Array<number>(TEXT_START).fill(0);
  const storage = input.reversePieceStorage ? [...pieces].reverse() : pieces;
  for (const piece of storage) {
    if (!piece.compressed && word.length % 2) {
      word.push(0);
    }
    piece.offset = word.length;
    for (const char of piece.text) {
      if (piece.compressed) {
        word.push(cp1252Bytes.get(char) ?? 0x3f);
      } else {
        for (let i = 0; i < char.length; i++) {
          const unit = char.charCodeAt(i);
          word.push(unit & 0xff, unit >> 8);
        }
      }
    }
  }
  const textEnd = word.length;
  const fcOf = (position: number) => {
    const piece = pieces.findLast((candidate) => candidate.cpStart <= position);
    if (!piece) {
      throw new Error(`no piece for CP ${position}`);
    }
    return piece.offset + (position - piece.cpStart) * (piece.compressed ? 1 : 2);
  };
  const charSize = (position: number) => (pieces.findLast((candidate) => candidate.cpStart <= position)?.compressed ? 1 : 2);

  // Character runs in stream order: each piece's properties, plus sprmCFSpec on special characters.
  const charRuns: { end: number; grpprl: number[] }[] = [];
  for (const piece of storage) {
    const size = piece.compressed ? 1 : 2;
    for (let i = 0; i < piece.text.length; i++) {
      if (SPECIAL.has(piece.text.charCodeAt(i))) {
        const fc = piece.offset + i * size;
        charRuns.push({ end: fc, grpprl: piece.grpprl }, { end: fc + size, grpprl: [...piece.grpprl, ...SPRM_SPECIAL] });
      }
    }
    charRuns.push({ end: piece.offset + piece.text.length * size, grpprl: piece.grpprl });
  }
  // Every paragraph keeps its own PAPX entry: readers find paragraph ends from these boundaries.
  const paraRuns: { end: number; grpprl: number[] }[] = [...marks.values()]
    .map((mark) => ({ end: fcOf(mark.cp) + charSize(mark.cp), grpprl: paragraphSprms(mark) }))
    .sort((a, b) => a.end - b.end);
  if ((paraRuns.at(-1)?.end ?? TEXT_START) < textEnd) {
    paraRuns.push({ end: textEnd, grpprl: [] });
  }

  padTo(word, PAGE);
  const chpx = writeFkps(word, mergeRuns(charRuns.sort((a, b) => a.end - b.end), TEXT_START), TEXT_START, 'chpx');
  const papx = writeFkps(word, paraRuns, TEXT_START, 'papx');

  // Table stream.
  const table: number[] = [];
  const pairs = new Map<number, [number, number]>();
  const put = (index: number, bytes: number[]) => {
    pairs.set(index, [table.length, bytes.length]);
    table.push(...bytes);
  };
  put(PAIR_STSHF, stylesheet());
  if (footnotes.length) {
    put(PAIR_FOOTNOTE_REFS, [...u32s([...footnoteRefs, ccpText]), ...footnoteRefs.flatMap(() => [1, 0])]);
    put(PAIR_FOOTNOTE_TEXT, u32s([...footnoteStarts, ccpFtn - 1, ccpFtn]));
  }
  put(PAIR_SECTIONS, [...u32s([0, ccpText]), ...u16s([0]), ...u32s([0xffffffff]), ...u16s([0]), ...u32s([0xffffffff])]);
  put(PAIR_CHPX, binTable(chpx));
  if (!input.omitParagraphTable) {
    put(PAIR_PAPX, binTable(papx));
  }
  put(PAIR_FONTS, fontTable());
  const fields = fieldTable(mainText, ccpText);
  if (fields.length) {
    put(PAIR_FIELDS, fields);
  }
  put(PAIR_DOP, new Array<number>(500).fill(0));
  put(PAIR_CLX, clx(pieces, cp, input.clxPropertyLists ?? []));

  const wordBuffer = Buffer.from(word);
  writeFib(wordBuffer, input, pairs, { ccpText, ccpFtn });
  const tableBuffer = Buffer.from(table);
  input.edit?.({ word: wordBuffer, table: tableBuffer });
  return writeCompoundFile(
    [
      { name: 'WordDocument', data: wordBuffer },
      { name: input.tableStream ?? '1Table', data: tableBuffer },
      ...(input.extraStreams ?? []),
    ],
    input.compound,
  );
}

function paragraphSprms(mark: ParagraphMark): number[] {
  if (!mark.inTable) {
    return [];
  }
  const sprms = [0x16, 0x24, 0x01];
  if (mark.rowEnd) {
    // sprmPFTtp, then sprmTDefTable with evenly spaced cells: cb, itcMac, rgdxaCenter, one empty TC80 per cell.
    const width = Math.floor(9000 / mark.cells);
    const centers = Array.from({ length: mark.cells + 1 }, (_, i) => i * width);
    const operand = [mark.cells, ...u16s(centers), ...new Array<number>(mark.cells * 20).fill(0)];
    sprms.push(0x17, 0x24, 0x01, 0x08, 0xd6, ...u16s([operand.length + 1]), ...operand);
  }
  return sprms;
}

/** Joins neighbouring runs with the same properties and drops empty ones. */
function mergeRuns(runs: { end: number; grpprl: number[] }[], start: number): { end: number; grpprl: number[] }[] {
  const merged: { end: number; grpprl: number[] }[] = [];
  let previousEnd = start;
  for (const run of runs) {
    if (run.end <= previousEnd) {
      continue;
    }
    const last = merged.at(-1);
    if (last && last.grpprl.join() === run.grpprl.join()) {
      last.end = run.end;
    } else {
      merged.push({ ...run });
    }
    previousEnd = run.end;
  }
  return merged;
}

/** Packs runs into 512-byte FKP pages appended to `word`; returns the bin table's page boundaries and page numbers. */
function writeFkps(word: number[], runs: { end: number; grpprl: number[] }[], start: number, kind: 'chpx' | 'papx'): { fcs: number[]; pages: number[] } {
  const maxRuns = kind === 'chpx' ? 0x65 : 0x1d;
  const entrySize = kind === 'chpx' ? 1 : 13;
  const fcs = [start];
  const pages: number[] = [];
  let index = 0;
  let first = start;
  while (index < runs.length) {
    const page = new Array<number>(PAGE).fill(0);
    const blobs = new Map<string, number>();
    let top = PAGE - 1;
    const included: { end: number; grpprl: number[] }[] = [];
    while (index < runs.length && included.length < maxRuns) {
      const run = runs[index];
      if (!run) {
        break;
      }
      const key = run.grpprl.join();
      const blob = run.grpprl.length === 0 || blobs.has(key) ? [] : propertyBlob(run.grpprl, kind);
      const nextTop = (top - blob.length) & ~1;
      if ((included.length + 2) * 4 + (included.length + 1) * entrySize > nextTop) {
        break;
      }
      if (blob.length) {
        top = nextTop;
        page.splice(top, blob.length, ...blob);
        blobs.set(key, top / 2);
      }
      included.push(run);
      index++;
    }
    const crun = included.length;
    if (crun === 0) {
      throw new Error('writeWordDocument: a property list does not fit in an FKP page');
    }
    page[PAGE - 1] = crun;
    let fc = first;
    writeU32(page, 0, fc);
    included.forEach((run, i) => {
      fc = run.end;
      writeU32(page, (i + 1) * 4, fc);
      page[(crun + 1) * 4 + i * entrySize] = run.grpprl.length ? (blobs.get(run.grpprl.join()) ?? 0) : 0;
    });
    pages.push(word.length / PAGE);
    word.push(...page);
    first = fc;
    fcs.push(fc);
  }
  return { fcs, pages };
}

/** A Chpx (cb, grpprl) or PapxInFkp (cb or 0 + cb', then istd 0 and the sprms). */
function propertyBlob(grpprl: number[], kind: 'chpx' | 'papx'): number[] {
  if (kind === 'chpx') {
    return [grpprl.length, ...grpprl];
  }
  const body = [0, 0, ...grpprl];
  return body.length % 2 ? [(body.length + 1) / 2, ...body] : [0, body.length / 2, ...body];
}

function binTable({ fcs, pages }: { fcs: number[]; pages: number[] }): number[] {
  return [...u32s(fcs), ...u32s(pages)];
}

function clx(pieces: Piece[], end: number, propertyLists: number[][]): number[] {
  const cps = [...pieces.map((piece) => piece.cpStart), end];
  const pcds = pieces.flatMap((piece) => [0, 0, ...u32s([piece.compressed ? (piece.offset * 2) | 0x40000000 : piece.offset]), 0, 0]);
  const plc = [...u32s(cps), ...pcds];
  return [...propertyLists.flatMap((grpprl) => [0x01, ...u16s([grpprl.length]), ...grpprl]), 0x02, ...u32s([plc.length]), ...plc];
}

/** PlcFld for the main text: one CP per field character, then Fld entries (ch + type, or ch + flags at the end). */
function fieldTable(text: string, ccpText: number): number[] {
  const positions: number[] = [];
  const flds: number[] = [];
  const open: { type: number; separated: boolean }[] = [];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code === 0x13) {
      const name = /^\s*(\w+)/.exec(text.slice(i + 1, i + 40))?.[1]?.toUpperCase() ?? '';
      open.push({ type: FIELD_TYPES[name] ?? 0, separated: false });
      positions.push(i);
      flds.push(0x13, open.at(-1)?.type ?? 0);
    } else if (code === 0x14) {
      const field = open.at(-1);
      if (field) {
        field.separated = true;
      }
      positions.push(i);
      flds.push(0x14, 0xff);
    } else if (code === 0x15) {
      const field = open.pop();
      positions.push(i);
      flds.push(0x15, field?.separated ? 0x80 : 0);
    }
  }
  return positions.length ? [...u32s([...positions, ccpText]), ...flds] : [];
}

/** A stylesheet whose only style is Normal (istd 0), followed by the empty fixed slots up to istd 14. */
function stylesheet(): number[] {
  const name = u16s([6, ...[...'Normal'].map((char) => char.charCodeAt(0)), 0]);
  // LPUpxPapx (istd 0, no sprms), then an empty LPUpxChpx.
  const upx = u16s([2, 0, 0]);
  const size = 10 + name.length + upx.length;
  // StdfBase: sti 0; stk paragraph, no base style; two UPXs, next style 0; bchUpe; grfstd.
  const std = [...u16s([0x0000, 0xfff1, 0x0002, size, 0]), ...name, ...upx];
  const fixedSlots = 15;
  const stshi = u16s([fixedSlots, 0x000a, 0x0001, 0x005b, fixedSlots, 0, 0, 0, 0]);
  return [...u16s([stshi.length]), ...stshi, ...u16s([std.length]), ...std, ...u16s(new Array<number>(fixedSlots - 1).fill(0))];
}

/** A font table with one font, Times New Roman. */
function fontTable(): number[] {
  const name = [...'Times New Roman\0'].flatMap((char) => u16s([char.charCodeAt(0)]));
  const ffn = [0, 0x16, ...u16s([400]), 0, 0, ...new Array<number>(10).fill(0), ...new Array<number>(24).fill(0), ...name];
  ffn[0] = ffn.length - 1;
  return [...u16s([1, 0]), ...ffn];
}

function writeFib(word: Buffer, input: WordDocumentInput, pairs: Map<number, [number, number]>, lengths: { ccpText: number; ccpFtn: number }): void {
  const tableStream = input.tableStream ?? '1Table';
  word.writeUInt16LE(input.ident ?? 0xa5ec, 0x00);
  word.writeUInt16LE(input.nFib ?? 0x00c1, 0x02);
  word.writeUInt16LE(0x0409, 0x06);
  word.writeUInt16LE((tableStream === '1Table' ? 0x0200 : 0) | (input.encrypted ? 0x0100 : 0) | 0x1000, 0x0a);
  word.writeUInt16LE(0x00bf, 0x0c);
  word.writeUInt16LE(0x000e, 0x20);
  const rgLw = 0x22 + 0x0e * 2 + 2;
  word.writeUInt16LE(0x0016, rgLw - 2);
  word.writeUInt32LE(word.length, rgLw);
  word.writeInt32LE(lengths.ccpText, rgLw + 3 * 4);
  word.writeInt32LE(lengths.ccpFtn, rgLw + 4 * 4);
  const rgFcLcb = rgLw + 0x16 * 4 + 2;
  word.writeUInt16LE(PAIR_COUNT, rgFcLcb - 2);
  for (const [index, [fc, lcb]] of pairs) {
    word.writeUInt32LE(fc, rgFcLcb + index * 8);
    word.writeUInt32LE(lcb, rgFcLcb + index * 8 + 4);
  }
  if (rgFcLcb + PAIR_COUNT * 8 + 2 !== FIB_SIZE) {
    throw new Error('FIB layout mismatch');
  }
}

function padTo(bytes: number[], size: number): void {
  while (bytes.length % size) {
    bytes.push(0);
  }
}

function writeU32(bytes: number[], offset: number, value: number): void {
  bytes.splice(offset, 4, value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff);
}

function u32s(values: number[]): number[] {
  return values.flatMap((value) => [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff]);
}

function u16s(values: number[]): number[] {
  return values.flatMap((value) => [value & 0xff, (value >>> 8) & 0xff]);
}
