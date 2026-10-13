import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { layoutDocument } from '../../../src/webview/doc/layout.ts';
import { DOC_LIMITS, readWordDocument } from '../../../src/webview/doc/wordDocument.ts';
import { SAMPLE_DOC } from '../../fixtures/doc.ts';
import { writeCompoundFile } from '../../support/compoundFileWriter.ts';
import { writeWordDocument, type WordDocumentInput } from '../../support/wordWriter.ts';

const FIB_FC_LCB = 0x9a;
const PAIR_PAPX = 13;
const PAIR_CLX = 33;

function doc(input: Partial<WordDocumentInput> = {}): Buffer {
  return writeWordDocument({ body: ['Hello\r', 'world\r'], ...input });
}

/** Offset in the table stream of the structure at fc/lcb pair `index`, read back from the FIB. */
function located(word: Buffer, index: number): { fc: number; lcb: number } {
  return { fc: word.readUInt32LE(FIB_FC_LCB + index * 8), lcb: word.readUInt32LE(FIB_FC_LCB + index * 8 + 4) };
}

describe('readWordDocument', () => {
  it('decodes cp1252 and UTF-16 pieces in CP order', () => {
    const document = readWordDocument(writeWordDocument(SAMPLE_DOC));
    const main = document.stories.find((story) => story.kind === 'main');
    assert.ok(main);
    const text = document.text.slice(main.start, main.end);
    assert.ok(text.startsWith('Document Lens sample\rThis synthetic Word 97–2003 document'));
    assert.ok(text.includes('Prices: 5 € – “quoted” text.\r'));
    assert.ok(text.includes('\x0c日本語のテキストも表示されます。\rSecond page'));
    assert.ok(text.includes('Item\x07Qty\x07Remarks\x07\x07'));
    assert.equal(document.truncated, false);
  });

  it('reports the stories and the footnote references', () => {
    const document = readWordDocument(writeWordDocument(SAMPLE_DOC));
    assert.deepEqual(
      document.stories.map((story) => story.kind),
      ['main', 'footnotes'],
    );
    const [main, footnotes] = document.stories;
    assert.ok(main && footnotes);
    assert.equal(document.text.slice(footnotes.start, footnotes.end), '\x02A footnote at the end of the text.\r\r');
    assert.equal(document.footnoteRefs.length, 1);
    assert.equal(document.text[document.footnoteRefs[0] ?? 0], '\x02');
    assert.deepEqual(document.endnoteRefs, []);
  });

  it('reads the table flags of paragraph marks, past a sprmTDefTable', () => {
    const document = readWordDocument(writeWordDocument(SAMPLE_DOC));
    const at = (needle: string) => {
      const cp = document.text.indexOf(needle);
      assert.ok(cp >= 0, needle);
      return document.paragraphAt(cp + needle.length - 1);
    };
    assert.deepEqual(at('Document Lens sample\r'), { inTable: false, rowEnd: false, innerRowEnd: false });
    assert.deepEqual(at('Item\x07'), { inTable: true, rowEnd: false, innerRowEnd: false });
    assert.deepEqual(at('Remarks\x07\x07'), { inTable: true, rowEnd: true, innerRowEnd: false });
    assert.deepEqual(at('Two\r'), { inTable: true, rowEnd: false, innerRowEnd: false });
    assert.deepEqual(at('After the table.\r'), { inTable: false, rowEnd: false, innerRowEnd: false });
    assert.deepEqual(at('日本語のテキストも表示されます。\r'), { inTable: false, rowEnd: false, innerRowEnd: false });
  });

  it('reads pieces stored out of order, as after fast saves', () => {
    const inOrder = readWordDocument(writeWordDocument(SAMPLE_DOC));
    const reversed = readWordDocument(writeWordDocument({ ...SAMPLE_DOC, reversePieceStorage: true }));
    assert.equal(reversed.text, inOrder.text);
    const rowEnd = reversed.text.indexOf('Fresh\x07\x07') + 6;
    assert.equal(reversed.paragraphAt(rowEnd)?.rowEnd, true);
  });

  it('finds tracked deletions and hidden text in the character properties', () => {
    const document = readWordDocument(writeWordDocument(SAMPLE_DOC));
    const shown = (from: number, to: number) => document.text.slice(from, to);
    assert.deepEqual(
      document.hidden.map(([from, to]) => shown(from, to)),
      ['removed words ', ' (hidden note)'],
    );
    const mixed = readWordDocument(
      doc({ body: ['a ', { text: 'gone', deleted: true }, { text: '消えた', unicode: true, hidden: true }, { text: '残る', unicode: true }, '\r'] }),
    );
    assert.deepEqual(
      mixed.hidden.map(([from, to]) => mixed.text.slice(from, to)),
      ['gone消えた'],
    );
  });

  it('ignores character properties that are damaged or point outside the stream', () => {
    const PAIR_CHPX = 12;
    const outside = doc({
      body: ['a ', { text: 'gone', deleted: true }, '\r'],
      edit({ word, table }) {
        table.writeUInt32LE(0x3fffff, located(word, PAIR_CHPX).fc + 8);
      },
    });
    assert.deepEqual(readWordDocument(outside).hidden, []);
    const capped = doc({ body: ['a ', { text: 'gone', deleted: true }, ' b ', { text: 'too', hidden: true }, '\r'] });
    assert.equal(readWordDocument(capped, { ...DOC_LIMITS, maxHiddenRuns: 1 }).hidden.length, 1);
  });

  it('uses the 0Table stream when the FIB says so', () => {
    assert.equal(readWordDocument(doc({ tableStream: '0Table' })).text, 'Hello\rworld\r');
  });

  it('skips property lists before the piece table', () => {
    assert.equal(readWordDocument(doc({ clxPropertyLists: [[0x16, 0x24, 0x01], []] })).text, 'Hello\rworld\r');
  });

  it('has no paragraph properties without a PAPX bin table', () => {
    const document = readWordDocument(doc({ omitParagraphTable: true }));
    assert.equal(document.paragraphAt(5), undefined);
  });

  it('cuts the text at the character limit', () => {
    const document = readWordDocument(writeWordDocument(SAMPLE_DOC), { ...DOC_LIMITS, maxCharacters: 30 });
    assert.equal(document.text.length, 30);
    assert.equal(document.truncated, true);
    assert.deepEqual(
      document.stories.map((story) => [story.kind, story.start, story.end]),
      [['main', 0, 30]],
    );
  });
});

describe('readWordDocument with unsupported or hostile files', () => {
  it('explains password-protected documents', () => {
    assert.throws(() => readWordDocument(doc({ encrypted: true })), /^Error: This \.doc file is password-protected/);
    const encryptedPackage = writeCompoundFile([{ name: 'EncryptedPackage', data: new Uint8Array(10) }]);
    assert.throws(() => readWordDocument(encryptedPackage), /password-protected/);
  });

  it('explains Word 6.0/95 documents and other compound files', () => {
    assert.throws(() => readWordDocument(doc({ nFib: 0x68 })), /saved by Word 6\.0\/95 or earlier/);
    assert.throws(() => readWordDocument(doc({ ident: 0xa5dc, nFib: 0x65 })), /saved by Word 6\.0\/95 or earlier/);
    assert.throws(() => readWordDocument(doc({ ident: 0x1234 })), /Not a Word document \(the WordDocument stream has no Word header\)/);
    assert.throws(() => readWordDocument(writeCompoundFile([{ name: 'Workbook', data: new Uint8Array(10) }])), /actually an Excel workbook/);
    assert.throws(() => readWordDocument(writeCompoundFile([{ name: 'PowerPoint Document', data: new Uint8Array(10) }])), /actually a PowerPoint presentation/);
    assert.throws(() => readWordDocument(writeCompoundFile([{ name: 'Contents', data: new Uint8Array(10) }])), /Not a Word document \(the file has no WordDocument stream\)/);
  });

  it('refuses a truncated FIB and a missing table stream', () => {
    const tiny = writeCompoundFile([{ name: 'WordDocument', data: Buffer.from([0xec, 0xa5, 0xc1, 0x00]) }]);
    assert.throws(() => readWordDocument(tiny), /Corrupt \.doc file: the file information block is truncated/);
    const wrongTable = doc({
      edit({ word }) {
        word.writeUInt16LE(word.readUInt16LE(0x0a) & ~0x0200, 0x0a);
      },
    });
    assert.throws(() => readWordDocument(wrongTable), /the 0Table stream is missing/);
  });

  it('refuses a piece table that is missing, misplaced or malformed', () => {
    const editClx = (change: (table: Buffer, clx: { fc: number; lcb: number }, word: Buffer) => void) =>
      doc({
        edit({ word, table }) {
          change(table, located(word, PAIR_CLX), word);
        },
      });
    assert.throws(() => readWordDocument(editClx((_, clx, word) => word.writeUInt32LE(0, FIB_FC_LCB + PAIR_CLX * 8 + 4))), /the piece table is missing/);
    assert.throws(() => readWordDocument(editClx((_, clx, word) => word.writeUInt32LE(0x7fffffff, FIB_FC_LCB + PAIR_CLX * 8))), /the piece table is outside the table stream/);
    assert.throws(() => readWordDocument(editClx((table, clx) => table.writeUInt8(0x07, clx.fc))), /the piece table is missing/);
    assert.throws(() => readWordDocument(editClx((table, clx) => table.writeUInt32LE(0xfffffff0, clx.fc + 1))), /the piece table has an invalid size/);
    // CPs: [0, 6, 12]; make the middle one larger than the last.
    assert.throws(() => readWordDocument(editClx((table, clx) => table.writeUInt32LE(100, clx.fc + 5 + 4))), /the piece table is out of order/);
    assert.throws(() => readWordDocument(editClx((table, clx) => table.writeUInt32LE(1, clx.fc + 5))), /does not start at the first character/);
    // First Pcd's fc: point the piece past the end of the WordDocument stream.
    assert.throws(() => readWordDocument(editClx((table, clx) => table.writeUInt32LE(0x40000000 | 0x3ffffff0, clx.fc + 5 + 12 + 2))), /a text piece lies outside the WordDocument stream/);
  });

  it('refuses more pieces than the limit', () => {
    const many = doc({ body: Array.from({ length: 5 }, (_, i) => `piece ${i}\r`) });
    assert.throws(() => readWordDocument(many, { ...DOC_LIMITS, maxPieces: 4 }), /too many text pieces to preview \(5; limit 4\)/);
    assert.equal(readWordDocument(many, { ...DOC_LIMITS, maxPieces: 5 }).text.length, 40);
  });

  it('refuses negative story lengths', () => {
    const negative = doc({
      edit({ word }) {
        word.writeInt32LE(-5, 0x40 + 3 * 4);
      },
    });
    assert.throws(() => readWordDocument(negative), /a story has a negative length/);
  });

  it('ignores story lengths beyond the piece table', () => {
    const long = doc({
      edit({ word }) {
        word.writeInt32LE(0x7ffffff0, 0x40 + 3 * 4);
      },
    });
    const document = readWordDocument(long);
    assert.equal(document.text, 'Hello\rworld\r');
    assert.equal(document.truncated, false);
  });

  it('falls back to no paragraph properties when the PAPX pages are damaged', () => {
    const outside = doc({
      edit({ word, table }) {
        const papx = located(word, PAIR_PAPX);
        // One page: [fcFirst, fcLast, pn]; point pn far outside the stream.
        table.writeUInt32LE(0x3fffff, papx.fc + 8);
      },
    });
    assert.equal(readWordDocument(outside).paragraphAt(5), undefined);
    const emptyPage = doc({
      edit({ word, table }) {
        const pn = table.readUInt32LE(located(word, PAIR_PAPX).fc + 8);
        word.writeUInt8(0, pn * 512 + 511);
      },
    });
    assert.equal(readWordDocument(emptyPage).paragraphAt(5), undefined);
  });

  it('survives property lists with sizes that run past the page', () => {
    const intact = writeWordDocument({ body: [{ row: ['a', 'b'] }, 'after\r'] });
    const broken = doc({
      body: [{ row: ['a', 'b'] }, 'after\r'],
      edit({ word, table }) {
        const pn = table.readUInt32LE(located(word, PAIR_PAPX).fc + 8);
        const page = pn * 512;
        const crun = word.readUInt8(page + 511);
        for (let i = 0; i < crun; i++) {
          const bOffset = word.readUInt8(page + (crun + 1) * 4 + i * 13);
          if (bOffset) {
            // cb = 0 then cb' = 255: a GrpPrlAndIstd longer than the page.
            word.writeUInt8(0, page + bOffset * 2);
            word.writeUInt8(255, page + bOffset * 2 + 1);
          }
        }
      },
    });
    assert.ok(readWordDocument(intact).paragraphAt(2)?.inTable);
    assert.doesNotThrow(() => readWordDocument(broken).paragraphAt(2));
  });

  it('either reads or reports randomly damaged files, never anything worse', () => {
    const sample = writeWordDocument(SAMPLE_DOC);
    let state = 1;
    const random = (below: number) => {
      state = (state * 1103515245 + 12345) >>> 0;
      return state % below;
    };
    for (let i = 0; i < 500; i++) {
      const damaged = new Uint8Array(sample);
      for (let flips = 1 + random(8); flips > 0; flips--) {
        // Mostly the headers and tables at the start of the file, where damage has the most effect.
        damaged[random(2) ? random(4096) : random(damaged.length)] = random(256);
      }
      try {
        layoutDocument(readWordDocument(damaged));
      } catch (error) {
        assert.ok(!(error instanceof TypeError || error instanceof RangeError), (error as Error).stack);
        assert.match((error as Error).message, /^(Corrupt|Not a|This \.doc file)/);
      }
    }
  });
});
