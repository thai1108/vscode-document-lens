import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { decodeBytes, decodeChars, paragraphRuns, TEXT_TYPE, textKind, toParagraphs } from '../../../src/webview/ppt/text.ts';
import { u16, u32 } from '../../support/pptWriter.ts';

describe('slide text', () => {
  it('decodes UTF-16 and 8-bit text atoms', () => {
    assert.equal(decodeChars(new Uint8Array(Buffer.from('四半期 review', 'utf16le'))), '四半期 review');
    // An odd trailing byte is ignored.
    assert.equal(decodeChars(new Uint8Array([0x41, 0x00, 0x42])), 'A');
    // 8-bit text is ISO 8859-1: 0x80–0x9f stay C1 controls instead of becoming Windows-1252 letters.
    assert.equal(decodeBytes(new Uint8Array([0x43, 0x61, 0x66, 0xe9, 0x80])), 'Café\u0080');
    assert.equal(decodeBytes(new Uint8Array(20_000).fill(0x61)).length, 20_000);
  });

  it('classifies text types', () => {
    assert.equal(textKind(TEXT_TYPE.title), 'title');
    assert.equal(textKind(TEXT_TYPE.centerTitle), 'title');
    assert.equal(textKind(TEXT_TYPE.body), 'body');
    assert.equal(textKind(TEXT_TYPE.centerBody), 'body');
    assert.equal(textKind(TEXT_TYPE.quarterBody), 'body');
    assert.equal(textKind(TEXT_TYPE.other), 'other');
    assert.equal(textKind(-1), 'other');
  });

  it('splits paragraphs on \\r, keeps \\v as a line break and drops empty paragraphs and control characters', () => {
    const paragraphs = toParagraphs('First\rSecond\vline\r \r\u0001Third\t*\u0007', TEXT_TYPE.other, []);
    assert.deepEqual(paragraphs, [
      { text: 'First', level: 0, bullet: false },
      { text: 'Second\nline', level: 0, bullet: false },
      { text: 'Third\t*', level: 0, bullet: false },
    ]);
  });

  it('gives body placeholders bullets by default and never titles', () => {
    assert.equal(toParagraphs('a', TEXT_TYPE.body, [])[0]?.bullet, true);
    assert.equal(toParagraphs('a', TEXT_TYPE.halfBody, [])[0]?.bullet, true);
    assert.equal(toParagraphs('a', TEXT_TYPE.centerBody, [])[0]?.bullet, false);
    assert.equal(toParagraphs('a', TEXT_TYPE.title, [{ count: 2, level: 3, bullet: true }])[0]?.bullet, false);
    assert.equal(toParagraphs('a', TEXT_TYPE.title, [{ count: 2, level: 3, bullet: true }])[0]?.level, 0);
  });

  it('applies the indent level and bullet flag of the run each paragraph starts in', () => {
    const runs = [
      { count: 4, level: 0 },
      { count: 4, level: 1, bullet: false },
      { count: 4, level: 2 },
    ];
    assert.deepEqual(
      toParagraphs('One\rTwo\rSix', TEXT_TYPE.body, runs).map((paragraph) => [paragraph.text, paragraph.level, paragraph.bullet]),
      [
        ['One', 0, true],
        ['Two', 1, false],
        ['Six', 2, true],
      ],
    );
    // Text beyond the runs keeps the last run's style.
    assert.equal(toParagraphs('a\rb\rc', TEXT_TYPE.body, [{ count: 2, level: 1 }])[2]?.level, 1);
  });

  it('reads paragraph runs past every optional field of TextPFException', () => {
    const allFields = 0x0000000f | 0x80 | 0x10 | 0x40 | 0x20 | 0x800 | 0x1000 | 0x2000 | 0x4000 | 0x100 | 0x400 | 0x8000 | 0x100000 | 0x10000 | 0xe0000 | 0x200000;
    const bytes = Buffer.concat([
      // Run 1: 6 characters, level 1, every field, bullet on, two tab stops.
      u32(6),
      u16(1),
      u32(allFields),
      u16(0x0001),
      u16(0x2022, 0, 100),
      u32(0xfe000000),
      u16(0, 100, 0, 0, 576, 288, 576),
      u16(2),
      u32(0x0000_0240, 0x0001_0480),
      u16(0, 0, 0),
      // Run 2: the rest, level 4 (capped from 9), bullet flags without the hasBullet mask.
      u32(10),
      u16(9),
      u32(0x0002),
      u16(0x0001),
      // Character runs follow; they are not read.
      u32(16, 0),
    ]);
    assert.deepEqual(paragraphRuns(new Uint8Array(bytes), 15), [
      { count: 6, level: 1, bullet: true },
      { count: 10, level: 4, bullet: undefined },
    ]);
  });

  it('stops at malformed paragraph runs instead of failing', () => {
    assert.deepEqual(paragraphRuns(new Uint8Array(u32(5)), 4), []);
    // A tab stop count that runs past the atom.
    const truncated = Buffer.concat([u32(5), u16(0), u32(0x100000), u16(1000)]);
    assert.deepEqual(paragraphRuns(new Uint8Array(truncated), 4), []);
    // A zero-length run ends the list.
    assert.deepEqual(paragraphRuns(new Uint8Array(Buffer.concat([u32(0), u16(0), u32(0)])), 4), []);
  });
});
