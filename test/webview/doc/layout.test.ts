import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { hyperlinkTarget, LAYOUT_LIMITS, layoutDocument, type Block, type DocumentLayout } from '../../../src/webview/doc/layout.ts';
import type { ParagraphProperties, StoryKind, WordDocument } from '../../../src/webview/doc/wordDocument.ts';

const PLAIN: ParagraphProperties = { inTable: false, rowEnd: false, innerRowEnd: false };
const CELL: ParagraphProperties = { inTable: true, rowEnd: false, innerRowEnd: false };
const ROW_END: ParagraphProperties = { inTable: true, rowEnd: true, innerRowEnd: false };

interface FakeStory {
  kind: StoryKind;
  text: string;
}

/**
 * A WordDocument made of stories. Unless `properties` is false or `overrides` says otherwise, '\r' ends a plain
 * paragraph, '\x07' a cell, and '\x07' right after another '\x07' a row (true of these fixtures, not of Word in general).
 */
function fake(stories: FakeStory[], options: { properties?: boolean; footnoteRefs?: number[]; endnoteRefs?: number[]; overrides?: Map<number, ParagraphProperties> } = {}): WordDocument {
  const text = stories.map((story) => story.text).join('');
  let start = 0;
  const ranges = stories.map((story) => {
    const range = { kind: story.kind, start, end: start + story.text.length };
    start = range.end;
    return range;
  });
  return {
    text,
    truncated: false,
    stories: ranges,
    footnoteRefs: options.footnoteRefs ?? [],
    endnoteRefs: options.endnoteRefs ?? [],
    paragraphAt(cp) {
      if (options.properties === false) {
        return undefined;
      }
      const override = options.overrides?.get(cp);
      if (override) {
        return override;
      }
      if (text[cp] !== '\x07') {
        return PLAIN;
      }
      return text[cp - 1] === '\x07' ? ROW_END : CELL;
    },
  };
}

function main(text: string, options?: Parameters<typeof fake>[1]): DocumentLayout {
  return layoutDocument(fake([{ kind: 'main', text }], options));
}

/** Blocks as compact strings: paragraphs as their text (links as [text](href), notes as ^n), tables as rows of cells. */
function show(blocks: Block[]): unknown[] {
  return blocks.map((block) => {
    switch (block.type) {
      case 'pageBreak':
        return '---';
      case 'table':
        return block.rows.map((row) => row.map((cell) => cell.map((paragraph) => show([paragraph])[0]).join(' / ')));
      case 'paragraph':
        return block.runs.map((run) => (run.href ? `[${run.text}](${run.href})` : run.noteRef ? `^${run.text}` : run.text)).join('');
    }
  });
}

function mainBlocks(text: string, options?: Parameters<typeof fake>[1]): unknown[] {
  return show(main(text, options).stories[0]?.blocks ?? []);
}

describe('layoutDocument', () => {
  it('splits paragraphs on paragraph marks and keeps empty ones as blank lines', () => {
    assert.deepEqual(mainBlocks('One\r\rTwo\r'), ['One', '', 'Two']);
  });

  it('keeps tabs and line breaks, drops other control characters', () => {
    assert.deepEqual(mainBlocks('a\tb\x0bc\x0ed\x01\x08\x05e\x1ef\x1fg\x00\r'), ['a\tb\nc\nde\u2011fg']);
  });

  it('shows page breaks between paragraphs without adding a blank line', () => {
    assert.deepEqual(mainBlocks('Before\r\x0c\rAfter\r'), ['Before', '---', 'After']);
    assert.deepEqual(mainBlocks('Before\x0cAfter\r'), ['Before', '---', 'After']);
    assert.deepEqual(mainBlocks('One\r\x0cTwo\r'), ['One', '---', 'Two']);
    assert.deepEqual(mainBlocks('\x0ca\x07\x07\rEnd\r'), ['---', [['a']], '', 'End']);
  });

  it('shows only the result of fields, including nested ones', () => {
    assert.deepEqual(mainBlocks('Page \x13 PAGE \x143\x15 of \x13 NUMPAGES \x149\x15\r'), ['Page 3 of 9']);
    assert.deepEqual(mainBlocks('\x13 IF \x13 MERGEFIELD x \x14v\x15 = "v" "yes" \x14yes\x15!\r'), ['yes!']);
    assert.deepEqual(mainBlocks('Index\x13 XE "entry" \x15 entry\r'), ['Index entry']);
    assert.deepEqual(mainBlocks('\x13 TOC \x14Heading\x0d\x15Body\r'), ['Heading', 'Body']);
  });

  it('turns HYPERLINK fields into links, http(s) only', () => {
    assert.deepEqual(mainBlocks('See \x13 HYPERLINK "https://example.com/a b" \x14site\x15.\r'), ['See [site](https://example.com/a%20b).']);
    assert.deepEqual(mainBlocks('\x13 HYPERLINK "javascript:alert(1)" \x14x\x15\r'), ['x']);
    assert.deepEqual(mainBlocks('\x13 HYPERLINK \\l "_Toc1" \x14toc\x15\r'), ['toc']);
  });

  it('hides a field that never ends only until the end of its story', () => {
    const layout = layoutDocument(
      fake([
        { kind: 'main', text: 'Shown\r\x13 broken field\rhidden\r' },
        { kind: 'footnotes', text: '\x02Note\r' },
      ]),
    );
    assert.deepEqual(show(layout.stories[0]?.blocks ?? []), ['Shown', '', '']);
    assert.deepEqual(show(layout.stories[1]?.blocks ?? []), ['^1Note']);
  });

  it('ignores field marks nested deeper than the limit', () => {
    const deep = '\x13'.repeat(100) + 'x' + '\x15'.repeat(100);
    const layout = layoutDocument(fake([{ kind: 'main', text: `${deep}after\r` }]), { ...LAYOUT_LIMITS, maxFieldDepth: 8 });
    assert.deepEqual(show(layout.stories[0]?.blocks ?? []), ['after']);
  });

  it('builds tables from cell and row marks, with multi-paragraph cells', () => {
    const text = 'Intro\ra\x07b\x07\x07c\rd\x07e\x07\x07After\r';
    const overrides = new Map([[text.indexOf('c\r') + 1, CELL]]);
    assert.deepEqual(mainBlocks(text, { overrides }), ['Intro', [['a', 'b'], ['c / d', 'e']], 'After']);
  });

  it('keeps empty cells when the paragraph table tells cells from row ends', () => {
    const text = 'a\x07\x07c\x07\x07';
    const overrides = new Map([
      [1, CELL],
      [2, CELL],
      [4, CELL],
      [5, ROW_END],
    ]);
    assert.deepEqual(mainBlocks(text, { overrides }), [[['a', '', 'c']]]);
  });

  it('separates tables split by a page break and closes a table at the end of the story', () => {
    assert.deepEqual(mainBlocks('a\x07\x07\x0cb\x07\x07'), [[['a']], '---', [['b']]]);
  });

  it('flattens nested tables into their outer cell, skipping their row-end marks', () => {
    const text = 'x\ry\r\r\x07\x07';
    const overrides = new Map([
      [1, CELL],
      [3, CELL],
      [4, { inTable: true, rowEnd: false, innerRowEnd: true }],
    ]);
    assert.deepEqual(mainBlocks(text, { overrides }), [[['x / y / ']]]);
  });

  it('guesses rows without paragraph properties', () => {
    assert.deepEqual(mainBlocks('Intro\ra\x07b\x07\x07c\x07d\x07\x07Next\r', { properties: false }), ['Intro', [['a', 'b'], ['c', 'd']], 'Next']);
  });

  it('numbers note references in the main text and in the note stories', () => {
    const text = 'A\x02 b\x02 c\x02.\r';
    const layout = layoutDocument(
      fake(
        [
          { kind: 'main', text },
          { kind: 'footnotes', text: '\x02First\r\x02Second\r\r' },
          { kind: 'endnotes', text: '\x02End\r\r' },
        ],
        { footnoteRefs: [1, 4], endnoteRefs: [7] },
      ),
    );
    assert.deepEqual(
      layout.stories.map((story) => [story.kind, show(story.blocks)]),
      [
        ['main', ['A^1 b^2 c^1.']],
        ['footnotes', ['^1First', '^2Second']],
        ['endnotes', ['^1End']],
      ],
    );
  });

  it('lists headers and text boxes compactly, without repeats or blank lines', () => {
    const layout = layoutDocument(
      fake([
        { kind: 'main', text: 'Body\r' },
        { kind: 'headers', text: '\x03\r\x04\r\rPage header\r\rPage header\rFooter\r\r' },
        { kind: 'comments', text: '\x05Same\r\x05Same\r' },
        { kind: 'textBoxes', text: '\r' },
      ]),
    );
    assert.deepEqual(
      layout.stories.map((story) => [story.kind, show(story.blocks)]),
      [
        ['main', ['Body']],
        ['headers', ['Page header', 'Footer']],
        ['comments', ['Same', 'Same']],
      ],
    );
  });

  it('keeps an empty main story and stops at the paragraph limit', () => {
    assert.deepEqual(main('').stories, [{ kind: 'main', blocks: [] }]);
    const layout = layoutDocument(fake([{ kind: 'main', text: 'p\r'.repeat(10) }, { kind: 'footnotes', text: '\x02n\r' }]), { ...LAYOUT_LIMITS, maxParagraphs: 4 });
    assert.equal(layout.truncated, true);
    // Four whole paragraphs, then the text of the one whose mark was over the limit.
    assert.deepEqual(show(layout.stories[0]?.blocks ?? []), ['p', 'p', 'p', 'p', 'p']);
    assert.equal(layout.stories.length, 1);
  });

  it('passes on truncation from the document', () => {
    const document = { ...fake([{ kind: 'main', text: 'cut' }]), truncated: true };
    const layout = layoutDocument(document);
    assert.equal(layout.truncated, true);
    assert.deepEqual(show(layout.stories[0]?.blocks ?? []), ['cut']);
  });
});

describe('hyperlinkTarget', () => {
  it('finds the address among switches', () => {
    assert.equal(hyperlinkTarget(' HYPERLINK "https://example.com/" '), 'https://example.com/');
    assert.equal(hyperlinkTarget('HYPERLINK \\o "tip" "http://example.com/x" \\t "_blank"'), 'http://example.com/x');
    assert.equal(hyperlinkTarget(' hyperlink https://example.com/bare '), 'https://example.com/bare');
    assert.equal(hyperlinkTarget(' HYPERLINK "https://example.com/" \\l "part"'), 'https://example.com/');
  });

  it('ignores other fields, in-document links and non-web addresses', () => {
    assert.equal(hyperlinkTarget(' PAGE '), undefined);
    assert.equal(hyperlinkTarget(' HYPERLINK \\l "_Toc1" '), undefined);
    assert.equal(hyperlinkTarget(' HYPERLINK "file:///etc/passwd" '), undefined);
    assert.equal(hyperlinkTarget(' HYPERLINK "mailto:a@example.com" '), undefined);
    assert.equal(hyperlinkTarget(' HYPERLINK "https://user:pw@example.com/" '), undefined);
    assert.equal(hyperlinkTarget(' HYPERLINK '), undefined);
  });
});
