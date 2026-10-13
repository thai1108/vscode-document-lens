// Turns the character stream of each story into blocks to render: paragraphs (split on paragraph marks), page breaks,
// and tables (cells end with 0x07; the paragraph table says which 0x07 ends a row). Fields show only their result,
// HYPERLINK fields become links, note references become numbers, and tracked deletions, hidden text and other control
// characters are dropped.
import { externalUrl } from '../../protocol.ts';
import type { ParagraphProperties, StoryKind, WordDocument } from './wordDocument.ts';

export const LAYOUT_LIMITS = {
  /** Paragraphs (including table cells) across all stories; a longer document is cut with a notice. */
  maxParagraphs: 100_000,
  /** Nesting of fields inside fields; deeper begin marks are ignored. */
  maxFieldDepth: 64,
  /** Characters of a field's code kept to look for a HYPERLINK target. */
  maxFieldCode: 2048,
};

export interface Run {
  text: string;
  href?: string;
  /** A footnote or endnote number. */
  noteRef?: boolean;
}

export interface Paragraph {
  type: 'paragraph';
  runs: Run[];
}

export interface PageBreak {
  type: 'pageBreak';
}

export interface Table {
  type: 'table';
  /** Rows of cells; each cell holds its paragraphs. */
  rows: Paragraph[][][];
}

export type Block = Paragraph | PageBreak | Table;

export interface Story {
  kind: StoryKind;
  blocks: Block[];
}

export interface DocumentLayout {
  stories: Story[];
  /** Characters or paragraphs past the limits were left out. */
  truncated: boolean;
}

const FIELD_BEGIN = 0x13;
const FIELD_SEPARATOR = 0x14;
const FIELD_END = 0x15;
const PARAGRAPH_MARK = 0x0d;
const CELL_MARK = 0x07;
const PAGE_BREAK = 0x0c;
const LINE_BREAK = 0x0b;
const COLUMN_BREAK = 0x0e;
const TAB = 0x09;
const NOTE_REF = 0x02;
const NON_BREAKING_HYPHEN = 0x1e;

export function layoutDocument(document: WordDocument, limits = LAYOUT_LIMITS): DocumentLayout {
  const budget = { paragraphs: limits.maxParagraphs };
  const stories: Story[] = [];
  let truncated = document.truncated;
  for (const range of document.stories) {
    const builder = new StoryBuilder(range.kind, document, budget, limits);
    const complete = builder.read(range.start, range.end);
    const blocks = builder.finish();
    if (blocks.length > 0 || range.kind === 'main') {
      stories.push({ kind: range.kind, blocks });
    }
    if (!complete) {
      truncated = true;
      break;
    }
  }
  return { stories, truncated };
}

interface Field {
  code: string;
  inResult: boolean;
  href?: string;
}

class StoryBuilder {
  private readonly kind: StoryKind;
  private readonly document: WordDocument;
  private readonly budget: { paragraphs: number };
  private readonly limits: typeof LAYOUT_LIMITS;
  private readonly blocks: Block[] = [];
  private readonly noteNumbers = new Map<number, string>();
  private runs: Run[] = [];
  private chars: string[] = [];
  private runHref: string | undefined;
  private runNote = false;
  private table: Table | undefined;
  private row: Paragraph[][] = [];
  private cell: Paragraph[] = [];
  private fields: Field[] = [];
  private ignoredFields = 0;
  private hiddenFields = 0;
  private href: string | undefined;
  private notesSeen = 0;
  /** A page break ended the paragraph text before its mark, so an empty paragraph right after it is not a blank line. */
  private afterBreak = false;
  private previousMark: 'paragraph' | 'cell' | 'row' | undefined;
  private readonly seenText = new Set<string>();

  constructor(kind: StoryKind, document: WordDocument, budget: { paragraphs: number }, limits: typeof LAYOUT_LIMITS) {
    this.kind = kind;
    this.document = document;
    this.budget = budget;
    this.limits = limits;
    if (kind === 'main') {
      document.footnoteRefs.forEach((cp, i) => this.noteNumbers.set(cp, String(i + 1)));
      document.endnoteRefs.forEach((cp, i) => this.noteNumbers.set(cp, String(i + 1)));
    }
  }

  /** Lays out CPs start..end; false when the paragraph budget ran out first. */
  read(start: number, end: number): boolean {
    const { text, hidden } = this.document;
    let next = hidden.findIndex(([, stop]) => stop > start);
    for (let cp = start; cp < end; cp++) {
      const range = hidden[next];
      if (range && range[0] <= cp) {
        cp = range[1] - 1;
        next++;
        continue;
      }
      const code = text.charCodeAt(cp);
      if (code === FIELD_BEGIN || code === FIELD_SEPARATOR || code === FIELD_END) {
        this.field(code);
      } else if (code === PARAGRAPH_MARK || code === CELL_MARK) {
        if (this.budget.paragraphs-- <= 0) {
          return false;
        }
        this.endParagraph(cp, code === CELL_MARK);
      } else if (this.hiddenFields > 0) {
        const field = this.fields.at(-1);
        if (field && !field.inResult && field.code.length < this.limits.maxFieldCode) {
          field.code += text[cp];
        }
      } else if (code === PAGE_BREAK) {
        this.pageBreak();
      } else if (code === LINE_BREAK || code === COLUMN_BREAK) {
        this.append('\n');
      } else if (code === NOTE_REF) {
        this.noteRef(cp);
      } else if (code === NON_BREAKING_HYPHEN) {
        this.append('\u2011');
      } else if (code >= 0x20 || code === TAB) {
        this.append(text[cp] ?? '');
      }
    }
    return true;
  }

  finish(): Block[] {
    if (this.chars.length > 0 || this.runs.length > 0) {
      this.addParagraph(this.takeParagraph());
    }
    this.closeTable();
    return this.blocks;
  }

  private field(code: number): void {
    if (code === FIELD_BEGIN) {
      if (this.fields.length >= this.limits.maxFieldDepth) {
        this.ignoredFields++;
        return;
      }
      this.fields.push({ code: '', inResult: false });
      this.hiddenFields++;
      return;
    }
    if (this.ignoredFields > 0) {
      this.ignoredFields -= code === FIELD_END ? 1 : 0;
      return;
    }
    const field = this.fields.at(-1);
    if (!field) {
      return;
    }
    if (code === FIELD_SEPARATOR) {
      if (!field.inResult) {
        field.inResult = true;
        field.href = hyperlinkTarget(field.code);
        this.hiddenFields--;
      }
    } else {
      this.fields.pop();
      if (!field.inResult) {
        this.hiddenFields--;
      }
    }
    this.href = this.fields.findLast((candidate) => candidate.href)?.href;
  }

  private noteRef(cp: number): void {
    const number = this.kind === 'main' ? this.noteNumbers.get(cp) : this.kind === 'footnotes' || this.kind === 'endnotes' ? String(++this.notesSeen) : undefined;
    if (number) {
      this.append(number, true);
    }
  }

  private append(text: string, noteRef = false): void {
    if (this.href !== this.runHref || noteRef !== this.runNote) {
      this.flushRun();
      this.runHref = this.href;
      this.runNote = noteRef;
    }
    this.chars.push(text);
    this.afterBreak = false;
  }

  private flushRun(): void {
    if (this.chars.length > 0) {
      const run: Run = { text: this.chars.join('') };
      if (this.runHref) {
        run.href = this.runHref;
      }
      if (this.runNote) {
        run.noteRef = true;
      }
      this.runs.push(run);
      this.chars = [];
    }
  }

  private takeParagraph(): Paragraph {
    this.flushRun();
    const paragraph: Paragraph = { type: 'paragraph', runs: this.runs };
    this.runs = [];
    return paragraph;
  }

  private endParagraph(cp: number, cellMark: boolean): void {
    const paragraph = this.takeParagraph();
    const properties = this.document.paragraphAt(cp);
    const empty = paragraph.runs.length === 0;
    const afterBreak = this.afterBreak;
    this.afterBreak = false;
    if (cellMark) {
      if (this.isRowEnd(properties, empty)) {
        this.endRow();
        this.previousMark = 'row';
      } else {
        this.cell.push(paragraph);
        this.row.push(this.cell);
        this.cell = [];
        this.previousMark = 'cell';
      }
      return;
    }
    this.previousMark = 'paragraph';
    if (properties?.innerRowEnd) {
      return;
    }
    // Without paragraph properties, a paragraph between two cells of a row is still taken to be inside the table.
    if (properties ? properties.inTable : this.row.length > 0) {
      this.cell.push(paragraph);
      return;
    }
    if (!(empty && afterBreak)) {
      this.addParagraph(paragraph);
    }
  }

  /** The row-end mark per the paragraph table, or, without one, a 0x07 straight after a cell mark with nothing between. */
  private isRowEnd(properties: ParagraphProperties | undefined, empty: boolean): boolean {
    return properties ? properties.rowEnd : empty && this.cell.length === 0 && this.previousMark === 'cell';
  }

  private endRow(): void {
    if (this.cell.length > 0) {
      this.row.push(this.cell);
      this.cell = [];
    }
    if (this.row.length > 0) {
      this.table ??= { type: 'table', rows: [] };
      this.table.rows.push(this.row);
      this.row = [];
    }
  }

  private closeTable(): void {
    this.endRow();
    if (this.table) {
      this.blocks.push(this.table);
      this.table = undefined;
    }
  }

  private addParagraph(paragraph: Paragraph): void {
    this.closeTable();
    if (this.kind !== 'main') {
      // Notes, headers and text boxes are listed compactly: no blank lines, and headers repeated per section once.
      const text = paragraph.runs.map((run) => run.text).join('');
      if (!text.trim() || (this.kind === 'headers' && this.seenText.has(text))) {
        return;
      }
      this.seenText.add(text);
    }
    this.blocks.push(paragraph);
  }

  private pageBreak(): void {
    if (this.row.length > 0 || this.cell.length > 0) {
      this.append('\n');
      return;
    }
    if (this.chars.length > 0 || this.runs.length > 0) {
      this.addParagraph(this.takeParagraph());
    }
    this.closeTable();
    this.blocks.push({ type: 'pageBreak' });
    this.afterBreak = true;
  }
}

/** The http(s) target of a HYPERLINK field code such as ` HYPERLINK "https://example.com" \t "_blank" `. */
export function hyperlinkTarget(code: string): string | undefined {
  const match = /^\s*HYPERLINK\b(.*)$/is.exec(code);
  if (!match) {
    return undefined;
  }
  // Switches with an argument (\l bookmark, \o tooltip, \t target frame) may come before the address.
  const rest = (match[1] ?? '').replace(/\\[lot]\s+("[^"]*"|\S+)/gi, ' ').replace(/\\[a-z]/gi, ' ');
  const target = /"([^"]*)"|(\S+)/.exec(rest);
  const address = target?.[1] ?? target?.[2];
  return address ? (externalUrl(address.trim()) ?? undefined) : undefined;
}
