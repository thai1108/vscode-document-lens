// Turns the character stream of each story into blocks to render: paragraphs (split on paragraph marks), page breaks,
// and tables (cells end with 0x07; the paragraph table says which 0x07 ends a row). Fields show only their result,
// HYPERLINK fields become links, note references become numbers, and tracked deletions, hidden text and other control
// characters are dropped.
import { externalUrl } from '../../protocol.ts';
import type { ParagraphProperties, StoryKind, WordDocument } from './wordDocument.ts';

export const LAYOUT_LIMITS = {
  /**
   * Elements across all stories, each of which becomes a DOM node: paragraphs, tables, rows, cells, page breaks and text
   * runs. A longer document is cut with a notice.
   */
  maxElements: 200_000,
  /** Characters shown across all stories, note numbers included; a longer document is cut with a notice. */
  maxOutputCharacters: 3_000_000,
  /** Cells per table row: Word 97–2003 tables have at most 63 columns, so further cells wrap to a new row. */
  maxColumns: 63,
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

/** What is left of LAYOUT_LIMITS.maxElements and maxOutputCharacters, shared by all stories. */
interface Budget {
  elements: number;
  characters: number;
}

export function layoutDocument(document: WordDocument, limits = LAYOUT_LIMITS): DocumentLayout {
  const budget: Budget = { elements: limits.maxElements, characters: limits.maxOutputCharacters };
  const stories: Story[] = [];
  let truncated = document.truncated;
  for (const range of document.stories) {
    const builder = new StoryBuilder(range.kind, document, budget, limits);
    const complete = builder.read(range.start, range.end);
    const blocks = builder.finish();
    if (blocks.length > 0 || range.kind === 'main') {
      stories.push({ kind: range.kind, blocks });
      // A labelled section around the story.
      budget.elements -= range.kind === 'main' ? 0 : 2;
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
  private readonly budget: Budget;
  private readonly limits: typeof LAYOUT_LIMITS;
  private readonly blocks: Block[] = [];
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

  constructor(kind: StoryKind, document: WordDocument, budget: Budget, limits: typeof LAYOUT_LIMITS) {
    this.kind = kind;
    this.document = document;
    this.budget = budget;
    this.limits = limits;
  }

  /** Lays out CPs start..end; false when the element or character budget ran out first. */
  read(start: number, end: number): boolean {
    const { text, hidden } = this.document;
    let next = hidden.findIndex(([, stop]) => stop > start);
    for (let cp = start; cp < end; cp++) {
      if (this.budget.elements <= 0 || this.budget.characters <= 0) {
        return false;
      }
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

  /**
   * An automatically numbered note reference (0x02). In the main text its number is its place among the automatic
   * references of its kind; in a note story, the notes' own 0x02 marks come in the same order. Custom marks are text.
   */
  private noteRef(cp: number): void {
    let index = -1;
    if (this.kind === 'main') {
      const { footnoteRefs, endnoteRefs } = this.document;
      index = indexOf(footnoteRefs, cp);
      index = index >= 0 ? index : indexOf(endnoteRefs, cp);
    } else if (this.kind === 'footnotes' || this.kind === 'endnotes') {
      index = this.notesSeen++;
    }
    if (index >= 0) {
      this.append(String(index + 1), true);
    }
  }

  private append(text: string, noteRef = false): void {
    if (this.href !== this.runHref || noteRef !== this.runNote) {
      this.flushRun();
      this.runHref = this.href;
      this.runNote = noteRef;
    }
    this.chars.push(text);
    this.budget.characters -= text.length;
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
      this.budget.elements--;
      this.chars = [];
    }
  }

  private takeParagraph(): Paragraph {
    this.flushRun();
    const paragraph: Paragraph = { type: 'paragraph', runs: this.runs };
    this.runs = [];
    this.budget.elements--;
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
        if (this.row.length >= this.limits.maxColumns) {
          this.pushRow();
        }
        this.row.push(this.cell);
        this.budget.elements--;
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
      this.budget.elements--;
      this.cell = [];
    }
    this.pushRow();
  }

  /** Adds the cells so far as a row of the current table, starting the table if needed. */
  private pushRow(): void {
    if (this.row.length > 0) {
      if (!this.table) {
        this.table = { type: 'table', rows: [] };
        this.budget.elements--;
      }
      this.table.rows.push(this.row);
      this.budget.elements--;
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
    this.budget.elements--;
    this.afterBreak = true;
  }
}

/** Binary search in ascending `values`: the index of `value`, or -1. */
function indexOf(values: readonly number[], value: number): number {
  let low = 0;
  let high = values.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    const candidate = values[mid] ?? 0;
    if (candidate === value) {
      return mid;
    }
    if (candidate < value) {
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  return -1;
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
