// Text of `<si>` (shared strings) and `<is>` (inline strings) elements: plain `<t>` or rich text runs `<r>`. Phonetic
// guides (`<rPh>`, the furigana Japanese Excel stores with a string) are not part of the text and are skipped.
import { SHEET_LIMITS } from '../limits.ts';
import type { ColorContext, ColorSpec } from '../model/colors.ts';
import type { Font, RichRun, StringTable } from '../model/types.ts';
import { readFontProperty, resolveRunFont } from './styles.ts';
import { decodeOoxmlEscapes, XmlScanner, type XmlTag } from './xml.ts';

/** Collects one string from the elements inside `<si>` or `<is>`; the owner forwards events and toggles text. */
export class RichTextReader {
  private readonly scanner: XmlScanner;
  private readonly colors: ColorContext;
  private text = '';
  private runs: RichRun[] = [];
  private run: { text: string; font?: Partial<Font & { colorSpec: ColorSpec }> } | undefined;
  private inPhonetic = false;
  private inRunProperties = false;
  private inText = false;

  constructor(scanner: XmlScanner, colors: ColorContext) {
    this.scanner = scanner;
    this.colors = colors;
  }

  reset(): void {
    this.text = '';
    this.runs = [];
    this.run = undefined;
    this.inPhonetic = false;
    this.inRunProperties = false;
    this.inText = false;
    this.scanner.wantText = false;
  }

  open(tag: XmlTag): void {
    switch (tag.name) {
      case 'rPh':
        this.inPhonetic = true;
        return;
      case 'r':
        if (!this.inPhonetic) {
          this.run = { text: '' };
        }
        return;
      case 'rPr':
        this.inRunProperties = this.run !== undefined;
        return;
      case 't':
        if (!this.inPhonetic && !tag.selfClosing) {
          this.inText = true;
          this.scanner.wantText = true;
        }
        return;
    }
    if (this.inRunProperties && this.run) {
      readFontProperty((this.run.font ??= {}), tag);
    }
  }

  close(name: string): void {
    switch (name) {
      case 'rPh':
        this.inPhonetic = false;
        return;
      case 'rPr':
        this.inRunProperties = false;
        return;
      case 't':
        this.inText = false;
        this.scanner.wantText = false;
        return;
      case 'r':
        if (this.run && !this.inPhonetic) {
          if (this.runs.length < SHEET_LIMITS.maxRichRuns) {
            this.runs.push(this.run.font ? { text: this.run.text, font: resolveRunFont(this.run.font, this.colors) } : { text: this.run.text });
          } else {
            const last = this.runs[this.runs.length - 1];
            if (last) {
              last.text += this.run.text;
            }
          }
          this.run = undefined;
        }
        return;
    }
  }

  onText(text: string): void {
    if (!this.inText || this.text.length >= SHEET_LIMITS.maxStringLength) {
      return;
    }
    const decoded = decodeOoxmlEscapes(text);
    this.text += decoded;
    if (this.run) {
      this.run.text += decoded;
    }
  }

  /** The string read since the last reset, and its runs if any of them is formatted. */
  result(): { text: string; runs?: RichRun[] } {
    const text = this.text.length > SHEET_LIMITS.maxStringLength ? this.text.slice(0, SHEET_LIMITS.maxStringLength) : this.text;
    return this.runs.some((run) => run.font) ? { text, runs: this.runs } : { text };
  }
}

/** Reads xl/sharedStrings.xml from text chunks. */
export async function readSharedStrings(chunks: AsyncIterable<string>, colors: ColorContext): Promise<{ table: StringTable; truncated: boolean }> {
  const texts: string[] = [];
  const rich = new Map<number, RichRun[]>();
  let truncated = false;
  let inItem = false;
  const scanner: XmlScanner = new XmlScanner({
    open(tag) {
      if (tag.name === 'si') {
        inItem = !tag.selfClosing;
        reader.reset();
        if (tag.selfClosing) {
          add('');
        }
      } else if (inItem) {
        reader.open(tag);
      }
    },
    close(name) {
      if (name === 'si' && inItem) {
        inItem = false;
        const { text, runs } = reader.result();
        if (runs && texts.length < SHEET_LIMITS.maxSharedStrings) {
          rich.set(texts.length, runs);
        }
        add(text);
      } else if (inItem) {
        reader.close(name);
      }
    },
    text(text) {
      reader.onText(text);
    },
  });
  const reader = new RichTextReader(scanner, colors);
  const add = (text: string) => {
    if (texts.length < SHEET_LIMITS.maxSharedStrings) {
      texts.push(text);
    } else {
      truncated = true;
    }
  };
  for await (const chunk of chunks) {
    scanner.write(chunk);
  }
  scanner.end();
  return { table: { texts, rich }, truncated };
}
