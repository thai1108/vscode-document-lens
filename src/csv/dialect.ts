// How a delimited text file is written: its delimiter (detected for .csv, a tab for .tsv) and line breaks.
import { CR, fieldValue, readRecords, stringUnits, type RecordSink } from './tokenizer.ts';

/** Delimiters a .csv file may use: commas, semicolons where the comma is the decimal separator, tabs. */
export const CSV_DELIMITERS = [',', ';', '\t'] as const;
export type Delimiter = (typeof CSV_DELIMITERS)[number];

/** Text looked at to choose the delimiter. */
const SAMPLE_CHARS = 64 * 1024;
const SAMPLE_RECORDS = 200;

/**
 * The delimiter of a .csv file: the one splitting most records into the same number (above one) of fields. Excel's
 * `sep=;` first line, when there is one, decides. Commas win ties, except over semicolons when the semicolon-split
 * fields hold numbers with a decimal comma (`1,5;2,5`).
 */
export function detectDelimiter(text: string): Delimiter {
  const hint = /^sep=([,;\t])\r?\n/.exec(text.slice(0, 8));
  if (hint) {
    return hint[1] as Delimiter;
  }
  const sample = text.length > SAMPLE_CHARS ? text.slice(0, SAMPLE_CHARS) : text;
  const truncated = sample.length < text.length;
  let best: Delimiter = ',';
  let bestScore = 0;
  for (const delimiter of CSV_DELIMITERS) {
    const { counts, decimalCommas } = sampleRecords(sample, delimiter, truncated);
    const score = consistency(counts);
    const wins = score > bestScore + 1e-9 || (delimiter === ';' && best === ',' && score > 0 && score >= bestScore - 0.05 && decimalCommas);
    if (wins) {
      best = delimiter;
      bestScore = score;
    }
  }
  return best;
}

/** Fields per record (blank lines left out) and whether any field is a number with a decimal comma. */
function sampleRecords(sample: string, delimiter: Delimiter, truncated: boolean): { counts: number[]; decimalCommas: boolean } {
  const counts: number[] = [];
  let fields = 0;
  let blank = false;
  let decimalCommas = false;
  const sink: RecordSink = {
    field(start, end) {
      fields++;
      blank = fields === 1 && start === end;
      if (!decimalCommas && end - start <= 32 && /^"?[+-]?\d+,\d+"?$/.test(sample.slice(start, end))) {
        decimalCommas = /^[+-]?\d+,\d+$/.test(fieldValue(sample, start, end));
      }
    },
    record(_end, next) {
      // The sample may cut the last record short.
      if (!(fields === 1 && blank) && !(truncated && next >= sample.length)) {
        counts.push(fields);
      }
      fields = 0;
      return true;
    },
  };
  readRecords(stringUnits(sample), delimiter.charCodeAt(0), sink, 0, SAMPLE_RECORDS);
  return { counts, decimalCommas };
}

/** Share of records with the most common field count, or 0 when that count is not above one. */
function consistency(counts: readonly number[]): number {
  const frequency = new Map<number, number>();
  for (const count of counts) {
    frequency.set(count, (frequency.get(count) ?? 0) + 1);
  }
  let mode = 0;
  let modeFrequency = 0;
  for (const [count, times] of frequency) {
    if (times > modeFrequency || (times === modeFrequency && count > mode)) {
      mode = count;
      modeFrequency = times;
    }
  }
  return mode > 1 ? modeFrequency / counts.length : 0;
}

export type LineEnding = '\r\n' | '\n' | '\r';

/** Line breaks between records, counted by kind. */
export class LineEndings {
  crlf = 0;
  lf = 0;
  cr = 0;

  /** Counts the break of a record ending at `end` with the next one starting at `next`, read from `text`. */
  add(text: string, end: number, next: number): void {
    if (next - end === 2) {
      this.crlf++;
    } else if (next - end === 1) {
      if (text.charCodeAt(end) === CR) {
        this.cr++;
      } else {
        this.lf++;
      }
    }
  }

  /** The most common one; CRLF (what Excel writes) for a file of one line. */
  get dominant(): LineEnding {
    if (this.lf > this.crlf && this.lf >= this.cr) {
      return '\n';
    }
    if (this.cr > this.crlf && this.cr > this.lf) {
      return '\r';
    }
    return '\r\n';
  }

  get mixed(): boolean {
    return [this.crlf, this.lf, this.cr].filter((count) => count > 0).length > 1;
  }
}

export function delimiterLabel(delimiter: string): string {
  switch (delimiter) {
    case ',':
      return 'Comma';
    case ';':
      return 'Semicolon';
    case '\t':
      return 'Tab';
    default:
      return delimiter;
  }
}

export function lineEndingLabel(lineEnding: LineEnding): string {
  return lineEnding === '\r\n' ? 'CRLF' : lineEnding === '\n' ? 'LF' : 'CR';
}
