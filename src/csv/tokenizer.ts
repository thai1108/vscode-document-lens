// Splits delimited text into records and fields the way Excel reads it: RFC 4180 quoting (a field in double quotes
// may hold delimiters, line breaks and doubled quotes), CRLF, LF or CR line breaks, and leniently past the RFC: rows
// of any length, text after a closing quote kept, a quote that is never closed runs to the end of the file.
//
// The tokenizer works on code units, so the same code reads a decoded string (the webview) and an encoded file's
// bytes (the extension host, which needs byte offsets to copy unchanged records as they are). Quotes, delimiters and
// line breaks are ASCII, and in UTF-8 and Shift_JIS no byte of a multi-byte character is below 0x40, so the bytes
// split exactly where the decoded text does.

export const QUOTE = 0x22;
export const CR = 0x0d;
export const LF = 0x0a;

/** Code units of the text: a string's char codes, or the bytes (or 16-bit units) of an encoded file. */
export interface Units {
  readonly length: number;
  /** The unit at `index`, for 0 <= index < length. */
  at(index: number): number;
}

export function stringUnits(text: string): Units {
  return { length: text.length, at: (index) => text.charCodeAt(index) };
}

export interface RecordSink {
  /** A field: units [start, end), quotes included. `open` when its quote is never closed (it runs to the end). */
  field(start: number, end: number, open: boolean): void;
  /** The record ends; its line break is units [end, next), empty for a last line without one. False stops reading. */
  record(end: number, next: number): boolean;
}

/**
 * Reads records from `from` (the start of a record) until the text ends, the sink stops or `budget` records have
 * been read; returns where the next record starts. Linear in the length of the text whatever it holds.
 */
export function readRecords(units: Units, delimiter: number, sink: RecordSink, from = 0, budget = Infinity): number {
  const length = units.length;
  let i = from;
  let count = 0;
  while (i < length && count < budget) {
    for (;;) {
      const start = i;
      let open = false;
      if (i < length && units.at(i) === QUOTE) {
        open = true;
        i++;
        while (i < length) {
          if (units.at(i) === QUOTE) {
            if (i + 1 < length && units.at(i + 1) === QUOTE) {
              i += 2;
              continue;
            }
            i++;
            open = false;
            break;
          }
          i++;
        }
      }
      while (i < length) {
        const unit = units.at(i);
        if (unit === delimiter || unit === CR || unit === LF) {
          break;
        }
        i++;
      }
      sink.field(start, i, open);
      if (i < length && units.at(i) === delimiter) {
        i++;
        continue;
      }
      break;
    }
    const end = i;
    if (i < length) {
      i += units.at(i) === CR && i + 1 < length && units.at(i + 1) === LF ? 2 : 1;
    }
    count++;
    if (!sink.record(end, i)) {
      break;
    }
  }
  return i;
}

/**
 * A field's value from its source text: surrounding quotes removed, doubled quotes made single, text after the
 * closing quote kept. Returns at most `limit` + 1 characters, so a caller capping values can tell one was longer
 * without building the whole of a huge field.
 */
export function fieldValue(text: string, start: number, end: number, limit = Infinity): string {
  const room = (value: string) => limit + 1 - value.length;
  if (text.charCodeAt(start) !== QUOTE) {
    return text.slice(start, Math.min(end, start + room('')));
  }
  let value = '';
  let i = start + 1;
  while (room(value) > 0) {
    const quote = text.indexOf('"', i);
    if (quote < 0 || quote >= end) {
      value += text.slice(i, Math.min(end, i + room(value)));
      break;
    }
    value += text.slice(i, Math.min(quote, i + room(value)));
    if (quote + 1 < end && text.charCodeAt(quote + 1) === QUOTE) {
      value += '"';
      i = quote + 2;
    } else {
      value += text.slice(quote + 1, Math.min(end, quote + 1 + room(value)));
      break;
    }
  }
  return value.length > limit + 1 ? value.slice(0, limit + 1) : value;
}
