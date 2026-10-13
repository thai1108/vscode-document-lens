// A small streaming XML tokenizer for SpreadsheetML parts. Sheet parts can be hundreds of megabytes, so text is fed
// in chunks and never held whole, and nothing is built per element beyond what the handler asks for. Namespace
// prefixes are dropped from element names (Excel and other writers disagree on them); DOCTYPEs are skipped and
// entities other than XML's five predefined ones and character references are left as they are, so a hostile DTD
// cannot expand anything.
import { SHEET_LIMITS } from '../limits.ts';

export interface XmlHandler {
  /** Element start; `tag` is reused for the next element, so read what you need during the call. */
  open?(tag: XmlTag): void;
  /** Element end, also called right after `open` for a self-closing element. */
  close?(name: string): void;
  /** Character data (entities decoded) while `scanner.wantText` is true. */
  text?(text: string): void;
}

export class XmlTag {
  /** Local name, without namespace prefix. */
  name = '';
  selfClosing = false;
  private source = '';
  private from = 0;
  private to = 0;

  set(source: string, name: string, from: number, to: number, selfClosing: boolean): void {
    this.source = source;
    this.name = name;
    this.from = from;
    this.to = to;
    this.selfClosing = selfClosing;
  }

  /** Attribute value by exact name, or else by local name for a prefixed attribute (`id` finds `r:id`). */
  attr(name: string): string | undefined {
    const s = this.source;
    let local: string | undefined;
    let i = this.from;
    while (i < this.to) {
      while (i < this.to && isSpace(s.charCodeAt(i))) {
        i++;
      }
      const nameStart = i;
      while (i < this.to && s.charCodeAt(i) !== 61 && !isSpace(s.charCodeAt(i))) {
        i++;
      }
      const nameEnd = i;
      while (i < this.to && isSpace(s.charCodeAt(i))) {
        i++;
      }
      if (s.charCodeAt(i) !== 61) {
        // Not name="value": skip the stray token.
        i = Math.max(i, nameEnd + 1);
        continue;
      }
      i++;
      while (i < this.to && isSpace(s.charCodeAt(i))) {
        i++;
      }
      const quote = s.charCodeAt(i);
      if (quote !== 34 && quote !== 39) {
        return local;
      }
      const valueEnd = s.indexOf(quote === 34 ? '"' : "'", i + 1);
      if (valueEnd < 0 || valueEnd > this.to) {
        return local;
      }
      const length = nameEnd - nameStart;
      if (length === name.length && s.startsWith(name, nameStart)) {
        return decodeEntities(s.slice(i + 1, valueEnd));
      }
      if (local === undefined && length > name.length && s.charCodeAt(nameEnd - name.length - 1) === 58 && s.startsWith(name, nameEnd - name.length) && !s.startsWith('xmlns:', nameStart)) {
        local = decodeEntities(s.slice(i + 1, valueEnd));
      }
      i = valueEnd + 1;
    }
    return local;
  }
}

export class XmlScanner {
  /** Whether character data is decoded and passed to the handler; off by default, which skips it cheaply. */
  wantText = false;
  private readonly handler: XmlHandler;
  private readonly tag = new XmlTag();
  private buffer = '';

  constructor(handler: XmlHandler) {
    this.handler = handler;
  }

  write(chunk: string): void {
    const s = this.buffer.length ? this.buffer + chunk : chunk;
    const consumed = this.scan(s);
    this.buffer = consumed < s.length ? s.slice(consumed) : '';
    if (this.buffer.length > SHEET_LIMITS.maxXmlToken) {
      throw new Error('Malformed XML: a single element or text run is too large');
    }
  }

  /** Flushes trailing character data; an unterminated tag at the end is dropped. */
  end(): void {
    if (this.buffer && this.wantText && this.buffer.indexOf('<') < 0) {
      this.handler.text?.(decodeEntities(this.buffer));
    }
    this.buffer = '';
  }

  /** Processes complete tokens and returns how many characters were consumed. */
  private scan(s: string): number {
    const handler = this.handler;
    let pos = 0;
    for (;;) {
      const lt = s.indexOf('<', pos);
      if (lt < 0) {
        // Text up to the end may continue in the next chunk.
        return this.wantText ? pos : s.length;
      }
      if (lt > pos && this.wantText) {
        handler.text?.(decodeEntities(s.slice(pos, lt)));
      }
      pos = lt;
      const next = s.charCodeAt(lt + 1);
      if (next === 33) {
        // <!-- -->, <![CDATA[ ]]> or a declaration such as <!DOCTYPE ...>.
        if (s.startsWith('<!--', lt)) {
          const end = s.indexOf('-->', lt + 4);
          if (end < 0) {
            return pos;
          }
          pos = end + 3;
        } else if (s.startsWith('<![CDATA[', lt)) {
          const end = s.indexOf(']]>', lt + 9);
          if (end < 0) {
            return pos;
          }
          if (this.wantText) {
            handler.text?.(s.slice(lt + 9, end));
          }
          pos = end + 3;
        } else {
          const end = declarationEnd(s, lt + 2);
          if (end < 0) {
            return pos;
          }
          pos = end + 1;
        }
        continue;
      }
      if (next === 63) {
        const end = s.indexOf('?>', lt + 2);
        if (end < 0) {
          return pos;
        }
        pos = end + 2;
        continue;
      }
      const end = tagEnd(s, lt + 1);
      if (end < 0) {
        return pos;
      }
      if (next === 47) {
        handler.close?.(localName(s, lt + 2, nameEnd(s, lt + 2, end)));
      } else {
        const selfClosing = s.charCodeAt(end - 1) === 47;
        const stop = nameEnd(s, lt + 1, end);
        const name = localName(s, lt + 1, stop);
        if (name) {
          this.tag.set(s, name, stop, selfClosing ? end - 1 : end, selfClosing);
          handler.open?.(this.tag);
          if (selfClosing) {
            handler.close?.(name);
          }
        }
      }
      pos = end + 1;
    }
  }
}

/** Scans a whole document held in memory. */
export function scanXml(text: string, handler: XmlHandler, wantText = false): void {
  const scanner = new XmlScanner(handler);
  scanner.wantText = wantText;
  scanner.write(text);
  scanner.end();
}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

export function decodeEntities(text: string): string {
  if (text.indexOf('&') < 0) {
    return text;
  }
  return text.replace(/&(#x[0-9a-fA-F]{1,6}|#[0-9]{1,7}|amp|lt|gt|quot|apos);/g, (match, entity: string) => {
    if (entity.charCodeAt(0) !== 35) {
      return ENTITIES[entity] ?? match;
    }
    const code = entity.charCodeAt(1) === 120 ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
    return code > 0 && code <= 0x10ffff && (code < 0xd800 || code > 0xdfff) ? String.fromCodePoint(code) : '�';
  });
}

/** OOXML escapes characters XML cannot carry as `_xHHHH_` (e.g. `_x000D_` for a carriage return). */
export function decodeOoxmlEscapes(text: string): string {
  if (text.indexOf('_x') < 0) {
    return text;
  }
  return text.replace(/_x([0-9a-fA-F]{4})_/g, (_match, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

function isSpace(code: number): boolean {
  return code === 32 || code === 10 || code === 13 || code === 9;
}

/** Index of the `>` closing a tag that starts before `from`, skipping quoted attribute values; -1 if incomplete. */
function tagEnd(s: string, from: number): number {
  for (let i = from; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code === 62) {
      return i;
    }
    if (code === 34 || code === 39) {
      const close = s.indexOf(code === 34 ? '"' : "'", i + 1);
      if (close < 0) {
        return -1;
      }
      i = close;
    }
  }
  return -1;
}

/** End of a `<!...>` declaration, skipping a bracketed internal DTD subset. */
function declarationEnd(s: string, from: number): number {
  let depth = 0;
  for (let i = from; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code === 91) {
      depth++;
    } else if (code === 93) {
      depth = Math.max(0, depth - 1);
    } else if (code === 62 && depth === 0) {
      return i;
    }
  }
  return -1;
}

function nameEnd(s: string, from: number, to: number): number {
  let i = from;
  while (i < to) {
    const code = s.charCodeAt(i);
    if (isSpace(code) || code === 47 || code === 62) {
      break;
    }
    i++;
  }
  return i;
}

/** The name without its namespace prefix (only the name's own characters are looked at). */
function localName(s: string, from: number, to: number): string {
  let start = from;
  for (let i = from; i < to; i++) {
    if (s.charCodeAt(i) === 58) {
      start = i + 1;
    }
  }
  return s.slice(start, to);
}
