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

/** How far the search for the end of the token at the start of the buffer got in the previous chunk. */
interface Resume {
  /** Characters of the buffer already searched. */
  searched: number;
  /** Quote character the search stopped inside (an attribute value), or 0. */
  quote: number;
  /** Bracket depth of a declaration's internal subset. */
  depth: number;
}

export class XmlScanner {
  /** Whether character data is decoded and passed to the handler; off by default, which skips it cheaply. */
  wantText = false;
  private readonly handler: XmlHandler;
  private readonly tag = new XmlTag();
  private buffer = '';
  /** Set when a chunk ends inside a token, so the next chunk continues the search instead of starting over. */
  private resume: Resume | undefined;

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
    this.resume = undefined;
  }

  /** Processes complete tokens and returns how many characters were consumed. */
  private scan(s: string): number {
    const handler = this.handler;
    const resume = this.resume;
    this.resume = undefined;
    let pos = 0;
    for (;;) {
      // Text pending from the previous chunk had no '<' in the part already searched.
      const lt = s.indexOf('<', pos === 0 && resume && s.charCodeAt(0) !== 60 ? resume.searched : pos);
      if (lt < 0) {
        if (!this.wantText) {
          return s.length;
        }
        // Text up to the end may continue in the next chunk.
        return this.stop(s, pos, 0, 0);
      }
      if (lt > pos && this.wantText) {
        handler.text?.(decodeEntities(s.slice(pos, lt)));
      }
      pos = lt;
      // The token the previous chunk ended in, if this is it.
      const again = lt === 0 && resume && s.charCodeAt(0) === 60 ? resume : undefined;
      const searched = again?.searched ?? 0;
      const next = s.charCodeAt(lt + 1);
      if (next === 33) {
        // <!-- -->, <![CDATA[ ]]> or a declaration such as <!DOCTYPE ...>.
        if (s.startsWith('<!--', lt)) {
          const end = s.indexOf('-->', Math.max(lt + 4, searched - 2));
          if (end < 0) {
            return this.stop(s, pos, 0, 0);
          }
          pos = end + 3;
        } else if (s.startsWith('<![CDATA[', lt)) {
          const end = s.indexOf(']]>', Math.max(lt + 9, searched - 2));
          if (end < 0) {
            return this.stop(s, pos, 0, 0);
          }
          if (this.wantText) {
            handler.text?.(s.slice(lt + 9, end));
          }
          pos = end + 3;
        } else {
          const end = this.declarationEnd(s, lt, Math.max(lt + 2, searched), again?.depth ?? 0);
          if (end < 0) {
            return pos;
          }
          pos = end + 1;
        }
        continue;
      }
      if (next === 63) {
        const end = s.indexOf('?>', Math.max(lt + 2, searched - 1));
        if (end < 0) {
          return this.stop(s, pos, 0, 0);
        }
        pos = end + 2;
        continue;
      }
      const end = this.tagEnd(s, lt, Math.max(lt + 1, searched), again?.quote ?? 0);
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

  /** Keeps the unfinished token from `pos` for the next chunk, remembering that all of `s` was searched. */
  private stop(s: string, pos: number, quote: number, depth: number): number {
    this.resume = { searched: s.length - pos, quote, depth };
    return pos;
  }

  /**
   * Index of the `>` closing the tag at `start`, skipping quoted attribute values; -1 (with the
   * search state kept) if it is not in `s` yet. `quote` is the quote an earlier search stopped inside.
   */
  private tagEnd(s: string, start: number, from: number, quote: number): number {
    let i = from;
    if (quote) {
      const close = s.indexOf(quote === 34 ? '"' : "'", i);
      if (close < 0) {
        this.stop(s, start, quote, 0);
        return -1;
      }
      i = close + 1;
    }
    for (; i < s.length; i++) {
      const code = s.charCodeAt(i);
      if (code === 62) {
        return i;
      }
      if (code === 34 || code === 39) {
        const close = s.indexOf(code === 34 ? '"' : "'", i + 1);
        if (close < 0) {
          this.stop(s, start, code, 0);
          return -1;
        }
        i = close;
      }
    }
    this.stop(s, start, 0, 0);
    return -1;
  }

  /** End of a `<!...>` declaration, skipping a bracketed internal DTD subset. */
  private declarationEnd(s: string, start: number, from: number, depth: number): number {
    let level = depth;
    for (let i = from; i < s.length; i++) {
      const code = s.charCodeAt(i);
      if (code === 91) {
        level++;
      } else if (code === 93) {
        level = Math.max(0, level - 1);
      } else if (code === 62 && level === 0) {
        return i;
      }
    }
    this.stop(s, start, 0, level);
    return -1;
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
