// Slide text: the characters of a TextCharsAtom/TextBytesAtom, split into paragraphs, with the indent level and
// bullet flag of each paragraph from the StyleTextPropAtom that follows it. Everything else about the text's
// formatting is ignored.

export interface Paragraph {
  readonly text: string;
  /** Outline level, 0–4. */
  readonly level: number;
  readonly bullet: boolean;
}

export type TextKind = 'title' | 'body' | 'other';

export interface TextBlock {
  readonly kind: TextKind;
  readonly paragraphs: readonly Paragraph[];
}

/** TextHeaderAtom.textType ([MS-PPT] TextTypeEnum). */
export const TEXT_TYPE = {
  title: 0,
  body: 1,
  notes: 2,
  other: 4,
  centerBody: 5,
  centerTitle: 6,
  halfBody: 7,
  quarterBody: 8,
} as const;

const MAX_LEVEL = 4;
const PARAGRAPH_BREAK = '\r';
const LINE_BREAK = '\v';
// C0/C1 control characters other than tab and newline; PowerPoint uses some of them as field markers.
// eslint-disable-next-line no-control-regex
const CONTROL_CHARACTERS = /[\u0000-\u0008\u000b-\u001f\u007f-\u009f￾￿]/g;

export function textKind(type: number): TextKind {
  switch (type) {
    case TEXT_TYPE.title:
    case TEXT_TYPE.centerTitle:
      return 'title';
    case TEXT_TYPE.body:
    case TEXT_TYPE.centerBody:
    case TEXT_TYPE.halfBody:
    case TEXT_TYPE.quarterBody:
      return 'body';
    default:
      return 'other';
  }
}

/** Body placeholders show bullets unless a paragraph turns them off; titles, subtitles and text boxes do not. */
function bulletByDefault(type: number): boolean {
  return type === TEXT_TYPE.body || type === TEXT_TYPE.halfBody || type === TEXT_TYPE.quarterBody;
}

/** TextCharsAtom: UTF-16LE. */
export function decodeChars(bytes: Uint8Array): string {
  return new TextDecoder('utf-16le').decode(bytes.subarray(0, bytes.byteLength & ~1));
}

/** TextBytesAtom: each byte is the low byte of a UTF-16 code unit whose high byte is zero (true ISO 8859-1). */
export function decodeBytes(bytes: Uint8Array): string {
  const chunk = 8192;
  let text = '';
  for (let i = 0; i < bytes.byteLength; i += chunk) {
    text += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return text;
}

export interface ParagraphRun {
  /** Characters the run covers, including each paragraph's break. */
  readonly count: number;
  readonly level: number;
  /** Set when the run states whether its paragraphs have bullets. */
  readonly bullet?: boolean;
}

// PFMasks bits ([MS-PPT] 2.9.20 TextPFException): which optional fields follow, in this order.
const PF = {
  bulletFlags: 0x000f,
  hasBullet: 0x0001,
  bulletChar: 0x0080,
  bulletFont: 0x0010,
  bulletSize: 0x0040,
  bulletColor: 0x0020,
  align: 0x0800,
  lineSpacing: 0x1000,
  spaceBefore: 0x2000,
  spaceAfter: 0x4000,
  leftMargin: 0x0100,
  indent: 0x0400,
  defaultTabSize: 0x8000,
  tabStops: 0x100000,
  fontAlign: 0x10000,
  wrapFlags: 0xe0000,
  textDirection: 0x200000,
} as const;

/** Optional fields as [mask, size in bytes]: those between bulletFlags and tabStops, then those after tabStops. */
const FIELDS_BEFORE_TABS: readonly [number, number][] = [
  [PF.bulletChar, 2],
  [PF.bulletFont, 2],
  [PF.bulletSize, 2],
  [PF.bulletColor, 4],
  [PF.align, 2],
  [PF.lineSpacing, 2],
  [PF.spaceBefore, 2],
  [PF.spaceAfter, 2],
  [PF.leftMargin, 2],
  [PF.indent, 2],
  [PF.defaultTabSize, 2],
];
const FIELDS_AFTER_TABS: readonly [number, number][] = [
  [PF.fontAlign, 2],
  [PF.wrapFlags, 2],
  [PF.textDirection, 2],
];

/**
 * The paragraph runs at the start of a StyleTextPropAtom (the character runs after them are not needed). The runs
 * together cover the text plus one character. A malformed atom just ends the list early: styles are optional here.
 */
export function paragraphRuns(bytes: Uint8Array, textLength: number): ParagraphRun[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const runs: ParagraphRun[] = [];
  let position = 0;
  let covered = 0;
  while (covered <= textLength && position + 10 <= bytes.byteLength) {
    const count = view.getUint32(position, true);
    const level = view.getUint16(position + 4, true);
    const masks = view.getUint32(position + 6, true);
    position += 10;
    let bullet: boolean | undefined;
    if (masks & PF.bulletFlags) {
      if (position + 2 > bytes.byteLength) {
        break;
      }
      if (masks & PF.hasBullet) {
        bullet = (view.getUint16(position, true) & 1) === 1;
      }
      position += 2;
    }
    position += fieldsSize(masks, FIELDS_BEFORE_TABS);
    if (masks & PF.tabStops) {
      if (position + 2 > bytes.byteLength) {
        break;
      }
      position += 2 + view.getUint16(position, true) * 4;
    }
    position += fieldsSize(masks, FIELDS_AFTER_TABS);
    if (position > bytes.byteLength || count === 0) {
      break;
    }
    runs.push({ count, level: Math.min(level, MAX_LEVEL), bullet });
    covered += count;
  }
  return runs;
}

function fieldsSize(masks: number, fields: readonly [number, number][]): number {
  return fields.reduce((size, [mask, bytes]) => size + (masks & mask ? bytes : 0), 0);
}

/** Splits text on paragraph breaks (\r); a vertical tab is a line break inside a paragraph. Empty paragraphs are dropped. */
export function toParagraphs(text: string, type: number, runs: readonly ParagraphRun[]): Paragraph[] {
  const title = textKind(type) === 'title';
  const paragraphs: Paragraph[] = [];
  let offset = 0;
  for (const raw of text.split(PARAGRAPH_BREAK)) {
    const run = runAt(runs, offset);
    offset += raw.length + 1;
    const clean = raw.replaceAll(LINE_BREAK, '\n').replace(CONTROL_CHARACTERS, '');
    if (!clean.trim()) {
      continue;
    }
    paragraphs.push({
      text: clean,
      level: title ? 0 : (run?.level ?? 0),
      bullet: !title && (run?.bullet ?? bulletByDefault(type)),
    });
  }
  return paragraphs;
}

function runAt(runs: readonly ParagraphRun[], offset: number): ParagraphRun | undefined {
  let end = 0;
  for (const run of runs) {
    end += run.count;
    if (offset < end) {
      return run;
    }
  }
  return runs.at(-1);
}
