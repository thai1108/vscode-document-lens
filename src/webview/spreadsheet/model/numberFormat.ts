// Number formats: Excel format codes applied to cell values. SheetJS's SSF does the formatting; this module adds the
// built-in format table, Japanese era and weekday codes SSF does not know, and the color a section asks for.
import { SSF } from 'xlsx';

/**
 * Built-in formats by id (ECMA-376 18.8.30). Ids 27-36 and 50-58 depend on the locale Excel ran in; they are only
 * written by East Asian Excel, and the Japanese meaning is used. Ids 14 and 22 follow the viewer's locale, like Excel.
 */
const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  27: '[$-411]ge.m.d',
  28: '[$-411]ggge"年"m"月"d"日"',
  29: '[$-411]ggge"年"m"月"d"日"',
  30: 'm/d/yy',
  31: 'yyyy"年"m"月"d"日"',
  32: 'h"時"mm"分"',
  33: 'h"時"mm"分"ss"秒"',
  34: 'yyyy"年"m"月"',
  35: 'm"月"d"日"',
  36: '[$-411]ge.m.d',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
  50: '[$-411]ge.m.d',
  51: '[$-411]ggge"年"m"月"d"日"',
  52: 'yyyy"年"m"月"',
  53: 'm"月"d"日"',
  54: '[$-411]ggge"年"m"月"d"日"',
  55: 'yyyy"年"m"月"',
  56: 'm"月"d"日"',
  57: '[$-411]ge.m.d',
  58: '[$-411]ggge"年"m"月"d"日"',
};

export function builtinFormat(id: number, locale = defaultLocale()): string {
  if (id === 14) {
    return locale.startsWith('ja') ? 'yyyy/m/d' : 'm/d/yyyy';
  }
  if (id === 22) {
    return locale.startsWith('ja') ? 'yyyy/m/d h:mm' : 'm/d/yyyy h:mm';
  }
  return BUILTIN_FORMATS[id] ?? 'General';
}

function defaultLocale(): string {
  return typeof navigator === 'undefined' ? 'en' : navigator.language;
}

export interface FormattedValue {
  text: string;
  /** Color from a `[Red]`-style section code. */
  color?: string;
}

const isDateCache = new Map<string, boolean>();

export function isDateFormat(format: string): boolean {
  const cached = isDateCache.get(format);
  if (cached !== undefined) {
    return cached;
  }
  let result = hasJapaneseDateCodes(format);
  try {
    result ||= SSF.is_date(format) === true;
  } catch {
    // An unreadable format is not a date format.
  }
  if (isDateCache.size > 10_000) {
    isDateCache.clear();
  }
  isDateCache.set(format, result);
  return result;
}

export function formatNumber(value: number, format: string, date1904: boolean): FormattedValue {
  const color = sectionColor(format, value);
  let code = hasJapaneseDateCodes(format) ? expandJapaneseDateCodes(format, value, date1904) : format;
  // A locale tag without a currency symbol ([$-411]) only selects the calendar; SSF would print its "$".
  code = code.replace(/\[\$-[0-9A-Za-z-]+\]/g, '');
  if (code.includes('.') && isDateFormat(format)) {
    // In a date a dot is a separator (2024.1.15) unless it starts fractional seconds (ss.00); SSF only knows the latter.
    code = code.replace(/"[^"]*"|\\.|\[[^\]]*\]|\.(?!0)/g, (token) => (token === '.' ? '"."' : token));
  }
  let text: string;
  try {
    text = SSF.format(code, value, { date1904 });
  } catch {
    text = generalNumber(value);
  }
  return color ? { text, color } : { text };
}

/** Text cells only use a format's fourth section (or a lone `@` section). */
export function formatText(text: string, format: string): FormattedValue {
  const sections = splitSections(format);
  const section = sections.length >= 4 ? sections[3] : sections.length === 1 && sections[0]?.includes('@') ? sections[0] : undefined;
  if (section === undefined) {
    return { text };
  }
  const color = colorCode(section);
  let formatted = text;
  try {
    formatted = SSF.format(format, text);
  } catch {
    // Keep the plain text.
  }
  return color ? { text: formatted, color } : { text: formatted };
}

export function generalNumber(value: number): string {
  try {
    return SSF.format('General', value);
  } catch {
    return String(value);
  }
}

/** Date parts of a serial date, or undefined if it is outside Excel's calendar. */
export function dateParts(value: number, date1904: boolean): { y: number; m: number; d: number; H: number; M: number; S: number; q: number } | undefined {
  try {
    const parts = SSF.parse_date_code(value, { date1904 }) as { y: number; m: number; d: number; H: number; M: number; S: number; q: number } | null;
    return parts ?? undefined;
  } catch {
    return undefined;
  }
}

/** ISO-like text for the cell bar: `2024-01-15`, `2024-01-15 13:45:00` or `13:45:00`. */
export function isoDate(value: number, date1904: boolean): string | undefined {
  const parts = dateParts(value, date1904);
  if (!parts) {
    return undefined;
  }
  const pad = (n: number) => String(n).padStart(2, '0');
  const date = `${parts.y}-${pad(parts.m)}-${pad(parts.d)}`;
  const time = `${pad(parts.H)}:${pad(parts.M)}:${pad(parts.S)}`;
  if (value >= 0 && value < 1) {
    return time;
  }
  return parts.H || parts.M || parts.S ? `${date} ${time}` : date;
}

/** Splits a format code at `;` outside quotes, brackets and escapes. */
export function splitSections(format: string): string[] {
  const sections: string[] = [];
  let start = 0;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '"') {
      const close = format.indexOf('"', i + 1);
      i = close < 0 ? format.length : close;
    } else if (ch === '[') {
      const close = format.indexOf(']', i + 1);
      i = close < 0 ? format.length : close;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      i++;
    } else if (ch === ';') {
      sections.push(format.slice(start, i));
      start = i + 1;
    }
  }
  sections.push(format.slice(start));
  return sections;
}

const NAMED_COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
};

/** The color code of the section Excel picks for `value`, e.g. red for negatives in `#,##0;[Red]-#,##0`. */
export function sectionColor(format: string, value: number): string | undefined {
  if (format.indexOf('[') < 0) {
    return undefined;
  }
  const sections = splitSections(format);
  const conditions = sections.map(sectionCondition);
  let section: string | undefined;
  if (conditions.some((condition) => condition !== undefined)) {
    section = sections.find((_text, i) => conditions[i]?.(value)) ?? sections.find((_text, i) => !conditions[i] && i < 3);
  } else if (sections.length === 1 || (sections.length === 2 && value >= 0) || (sections.length >= 3 && value > 0)) {
    section = sections[0];
  } else if (value < 0) {
    section = sections[1];
  } else {
    section = sections[2];
  }
  return section === undefined ? undefined : colorCode(section);
}

function colorCode(section: string): string | undefined {
  for (const [, code] of section.matchAll(/\[([^\]]*)\]/g)) {
    const name = (code ?? '').trim().toLowerCase();
    const named = NAMED_COLORS[name];
    if (named) {
      return named;
    }
    const indexed = /^colou?r\s*(\d{1,2})$/.exec(name);
    if (indexed) {
      return PALETTE_COLORS[Number(indexed[1]) - 1];
    }
  }
  return undefined;
}

/** `[Color1]`-`[Color56]` refer to the 56-color palette. */
const PALETTE_COLORS = [
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080',
  '#c0c0c0', '#808080', '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff', '#000080', '#ff00ff', '#ffff00', '#00ffff',
  '#800080', '#800000', '#008080', '#0000ff', '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99', '#3366ff', '#33cccc',
  '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696', '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

function sectionCondition(section: string): ((value: number) => boolean) | undefined {
  const match = /\[(<=|>=|<>|<|>|=)\s*(-?[0-9.]+(?:[eE][-+]?\d+)?)\s*\]/.exec(section);
  if (!match) {
    return undefined;
  }
  const limit = Number(match[2]);
  switch (match[1]) {
    case '<':
      return (value) => value < limit;
    case '<=':
      return (value) => value <= limit;
    case '>':
      return (value) => value > limit;
    case '>=':
      return (value) => value >= limit;
    case '<>':
      return (value) => value !== limit;
    default:
      return (value) => value === limit;
  }
}

// Japanese calendar codes: g/gg/ggg era name, e/ee era year, aaa/aaaa weekday. SSF reads `e` as the Gregorian year
// and does not know the others, so they are replaced by quoted literals before SSF sees the format.
const ERAS = [
  { start: [2019, 5, 1], letter: 'R', short: '令', name: '令和' },
  { start: [1989, 1, 8], letter: 'H', short: '平', name: '平成' },
  { start: [1926, 12, 25], letter: 'S', short: '昭', name: '昭和' },
  { start: [1912, 7, 30], letter: 'T', short: '大', name: '大正' },
  { start: [1868, 1, 1], letter: 'M', short: '明', name: '明治' },
] as const;
const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];

const japaneseCodesCache = new Map<string, boolean>();

function hasJapaneseDateCodes(format: string): boolean {
  let result = japaneseCodesCache.get(format);
  if (result === undefined) {
    const bare = stripLiterals(format).replace(/general/gi, '');
    const japaneseLocale = /\[\$-(?:411|ja-JP)\]/i.test(format);
    result = /g/i.test(bare) || /aaa/.test(bare) || (japaneseLocale && /e/i.test(bare) && /[dmy]/i.test(bare) && !/e[+-]/i.test(bare));
    if (japaneseCodesCache.size > 10_000) {
      japaneseCodesCache.clear();
    }
    japaneseCodesCache.set(format, result);
  }
  return result;
}

function stripLiterals(format: string): string {
  return format.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '');
}

function expandJapaneseDateCodes(format: string, value: number, date1904: boolean): string {
  const parts = dateParts(value, date1904);
  if (!parts) {
    return format;
  }
  const era = ERAS.find(({ start: [y, m, d] }) => parts.y > y || (parts.y === y && (parts.m > m || (parts.m === m && parts.d >= d)))) ?? ERAS[ERAS.length - 1];
  const eraYear = era ? parts.y - era.start[0] + 1 : parts.y;
  const weekday = WEEKDAYS[Math.floor(value + (date1904 ? 1462 : 0) + 6) % 7] ?? '';
  const literal = (text: string) => `"${text}"`;
  return format.replace(/"[^"]*"|\\.|\[[^\]]*\]|g{1,3}|G{1,3}|e{1,2}|E{1,2}|a{3,4}/g, (token) => {
    const lower = token.toLowerCase();
    if (lower.startsWith('g')) {
      return literal(lower.length === 1 ? (era?.letter ?? '') : lower.length === 2 ? (era?.short ?? '') : (era?.name ?? ''));
    }
    if (lower.startsWith('e')) {
      return literal(lower.length === 2 ? String(eraYear).padStart(2, '0') : String(eraYear));
    }
    if (lower.startsWith('aaa')) {
      return literal(lower.length === 4 ? `${weekday}曜日` : weekday);
    }
    return token;
  });
}
