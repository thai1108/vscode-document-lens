// CSS hygiene for what docx-preview generates from the document. The values come from the file (colours, font names,
// lengths, VML style text), so a hostile file can inject declarations or whole rules. Everything is re-read through
// the browser's CSS parser: no value may load anything but blob:/data: URLs, lengths are clamped, and style rules are
// confined to the document's own subtree with @scope.
import { DOCX_LIMITS } from './limits.ts';

const POINTS_PER_UNIT: Record<string, number> = { pt: 1, px: 0.75, pc: 12, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, q: 72 / 101.6 };
const RELATIVE_UNITS = new Set(['em', 'rem', 'ex', 'ch', 'lh', 'vw', 'vh', 'vmin', 'vmax']);

/** Functions that can make the browser fetch something. */
const FETCHING_FUNCTION = /(?:^|[^\w-])(?:url|src|image|image-set|-webkit-image-set|cross-fade|-webkit-cross-fade|element|-moz-element)\(/i;
/** A url() argument or a quoted string (image-set() accepts plain strings as URLs). */
const URL_OR_STRING = /url\(\s*(?:"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'|([^)"'\s]*))\s*\)|"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'/gi;
/** A quoted string or url() (left alone), or a number with a unit. */
const LENGTH_TOKEN = /"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|url\([^)]*\)|(?<![\w.-])([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)([a-z]+|%)(?![\w-])/gi;
const PLAIN_NUMBER = /^\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)\s*$/i;

/** Declarations docx-preview never writes and a document has no business setting. */
const BLOCKED_PROPERTIES = new Set(['zoom', 'behavior', '-moz-binding']);

/** True when every URL the value could fetch is a blob: or data: URL. */
export function isSafeCssValue(value: string): boolean {
  if (!FETCHING_FUNCTION.test(value)) {
    return true;
  }
  for (const match of value.matchAll(URL_OR_STRING)) {
    const target = (match[1] ?? match[2] ?? match[3] ?? match[4] ?? match[5] ?? '').trim();
    if (!/^(?:blob|data):/i.test(target)) {
      return false;
    }
  }
  return true;
}

/** Clamps every length in a CSS value to DOCX_LIMITS; strings and url() arguments are left untouched. */
export function clampLengths(value: string): string {
  return value.replace(LENGTH_TOKEN, (token, number: string | undefined, unit: string | undefined) => {
    if (number === undefined || unit === undefined) {
      return token;
    }
    const max = maxForUnit(unit.toLowerCase());
    const amount = Number(number);
    return max === undefined || !(Math.abs(amount) > max) ? token : `${Math.sign(amount) * max}${unit}`;
  });
}

function maxForUnit(unit: string): number | undefined {
  const points = POINTS_PER_UNIT[unit];
  if (points !== undefined) {
    return Math.round((DOCX_LIMITS.maxLengthPt / points) * 100) / 100;
  }
  if (unit === '%') {
    return DOCX_LIMITS.maxPercentage;
  }
  return RELATIVE_UNITS.has(unit) ? DOCX_LIMITS.maxRelativeLength : undefined;
}

/** Applies the value rules to one declaration block (an element's inline style or a rule's style). */
export function sanitizeDeclarations(style: CSSStyleDeclaration): void {
  const names = Array.from({ length: style.length }, (_, i) => style.item(i));
  for (const name of names) {
    const value = style.getPropertyValue(name);
    // A custom property keeps its raw text, where an escape such as `u\72 l(` only becomes url( once substituted.
    const rawEscape = name.startsWith('--') && value.includes('\\');
    if (BLOCKED_PROPERTIES.has(name) || rawEscape || !isSafeCssValue(value)) {
      style.removeProperty(name);
      continue;
    }
    const clamped = clampProperty(name, value);
    if (clamped !== value) {
      style.setProperty(name, clamped, style.getPropertyPriority(name));
    }
  }
}

function clampProperty(name: string, value: string): string {
  const plain = PLAIN_NUMBER.exec(value)?.[1];
  if (plain !== undefined && (name === 'line-height' || name === 'column-count')) {
    const max = name === 'line-height' ? DOCX_LIMITS.maxLineHeight : DOCX_LIMITS.maxColumns;
    return Math.abs(Number(plain)) > max ? String(max) : value;
  }
  return clampLengths(value);
}

/**
 * Re-reads a style sheet docx-preview generated and returns a safe version: style rules (without nested rules) with
 * safe, clamped declarations wrapped in `@scope (<scope>)`, and @font-face rules whose sources are all blob:/data:.
 * @import, @media and every other at-rule are dropped; docx-preview does not generate them with the viewer's options.
 */
export function sanitizeStyleSheet(text: string, scope: string): string {
  const sheet = new window.CSSStyleSheet();
  try {
    sheet.replaceSync(text);
  } catch {
    return '';
  }
  const fonts: string[] = [];
  const rules: string[] = [];
  for (const rule of Array.from(sheet.cssRules)) {
    if (rule instanceof window.CSSStyleRule) {
      while (rule.cssRules?.length) {
        rule.deleteRule(0);
      }
      sanitizeDeclarations(rule.style);
      if (rule.style.length > 0) {
        rules.push(rule.cssText);
      }
    } else if (rule instanceof window.CSSFontFaceRule) {
      const source = rule.style.getPropertyValue('src');
      if (/url\(/i.test(source) && isSafeCssValue(source)) {
        sanitizeDeclarations(rule.style);
        fonts.push(rule.cssText);
      }
    }
  }
  const scoped = rules.length > 0 ? [`@scope (${scope}) {`, ...rules, '}'] : [];
  return [...fonts, ...scoped].join('\n');
}
