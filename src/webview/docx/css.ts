// CSS hygiene for what docx-preview generates from the document. The values come from the file (colours, font names,
// lengths, VML style text), so a hostile file can inject declarations or whole rules. Everything is re-read through
// the browser's CSS parser and handled as the declarations it serializes to (style.cssText): a shorthand holding a
// var() shows up there, while its longhands report empty values. No value may load anything but blob:/data: URLs,
// lengths are clamped, var() may only refer to docx-preview's own --docx-* properties (which may only hold colours,
// names or blob: URLs), style rules are confined to the document with @scope, and embedded fonts get names of their
// own so they cannot stand in for fonts the viewer itself uses.
import { DOCX_LIMITS } from './limits.ts';

/** Embedded font families (lower case, unquoted) to the names their @font-face rules are given. */
export type FontNames = ReadonlyMap<string, string>;

const POINTS_PER_UNIT: Record<string, number> = { pt: 1, px: 0.75, pc: 12, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, q: 72 / 101.6 };
const RELATIVE_UNITS = new Set(['em', 'rem', 'ex', 'ch', 'lh', 'vw', 'vh', 'vmin', 'vmax']);

/** Functions that can make the browser fetch something. */
const FETCHING_FUNCTION = /(?:^|[^\w-])(?:url|src|image|image-set|-webkit-image-set|cross-fade|-webkit-cross-fade|element|-moz-element)\(/i;
/** A url() argument or a quoted string (image-set() accepts plain strings as URLs). */
const URL_OR_STRING = /url\(\s*(?:"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'|([^)"'\s]*))\s*\)|"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'/gi;
/** A quoted string or url() (left alone), or a number with a unit. */
const LENGTH_TOKEN = /"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|url\([^)]*\)|(?<![\w.-])([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)([a-z]+|%)(?![\w-])/gi;
const PLAIN_NUMBER = /^\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)\s*$/i;
/** Substitutions resolved only when the value is used, out of reach of the checks here. */
const SUBSTITUTION = /(?:^|[^\w-])(?:var|env|attr)\(/i;
const OTHER_SUBSTITUTION = /(?:^|[^\w-])(?:env|attr)\(/i;
/** Math functions, whose products and quotients of numbers can make any length. */
const MATH_FUNCTION = /(?:^|[^\w-])(?:calc|min|max|clamp|round|mod|rem|pow|sqrt|hypot|log|exp|abs|sign|sin|cos|tan|asin|acos|atan|atan2)\(/gi;
/**
 * What docx-preview's --docx-* properties hold: a hex colour, a blob:/data: URL (numbering pictures) or a list of
 * names (theme fonts, colour keywords). A name starts with a letter, so it can never pass for a number or a length.
 */
const CUSTOM_PROPERTY_VALUE = /^\s*(?:#[0-9a-f]{3,8}|url\(\s*(?:"[^"\\]*"|[^)"'\\\s]*)\s*\)|[\p{L}_][\p{L}\p{M}\p{N} _.,-]*)\s*$/iu;

/** Declarations docx-preview never writes and a document has no business setting. */
const BLOCKED_PROPERTIES = new Set(['zoom', 'behavior', '-moz-binding']);
/** The only descriptors kept in an embedded font's @font-face rule. */
const FONT_FACE_DESCRIPTORS = ['font-weight', 'font-style'];

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

interface Declaration {
  name: string;
  value: string;
  important: boolean;
}

/**
 * The value to keep for one declaration, or undefined to drop it. Exported for tests; sanitizeDeclarations applies it
 * to a whole declaration block.
 */
export function sanitizeDeclaration(name: string, value: string, fonts: FontNames = new Map()): string | undefined {
  const property = name.toLowerCase();
  if (BLOCKED_PROPERTIES.has(property) || OTHER_SUBSTITUTION.test(value) || !isSafeCssValue(value) || !safeSubstitutions(value)) {
    return undefined;
  }
  if (property.startsWith('--')) {
    if (!property.startsWith('--docx-') || !CUSTOM_PROPERTY_VALUE.test(value)) {
      return undefined;
    }
    return property.endsWith('-font') ? renameFonts(value, fonts) : value;
  }
  if (hasMathProduct(value)) {
    return undefined;
  }
  const plain = PLAIN_NUMBER.exec(value)?.[1];
  if (plain !== undefined && (property === 'line-height' || property === 'column-count')) {
    const max = property === 'line-height' ? DOCX_LIMITS.maxLineHeight : DOCX_LIMITS.maxColumns;
    return Math.abs(Number(plain)) > max ? String(max) : value;
  }
  const clamped = clampLengths(value);
  return property === 'font-family' ? renameFonts(clamped, fonts) : clamped;
}

/** var() may only refer to docx-preview's own --docx-* properties, without a fallback (which could hold anything). */
function safeSubstitutions(value: string): boolean {
  if (!SUBSTITUTION.test(value)) {
    return true;
  }
  for (const match of value.matchAll(/var\(/gi)) {
    const inner = functionArguments(value, (match.index ?? 0) + match[0].length);
    if (!/^\s*--docx-[\w-]*\s*$/i.test(inner)) {
      return false;
    }
  }
  return true;
}

function hasMathProduct(value: string): boolean {
  for (const match of value.matchAll(MATH_FUNCTION)) {
    if (/[*/]/.test(functionArguments(value, (match.index ?? 0) + match[0].length))) {
      return true;
    }
  }
  return false;
}

/** The text between an opening parenthesis (just before `start`) and its closing one. */
function functionArguments(value: string, start: number): string {
  let depth = 1;
  let quote = '';
  for (let index = start; index < value.length; index++) {
    const char = value[index];
    if (quote) {
      if (char === '\\') {
        index++;
      } else if (char === quote) {
        quote = '';
      }
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === '(') {
      depth++;
    } else if (char === ')' && --depth === 0) {
      return value.slice(start, index);
    }
  }
  return value.slice(start);
}

/** Splits serialized declarations (style.cssText) at the semicolons outside strings and brackets. */
function parseDeclarations(cssText: string): Declaration[] {
  const declarations: Declaration[] = [];
  let depth = 0;
  let quote = '';
  let start = 0;
  const add = (text: string) => {
    const colon = text.indexOf(':');
    if (colon <= 0) {
      return;
    }
    const important = /!\s*important\s*$/i.test(text);
    const value = text.slice(colon + 1).replace(/!\s*important\s*$/i, '').trim();
    declarations.push({ name: text.slice(0, colon).trim(), value, important });
  };
  for (let index = 0; index < cssText.length; index++) {
    const char = cssText[index];
    if (quote) {
      if (char === '\\') {
        index++;
      } else if (char === quote) {
        quote = '';
      }
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === '(' || char === '[' || char === '{') {
      depth++;
    } else if (char === ')' || char === ']' || char === '}') {
      depth--;
    } else if (char === ';' && depth === 0) {
      add(cssText.slice(start, index));
      start = index + 1;
    }
  }
  add(cssText.slice(start));
  return declarations;
}

/** Applies sanitizeDeclaration to a declaration block (an element's inline style or a rule's style). */
export function sanitizeDeclarations(style: CSSStyleDeclaration, fonts: FontNames = new Map()): void {
  let changed = false;
  const kept: string[] = [];
  for (const { name, value, important } of parseDeclarations(style.cssText)) {
    const safe = sanitizeDeclaration(name, value, fonts);
    if (safe === undefined) {
      changed = true;
      continue;
    }
    changed ||= safe !== value;
    kept.push(`${name}: ${safe}${important ? ' !important' : ''}`);
  }
  if (changed) {
    style.cssText = kept.join('; ');
  }
}

/** Replaces embedded font families in a font-family list with their own names. */
export function renameFonts(list: string, fonts: FontNames): string {
  if (fonts.size === 0) {
    return list;
  }
  return splitTopLevel(list, ',')
    .map((family) => fonts.get(familyKey(family)) ?? family.trim())
    .join(', ');
}

/**
 * The embedded fonts declared by docx-preview's style sheets (@font-face rules with blob:/data: sources), each family
 * mapped to a name of the viewer's own: `dl-docx-font-<n>`.
 */
export function embeddedFontNames(sheets: readonly string[]): FontNames {
  const names = new Map<string, string>();
  for (const text of sheets) {
    for (const rule of parseRules(text)) {
      if (rule instanceof window.CSSFontFaceRule && safeFontSource(rule.style.getPropertyValue('src'))) {
        const key = familyKey(rule.style.getPropertyValue('font-family'));
        if (key && !names.has(key)) {
          names.set(key, `dl-docx-font-${names.size + 1}`);
        }
      }
    }
  }
  return names;
}

/**
 * Re-reads a style sheet docx-preview generated and returns a safe version: style rules (without nested rules) with
 * safe declarations wrapped in `@scope (<scope>)`, and @font-face rules of embedded fonts, renamed after `fonts` and
 * reduced to their family, source, weight and style. @import, @media and every other at-rule are dropped;
 * docx-preview does not generate them with the viewer's options.
 */
export function sanitizeStyleSheet(text: string, scope: string, fonts: FontNames = new Map()): string {
  const faces: string[] = [];
  const rules: string[] = [];
  for (const rule of parseRules(text)) {
    if (rule instanceof window.CSSStyleRule) {
      while (rule.cssRules?.length) {
        rule.deleteRule(0);
      }
      sanitizeDeclarations(rule.style, fonts);
      if (rule.style.length > 0) {
        rules.push(rule.cssText);
      }
    } else if (rule instanceof window.CSSFontFaceRule) {
      const source = rule.style.getPropertyValue('src');
      const family = fonts.get(familyKey(rule.style.getPropertyValue('font-family')));
      if (family && safeFontSource(source)) {
        const descriptors = FONT_FACE_DESCRIPTORS.map((name) => [name, rule.style.getPropertyValue(name)] as const)
          .filter(([, value]) => /^[\w\s%.-]+$/.test(value))
          .map(([name, value]) => `${name}: ${value}`);
        faces.push(`@font-face { ${[`font-family: ${family}`, `src: ${source}`, ...descriptors].join('; ')}; }`);
      }
    }
  }
  const scoped = rules.length > 0 ? [`@scope (${scope}) {`, ...rules, '}'] : [];
  return [...faces, ...scoped].join('\n');
}

function parseRules(text: string): CSSRule[] {
  const sheet = new window.CSSStyleSheet();
  try {
    sheet.replaceSync(text);
  } catch {
    return [];
  }
  return Array.from(sheet.cssRules);
}

function safeFontSource(source: string): boolean {
  return /url\(/i.test(source) && isSafeCssValue(source) && !SUBSTITUTION.test(source);
}

/** A font family as CSS matches it: unquoted, case-insensitively. */
function familyKey(family: string): string {
  const trimmed = family.trim();
  const unquoted = /^(["']).*\1$/s.test(trimmed) ? trimmed.slice(1, -1).replace(/\\(.)/g, '$1') : trimmed.replace(/\s+/g, ' ');
  return unquoted.toLowerCase();
}

function splitTopLevel(text: string, separator: string): string[] {
  const parts: string[] = [];
  let quote = '';
  let depth = 0;
  let start = 0;
  for (let index = 0; index < text.length; index++) {
    const char = text[index];
    if (quote) {
      if (char === '\\') {
        index++;
      } else if (char === quote) {
        quote = '';
      }
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === '(') {
      depth++;
    } else if (char === ')') {
      depth--;
    } else if (char === separator && depth === 0) {
      parts.push(text.slice(start, index));
      start = index + 1;
    }
  }
  parts.push(text.slice(start));
  return parts;
}
