// Turns run typefaces (including theme references such as +mn-lt) into a CSS font-family list. The latin face comes
// first and the East Asian face after it, so the browser's per-glyph fallback picks the East Asian face for CJK
// characters; common Office fonts get metric-compatible or look-alike substitutes for machines without them.
import type { Theme } from './theme.ts';

const SANS_JP = ['YuGothic', 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Noto Sans CJK JP', 'Meiryo'];
const SERIF_JP = ['YuMincho', 'Hiragino Mincho ProN', 'Noto Serif JP', 'Noto Serif CJK JP'];

const SUBSTITUTES: Record<string, readonly string[]> = {
  calibri: ['Carlito', 'Segoe UI', 'Helvetica Neue', 'Arial'],
  'calibri light': ['Carlito', 'Segoe UI Light', 'Helvetica Neue', 'Arial'],
  cambria: ['Caladea', 'Georgia', 'Times New Roman'],
  arial: ['Liberation Sans', 'Helvetica Neue', 'Helvetica'],
  'arial narrow': ['Liberation Sans Narrow', 'Helvetica Neue'],
  helvetica: ['Helvetica Neue', 'Arial', 'Liberation Sans'],
  'segoe ui': ['Helvetica Neue', 'Arial'],
  'times new roman': ['Liberation Serif', 'Times'],
  'courier new': ['Liberation Mono', 'Courier', 'Menlo'],
  consolas: ['Menlo', 'Monaco', 'Liberation Mono'],
  'century gothic': ['Avenir Next', 'Futura'],
  'yu gothic': SANS_JP,
  'yu gothic ui': SANS_JP,
  'yu gothic light': SANS_JP,
  'yu gothic medium': SANS_JP,
  '游ゴシック': SANS_JP,
  '游ゴシック light': SANS_JP,
  '游ゴシック medium': SANS_JP,
  meiryo: ['Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'YuGothic', 'Noto Sans JP'],
  'meiryo ui': ['Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'YuGothic', 'Noto Sans JP'],
  'メイリオ': ['Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'YuGothic', 'Noto Sans JP'],
  'ms gothic': ['Osaka', 'Hiragino Kaku Gothic ProN', 'Noto Sans JP'],
  'ms pgothic': ['Hiragino Kaku Gothic ProN', 'Hiragino Sans', 'Noto Sans JP'],
  'ms ui gothic': ['Hiragino Kaku Gothic ProN', 'Hiragino Sans', 'Noto Sans JP'],
  'ｍｓ ゴシック': ['Osaka', 'Hiragino Kaku Gothic ProN', 'Noto Sans JP'],
  'ｍｓ ｐゴシック': ['Hiragino Kaku Gothic ProN', 'Hiragino Sans', 'Noto Sans JP'],
  'ms mincho': SERIF_JP,
  'ms pmincho': SERIF_JP,
  'ｍｓ 明朝': SERIF_JP,
  'ｍｓ ｐ明朝': SERIF_JP,
  'yu mincho': SERIF_JP,
  '游明朝': SERIF_JP,
  'malgun gothic': ['Apple SD Gothic Neo', 'Noto Sans KR'],
  '맑은 고딕': ['Apple SD Gothic Neo', 'Noto Sans KR'],
  'microsoft yahei': ['PingFang SC', 'Hiragino Sans GB', 'Noto Sans SC'],
  '微软雅黑': ['PingFang SC', 'Hiragino Sans GB', 'Noto Sans SC'],
  dengxian: ['PingFang SC', 'Noto Sans SC'],
  '等线': ['PingFang SC', 'Noto Sans SC'],
  'microsoft jhenghei': ['PingFang TC', 'Noto Sans TC'],
  '微軟正黑體': ['PingFang TC', 'Noto Sans TC'],
  simsun: ['Songti SC', 'Noto Serif SC'],
  '宋体': ['Songti SC', 'Noto Serif SC'],
};

const SCRIPT_BY_LANGUAGE: Record<string, string> = { ja: 'Jpan', zh: 'Hans', ko: 'Hang' };

/** Resolves a theme font reference (+mj-lt, +mn-ea, ...) to a typeface name; other names are returned as is. */
export function resolveTypeface(typeface: string | undefined, theme: Theme, lang?: string): string | undefined {
  if (!typeface) {
    return undefined;
  }
  const match = /^\+(mj|mn)-(lt|ea|cs)$/.exec(typeface);
  if (!match) {
    return typeface;
  }
  const scheme = match[1] === 'mj' ? theme.major : theme.minor;
  if (match[2] === 'lt') {
    return scheme.latin || undefined;
  }
  if (match[2] === 'cs') {
    return scheme.cs || undefined;
  }
  if (scheme.ea) {
    return scheme.ea;
  }
  const language = lang?.toLowerCase() ?? '';
  const script = language === 'zh-tw' || language === 'zh-hk' ? 'Hant' : SCRIPT_BY_LANGUAGE[language.slice(0, 2)];
  return (script && scheme.scripts[script]) || undefined;
}

/** A CSS font-family value for the given faces, most specific first. */
export function fontFamily(...typefaces: (string | undefined)[]): string {
  const names: string[] = [];
  let serif = false;
  let mono = false;
  for (const typeface of typefaces) {
    const name = cleanName(typeface);
    if (!name) {
      continue;
    }
    names.push(name, ...(SUBSTITUTES[name.toLowerCase()] ?? []));
    serif ||= /times|mincho|明朝|georgia|cambria|garamond|serif|book antiqua|palatino|songti|simsun|宋体/i.test(name);
    mono ||= /courier|consolas|mono|menlo/i.test(name);
  }
  const unique = [...new Set(names)];
  return [...unique.map((name) => `"${name}"`), mono ? 'monospace' : serif ? 'serif' : 'sans-serif'].join(', ');
}

/** Typeface names come from the file: keep them to plain names so they cannot break out of the CSS value. */
function cleanName(typeface: string | undefined): string | undefined {
  const name = typeface
    ?.replace(/["\\;{}<>]/g, '')
    .replace(/\p{Cc}/gu, '')
    .trim()
    .slice(0, 80);
  return name || undefined;
}
