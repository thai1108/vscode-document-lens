// Excel's column and row units in CSS pixels at 100% zoom. Column widths are counted in "characters" of the default
// font's widest digit (ECMA-376 18.3.1.13); row heights are points.
import type { Font } from './types.ts';

/** Digit width as a share of the font size, for fonts common as a workbook's default font. */
const DIGIT_WIDTHS: [RegExp, number][] = [
  [/calibri|carlito/i, 0.507],
  [/游|yu gothic|yu mincho|meiryo|メイリオ/i, 0.55],
  [/ゴシック|明朝|gothic|mincho/i, 0.5],
  [/times/i, 0.5],
  [/verdana/i, 0.636],
  [/tahoma/i, 0.546],
];
const DEFAULT_DIGIT_WIDTH = 0.556;

/** Width of the widest digit of `font` in pixels: 7 for Calibri 11 or Arial 10, 8 for Yu Gothic 11, as in Excel. */
export function maxDigitWidth(font: Font): number {
  const ratio = DIGIT_WIDTHS.find(([pattern]) => pattern.test(font.name))?.[1] ?? DEFAULT_DIGIT_WIDTH;
  return Math.max(1, Math.round(pointsToPx(font.size) * ratio));
}

/** A `<col width>` (characters, padding included) in pixels. */
export function columnWidthPx(width: number, digitWidth: number): number {
  return Math.max(0, Math.trunc(((256 * width + Math.trunc(128 / digitWidth)) / 256) * digitWidth));
}

/** The default width of a sheet without `defaultColWidth`: base characters plus padding, rounded up to 8 pixels. */
export function baseColumnWidthPx(baseWidth: number, digitWidth: number): number {
  return Math.ceil((baseWidth * digitWidth + 5) / 8) * 8;
}

export function pointsToPx(points: number): number {
  return (points * 4) / 3;
}

/** Columns `first`..`last` (inclusive) and their size. */
export type SizeRange = readonly [first: number, last: number, size: number];

/**
 * The sizes ranges give entries 0..limit-1, later ranges winning where they overlap. Each entry is set once, however
 * many (full-width) ranges a file repeats.
 */
export function sizesFromRanges(ranges: readonly SizeRange[], limit: number): Map<number, number> {
  const sizes = new Float64Array(limit).fill(Number.NaN);
  // next[i] is the first entry at or after i not set yet (path-compressed as entries are set).
  const next = Int32Array.from({ length: limit + 1 }, (_value, i) => i);
  const find = (index: number): number => {
    let root = index;
    while (root < limit && next[root] !== root) {
      root = next[root] ?? limit;
    }
    for (let i = index; i < root; ) {
      const following = next[i] ?? limit;
      next[i] = root;
      i = following;
    }
    return root;
  };
  for (let k = ranges.length - 1; k >= 0; k--) {
    const [first, last, size] = ranges[k] ?? [0, -1, 0];
    for (let i = find(Math.max(0, first)); i <= Math.min(last, limit - 1); i = find(i + 1)) {
      sizes[i] = size;
      next[i] = i + 1;
    }
  }
  const result = new Map<number, number>();
  sizes.forEach((size, i) => {
    if (!Number.isNaN(size)) {
      result.set(i, size);
    }
  });
  return result;
}

/** Excel's row height for a font when the file does not say: 15pt for Calibri 11, 18.75pt for Yu Gothic 11. */
export function defaultRowHeightPt(font: Font): number {
  return (/游|yu gothic|meiryo|メイリオ/i.test(font.name) ? 18.75 : 15) * (font.size / 11);
}
