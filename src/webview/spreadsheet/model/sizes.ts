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

/** Excel's row height for a font when the file does not say: 15pt for Calibri 11, 18.75pt for Yu Gothic 11. */
export function defaultRowHeightPt(font: Font): number {
  return (/游|yu gothic|meiryo|メイリオ/i.test(font.name) ? 18.75 : 15) * (font.size / 11);
}
