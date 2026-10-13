// One CSS class per cell format, created the first time a cell with that format is drawn. Rules live in a
// nonce-tagged <style> element (the webview's CSP refuses others) and are filled through the CSSOM, so values read
// from the file are only ever property values: an invalid one is dropped by the browser, it cannot add rules.
import type { BorderEdge, CellFormat, Font } from '../model/types.ts';

const BORDER_CSS: Record<BorderEdge['style'], [width: number, style: string]> = {
  thin: [1, 'solid'],
  hair: [1, 'dotted'],
  dotted: [1, 'dotted'],
  dashed: [1, 'dashed'],
  dashDot: [1, 'dashed'],
  dashDotDot: [1, 'dotted'],
  medium: [2, 'solid'],
  mediumDashed: [2, 'dashed'],
  mediumDashDot: [2, 'dashed'],
  mediumDashDotDot: [2, 'dotted'],
  slantDashDot: [2, 'dashed'],
  thick: [3, 'solid'],
  double: [3, 'double'],
};

const VERTICAL: Record<CellFormat['vertical'], string> = {
  top: 'flex-start',
  center: 'center',
  bottom: 'flex-end',
  justify: 'flex-start',
  distributed: 'center',
};

export class StyleRules {
  private readonly element: HTMLStyleElement;
  private readonly formats: readonly CellFormat[];
  private readonly scope: string;
  private readonly created = new Set<number>();

  constructor(document: Document, nonce: string, formats: readonly CellFormat[], scope: string) {
    this.formats = formats;
    this.scope = scope;
    this.element = document.createElement('style');
    this.element.nonce = nonce;
    this.element.setAttribute('nonce', nonce);
    document.head.append(this.element);
  }

  /** Class name for the cell format `index` (formats[0] for unknown indexes). */
  classFor(index: number): string {
    const format = this.formats[index] ? index : 0;
    if (!this.created.has(format)) {
      this.created.add(format);
      const sheet = this.element.sheet;
      if (sheet) {
        const rule = sheet.cssRules[sheet.insertRule(`${this.scope} .x${format} {}`, sheet.cssRules.length)] as CSSStyleRule | undefined;
        if (rule?.style) {
          applyFormat(rule.style, this.formats[format]);
        }
      }
    }
    return `x${format}`;
  }

  dispose(): void {
    this.element.remove();
  }
}

function applyFormat(style: CSSStyleDeclaration, format: CellFormat | undefined): void {
  if (!format) {
    return;
  }
  if (format.fill) {
    style.setProperty('background-color', format.fill);
  }
  applyFont(style, format.font);
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const edge = format.borders[side];
    if (edge) {
      style.setProperty(`border-${side}`, borderCss(edge));
    }
  }
  style.setProperty('align-items', VERTICAL[format.vertical]);
  if (format.wrap) {
    style.setProperty('white-space', 'pre-wrap');
    style.setProperty('overflow-wrap', 'anywhere');
  }
  if (format.indent) {
    style.setProperty('padding-left', `calc(var(--dl-zoom) * ${3 + format.indent * 9}px)`);
  }
}

/** A cell border edge as a CSS `border-*` value. */
export function borderCss(edge: BorderEdge): string {
  const [width, line] = BORDER_CSS[edge.style];
  return `${width}px ${line} ${edge.color}`;
}

/** Font properties for a cell format or a rich text run; only the ones `font` sets. */
export function applyFont(style: CSSStyleDeclaration, font: Partial<Font>): void {
  if (font.name !== undefined) {
    style.setProperty('font-family', fontFamily(font.name));
  }
  if (font.size !== undefined) {
    style.setProperty('font-size', `calc(var(--dl-zoom) * ${font.size}pt)`);
  }
  if (font.bold !== undefined) {
    style.setProperty('font-weight', font.bold ? '700' : '400');
  }
  if (font.italic !== undefined) {
    style.setProperty('font-style', font.italic ? 'italic' : 'normal');
  }
  if (font.color !== undefined) {
    style.setProperty('color', font.color);
  }
  if (font.underline !== undefined || font.strike !== undefined) {
    const lines = [font.underline && font.underline !== 'none' ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean);
    style.setProperty('text-decoration-line', lines.join(' ') || 'none');
    if (font.underline === 'double') {
      style.setProperty('text-decoration-style', 'double');
    }
  }
  if (font.vertAlign) {
    style.setProperty('vertical-align', font.vertAlign === 'superscript' ? 'super' : 'sub');
  }
}

const FALLBACKS: [RegExp, string][] = [
  [/明朝|mincho/i, '"Hiragino Mincho ProN", "Yu Mincho", "MS Mincho", serif'],
  [/ゴシック|gothic|meiryo|メイリオ|游|osaka/i, '"Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic", Meiryo, sans-serif'],
  [/courier|consolas|mono|menlo/i, 'monospace'],
  [/times|cambria|georgia|garamond|serif/i, 'serif'],
];

/** The file's font first (quoted, stripped of characters that could end the value), then a fallback that fits it. */
export function fontFamily(name: string): string {
  const clean = name.replace(/["\\;{}<>\p{Cc}]/gu, '').trim();
  const fallback = FALLBACKS.find(([pattern]) => pattern.test(clean))?.[1] ?? 'Calibri, Carlito, "Segoe UI", sans-serif';
  return clean ? `"${clean}", ${fallback}` : fallback;
}
