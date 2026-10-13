// Text bodies as HTML: real, selectable DOM text laid out by the browser inside the shape's text rectangle. Every
// value comes from the file, so text is only ever set with textContent and styles go through the CSSOM.
import { externalUrl } from '../../protocol.ts';
import { colorIn, hexToRgba, toCss, type ColorContext, type Rgba } from './color.ts';
import { fontFamily, resolveTypeface } from './fonts.ts';
import type { Rect } from './geometry.ts';
import { relationshipTarget, type RenderContext } from './renderContext.ts';
import type { ParagraphStyle, RunStyle, TextStyleResolver } from './textStyle.ts';
import { EMU_PER_PX, PX_PER_PT } from './units.ts';
import { attr, child, children, kids, num, relAttr } from './xml.ts';

export interface BodyProperties {
  /** CSS pixels. */
  readonly insets: { readonly l: number; readonly t: number; readonly r: number; readonly b: number };
  readonly anchor: string;
  readonly wrap: boolean;
  readonly vert: string;
  /** Degrees. */
  readonly rotation: number;
  readonly upright: boolean;
  readonly columns: number;
  readonly columnGap: number;
  /** normAutofit: font scale (1 = 100%) and line spacing reduction (0.2 = 20 points of percentage less). */
  readonly fontScale: number;
  readonly lineReduction: number;
}

const DEFAULT_INSETS = { lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720 };
const LINE_HEIGHT = 1.2;

/** Merges a:bodyPr from the most specific (the shape) to the least (its master placeholder). */
export function readBodyProperties(bodies: readonly (Element | undefined)[]): BodyProperties {
  const present = bodies.filter((body): body is Element => body !== undefined);
  const value = (name: string) => present.map((body) => attr(body, name)).find((item) => item !== undefined);
  const inset = (name: keyof typeof DEFAULT_INSETS) => {
    const raw = Number(value(name) ?? DEFAULT_INSETS[name]);
    return (Number.isFinite(raw) ? raw : DEFAULT_INSETS[name]) / EMU_PER_PX;
  };
  const autofit = present.map((body) => kids(body).find((node) => ['normAutofit', 'noAutofit', 'spAutoFit'].includes(node.localName))).find((item) => item !== undefined);
  const normal = autofit?.localName === 'normAutofit' ? autofit : undefined;
  return {
    insets: { l: inset('lIns'), t: inset('tIns'), r: inset('rIns'), b: inset('bIns') },
    anchor: value('anchor') ?? 't',
    wrap: value('wrap') !== 'none',
    vert: value('vert') ?? 'horz',
    rotation: (Number(value('rot')) || 0) / 60000,
    upright: value('upright') === '1',
    columns: Math.max(1, Math.min(16, Number(value('numCol')) || 1)),
    columnGap: (Number(value('spcCol')) || 0) / EMU_PER_PX,
    fontScale: Math.max(0.05, Math.min(1, (num(normal, 'fontScale') ?? 100000) / 100000)),
    lineReduction: Math.max(0, Math.min(0.9, (num(normal, 'lnSpcReduction') ?? 0) / 100000)),
  };
}

export interface TextOptions {
  /** Text rectangle inside the shape, in CSS pixels. */
  readonly box: Rect;
  readonly body: BodyProperties;
  readonly styles: TextStyleResolver;
  readonly context: RenderContext;
  /** Text is turned upside down on vertically flipped shapes (it is never mirrored). */
  readonly flipV?: boolean;
  /** The shape's rotation, which upright text undoes. */
  readonly shapeRotation?: number;
}

/** True when the body has any visible character, so empty placeholders draw nothing. */
export function hasText(txBody: Element | undefined): boolean {
  return children(txBody, 'p').some((p) => kids(p).some((node) => (node.localName === 'r' || node.localName === 'fld') && (child(node, 't')?.textContent ?? '') !== ''));
}

export function renderTextBody(txBody: Element, options: TextOptions): HTMLElement {
  const { box, body } = options;
  const container = document.createElement('div');
  container.className = `pptx-text pptx-anchor-${body.anchor === 'ctr' ? 'middle' : body.anchor === 'b' ? 'bottom' : 'top'}`;
  if (!body.wrap) {
    container.classList.add('pptx-nowrap');
  }
  const innerWidth = Math.max(0, box.w - body.insets.l - body.insets.r);
  const innerHeight = Math.max(0, box.h - body.insets.t - body.insets.b);
  const transforms: string[] = [];
  const sideways = body.vert === 'vert' || body.vert === 'vert270' || body.vert === 'wordArtVertRtl';
  // Sideways text is laid out in the box turned by 90°, centered on the real box, then rotated into place.
  const [width, height] = sideways ? [innerHeight, innerWidth] : [innerWidth, innerHeight];
  if (sideways) {
    transforms.push(`rotate(${body.vert === 'vert270' ? -90 : 90}deg)`);
  } else if (body.vert === 'eaVert' || body.vert === 'mongolianVert' || body.vert === 'wordArtVert') {
    container.classList.add(body.vert === 'wordArtVert' ? 'pptx-vertical-upright' : 'pptx-vertical');
  }
  if (body.rotation) {
    transforms.push(`rotate(${body.rotation}deg)`);
  }
  if (options.flipV) {
    transforms.push('rotate(180deg)');
  }
  if (body.upright && options.shapeRotation) {
    transforms.push(`rotate(${-options.shapeRotation}deg)`);
  }
  setBox(container, box.x + body.insets.l + (innerWidth - width) / 2, box.y + body.insets.t + (innerHeight - height) / 2, width, height);
  if (transforms.length > 0) {
    container.style.setProperty('transform', transforms.join(' '));
  }

  let flow: HTMLElement = container;
  if (body.columns > 1) {
    flow = document.createElement('div');
    flow.style.setProperty('column-count', String(body.columns));
    flow.style.setProperty('column-gap', px(body.columnGap));
    container.append(flow);
  }
  flow.append(...renderParagraphs(txBody, options));
  return container;
}

export type ParagraphOptions = Pick<TextOptions, 'body' | 'styles' | 'context'>;

/** The paragraphs of a text body, for a positioned text box or a table cell. */
export function renderParagraphs(txBody: Element, options: ParagraphOptions): HTMLElement[] {
  const counters = new Map<number, { scheme: string; value: number }>();
  const result: HTMLElement[] = [];
  for (const paragraph of children(txBody, 'p')) {
    if (options.context.budget.runs <= 0) {
      options.context.budget.truncated = true;
      break;
    }
    result.push(renderParagraph(paragraph, options, counters));
  }
  return result;
}

function renderParagraph(paragraph: Element, options: ParagraphOptions, counters: Map<number, { scheme: string; value: number }>): HTMLElement {
  const { styles, body, context } = options;
  const pPr = child(paragraph, 'pPr');
  const style = styles.paragraph(pPr);
  const content = kids(paragraph).filter((node) => node.localName === 'r' || node.localName === 'br' || node.localName === 'fld');
  const firstRun = content.find((node) => node.localName !== 'br');
  const leadStyle = styles.run(style.level, pPr, child(firstRun ?? paragraph, firstRun ? 'rPr' : 'endParaRPr'));
  const fontSize = leadStyle.size * body.fontScale * PX_PER_PT;

  const element = document.createElement('p');
  element.className = 'pptx-p';
  const css = element.style;
  css.setProperty('text-align', ALIGN[style.align] ?? 'left');
  if (!body.wrap) {
    css.setProperty('align-self', style.align === 'ctr' ? 'center' : style.align === 'r' ? 'flex-end' : 'flex-start');
  }
  if (style.align === 'dist') {
    css.setProperty('text-align-last', 'justify');
  }
  if (style.rtl) {
    element.dir = 'rtl';
  }
  css.setProperty(style.rtl ? 'padding-right' : 'padding-left', px(Math.max(0, style.marL / EMU_PER_PX)));
  css.setProperty('text-indent', px(style.indent / EMU_PER_PX));
  css.setProperty('font-size', px(fontSize));
  css.setProperty('line-height', lineHeight(style, body));
  css.setProperty('tab-size', px(style.defaultTab / EMU_PER_PX));
  const before = spacingPx(style.spaceBefore, fontSize);
  const after = spacingPx(style.spaceAfter, fontSize);
  if (before) {
    css.setProperty('margin-top', px(before));
  }
  if (after) {
    css.setProperty('margin-bottom', px(after));
  }

  const visible = content.some((node) => node.localName !== 'br' && (child(node, 't')?.textContent ?? '') !== '');
  updateCounters(counters, style, visible);
  if (visible && style.bullet.kind !== 'none') {
    element.append(bulletElement(style, leadStyle, options, counters.get(style.level)?.value ?? 1));
  }
  for (const node of content) {
    if (context.budget.runs-- <= 0) {
      context.budget.truncated = true;
      break;
    }
    if (node.localName === 'br') {
      element.append(document.createElement('br'));
      continue;
    }
    const rPr = child(node, 'rPr');
    const run = styles.run(style.level, pPr, rPr);
    let text = child(node, 't')?.textContent ?? '';
    if (node.localName === 'fld' && attr(node, 'type') === 'slidenum') {
      text = String(context.slideNumber);
    }
    if (text === '') {
      continue;
    }
    element.append(runElement(text, run, rPr, options));
  }
  const last = element.lastChild;
  if (!last || last.nodeName === 'BR' || (last instanceof HTMLElement && last.classList.contains('pptx-bullet'))) {
    // An empty paragraph (or a trailing line break) still takes a line, as in PowerPoint.
    element.append(document.createElement('br'));
  }
  return element;
}

const ALIGN: Readonly<Record<string, string>> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', justLow: 'justify', dist: 'justify', thaiDist: 'justify' };

function lineHeight(style: ParagraphStyle, body: BodyProperties): string {
  const spacing = style.lineSpacing;
  if (spacing?.pts !== undefined) {
    return px(spacing.pts * PX_PER_PT);
  }
  const percent = Math.max(0.1, (spacing?.pct ?? 1) - body.lineReduction);
  return String(Math.round(LINE_HEIGHT * percent * 1000) / 1000);
}

function spacingPx(spacing: { pct?: number; pts?: number } | undefined, fontSize: number): number {
  if (!spacing) {
    return 0;
  }
  return spacing.pts !== undefined ? spacing.pts * PX_PER_PT : (spacing.pct ?? 0) * fontSize * LINE_HEIGHT;
}

/** Auto-numbering restarts when a level's scheme changes or a paragraph at that level has no number. */
function updateCounters(counters: Map<number, { scheme: string; value: number }>, style: ParagraphStyle, visible: boolean): void {
  for (const level of [...counters.keys()]) {
    if (level > style.level) {
      counters.delete(level);
    }
  }
  if (style.bullet.kind !== 'autonum') {
    if (visible) {
      counters.delete(style.level);
    }
    return;
  }
  if (!visible) {
    return;
  }
  const scheme = style.bullet.scheme ?? 'arabicPeriod';
  const current = counters.get(style.level);
  counters.set(style.level, current && current.scheme === scheme ? { scheme, value: current.value + 1 } : { scheme, value: style.bullet.startAt ?? 1 });
}

function bulletElement(style: ParagraphStyle, lead: RunStyle, options: ParagraphOptions, number: number): HTMLElement {
  const { context, body } = options;
  const bullet = style.bullet;
  const span = document.createElement('span');
  span.className = 'pptx-bullet';
  const css = span.style;
  let text = bullet.kind === 'autonum' ? autoNumber(bullet.scheme ?? 'arabicPeriod', number) : bullet.kind === 'char' ? (bullet.char ?? '•') : '•';
  let family = runFamily(lead, context);
  const bulletFont = attr(bullet.font, 'typeface');
  if (bulletFont && bullet.kind === 'char') {
    const mapped = symbolCharacter(bulletFont, text);
    if (mapped) {
      text = mapped;
    } else {
      family = fontFamily(resolveTypeface(bulletFont, context.theme, lead.lang));
    }
  }
  span.textContent = text;
  const size = bullet.sizePts !== undefined ? bullet.sizePts : lead.size * (bullet.sizePct ?? 1);
  css.setProperty('font-size', px(size * body.fontScale * PX_PER_PT));
  css.setProperty('font-family', family);
  css.setProperty('color', toCss((bullet.color && colorIn(bullet.color, context.colors)) ?? runColor(lead, context.colors)));
  if (bullet.kind === 'autonum') {
    css.setProperty('font-weight', lead.bold ? '700' : '400');
  }
  if (style.indent < 0) {
    css.setProperty('min-width', px(-style.indent / EMU_PER_PX));
  } else {
    css.setProperty('padding-right', '0.25em');
  }
  return span;
}

function runElement(text: string, run: RunStyle, rPr: Element | undefined, options: ParagraphOptions): HTMLElement {
  const { context, body } = options;
  const span = document.createElement('span');
  span.textContent = text;
  const css = span.style;
  let size = run.size * body.fontScale * PX_PER_PT;
  if (run.baseline !== 0) {
    css.setProperty('vertical-align', px((size * run.baseline) / 100));
    size *= 2 / 3;
  }
  css.setProperty('font-size', px(size));
  css.setProperty('font-family', runFamily(run, context));
  if (run.bold) {
    css.setProperty('font-weight', '700');
  }
  if (run.italic) {
    css.setProperty('font-style', 'italic');
  }
  if (run.spacing) {
    css.setProperty('letter-spacing', px(run.spacing * PX_PER_PT));
  }
  if (run.cap === 'all') {
    css.setProperty('text-transform', 'uppercase');
  } else if (run.cap === 'small') {
    css.setProperty('font-variant', 'small-caps');
  }
  const highlight = run.highlight && colorIn(run.highlight, context.colors);
  if (highlight) {
    css.setProperty('background-color', toCss(highlight));
  }
  const link = hyperlink(child(rPr, 'hlinkClick'), context);
  let color = runColor(run, context.colors);
  let underline = run.underline;
  if (link) {
    color = hexToRgba(context.colors.scheme[context.colors.map['hlink'] ?? 'hlink']) ?? color;
    underline = underline === 'none' ? 'sng' : underline;
  }
  css.setProperty('color', toCss(color));
  const decorations = [underline !== 'none' ? 'underline' : '', run.strike !== 'noStrike' ? 'line-through' : ''].filter(Boolean);
  if (decorations.length > 0) {
    css.setProperty('text-decoration-line', decorations.join(' '));
    const decorationStyle = /dbl/i.test(underline) || run.strike === 'dblStrike' ? 'double' : /wavy/i.test(underline) ? 'wavy' : /dot/i.test(underline) ? 'dotted' : /dash/i.test(underline) ? 'dashed' : '';
    if (decorationStyle) {
      css.setProperty('text-decoration-style', decorationStyle);
    }
  }
  if (!link) {
    return span;
  }
  link.append(span);
  return link;
}

function runFamily(run: RunStyle, context: RenderContext): string {
  const theme = context.theme;
  const latin = resolveTypeface(run.latin ?? '+mn-lt', theme, run.lang);
  const eastAsian = resolveTypeface(run.ea ?? '+mn-ea', theme, run.lang) ?? resolveTypeface('+mn-ea', theme, 'ja-JP');
  return fontFamily(latin, eastAsian, resolveTypeface(run.cs, theme, run.lang));
}

export function runColor(run: RunStyle, colors: ColorContext): Rgba {
  const fill = run.fill;
  const fallback = hexToRgba(colors.scheme[colors.map['tx1'] ?? 'dk1']) ?? { r: 0, g: 0, b: 0, a: 1 };
  switch (fill?.localName) {
    case undefined:
    case 'blipFill':
    case 'grpFill':
      return fallback;
    case 'gradFill': {
      const stops = children(child(fill, 'gsLst'), 'gs');
      return colorIn(stops[0], colors) ?? fallback;
    }
    case 'pattFill':
      return colorIn(child(fill, 'fgClr'), colors) ?? fallback;
    case 'noFill':
      return { r: 0, g: 0, b: 0, a: 0 };
    default:
      // solidFill, or a style element holding the default color (fontRef, tcTxStyle).
      return colorIn(fill, colors) ?? fallback;
  }
}

/** An <a> for a run or shape hyperlink: http(s) URLs (opened by the shell through the host) or slide jumps. */
export function hyperlink(click: Element | undefined, context: RenderContext): HTMLAnchorElement | undefined {
  if (!click) {
    return undefined;
  }
  const action = attr(click, 'action') ?? '';
  const link = document.createElement('a');
  link.className = 'pptx-link';
  if (action.startsWith('ppaction://hlinkshowjump')) {
    const jump = /jump=(\w+)/.exec(action)?.[1];
    if (!jump || !['nextslide', 'previousslide', 'firstslide', 'lastslide'].includes(jump)) {
      return undefined;
    }
    link.dataset['slideJump'] = jump;
    return link;
  }
  const target = relationshipTarget(context, relAttr(click, 'id'));
  if (!target) {
    return undefined;
  }
  if (action.startsWith('ppaction://hlinksldjump')) {
    const index = context.deck.slideParts.indexOf(target.target);
    if (index < 0) {
      return undefined;
    }
    link.dataset['slideIndex'] = String(index);
    return link;
  }
  const url = target.external ? externalUrl(target.target) : null;
  if (!url) {
    return undefined;
  }
  link.href = url;
  link.title = url;
  return link;
}

// Bullet characters in the Wingdings and Symbol fonts, by their code in those fonts (or its U+F0xx alias).
const SYMBOL_FONTS: Readonly<Record<string, Readonly<Record<number, string>>>> = {
  wingdings: { 0x6c: '●', 0x6e: '■', 0x6f: '□', 0x71: '❑', 0x75: '◆', 0x76: '❖', 0xa7: '▪', 0xa8: '◻', 0xd8: '➢', 0xe0: '→', 0xe8: '➔', 0xfc: '✓', 0xfb: '✗', 0x9f: '•', 0xa1: '○', 0x77: '⬥' },
  symbol: { 0xb7: '•', 0xae: '→', 0xde: '⇒', 0xa8: '♦', 0xa7: '♣', 0xa9: '♥', 0xaa: '♠', 0x2d: '−', 0xd7: '·' },
};

export function symbolCharacter(typeface: string, text: string): string | undefined {
  const table = SYMBOL_FONTS[typeface.toLowerCase().replace(/\s+\d+$/, '')];
  const code = text.codePointAt(0);
  if (!table || code === undefined) {
    return undefined;
  }
  return table[code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code];
}

export function autoNumber(scheme: string, value: number): string {
  const n = Math.max(1, Math.trunc(value));
  let digits: string;
  if (scheme.startsWith('romanUc')) {
    digits = roman(n);
  } else if (scheme.startsWith('romanLc')) {
    digits = roman(n).toLowerCase();
  } else if (scheme.startsWith('alphaUc')) {
    digits = alpha(n);
  } else if (scheme.startsWith('alphaLc')) {
    digits = alpha(n).toLowerCase();
  } else if (scheme.startsWith('circleNum') && n <= 20) {
    return String.fromCodePoint((scheme.includes('WdBlack') ? 0x2775 : 0x245f) + n);
  } else if (scheme.startsWith('arabicDb') || scheme.startsWith('ea1Jpn') || scheme.startsWith('ea1Chs') || scheme.startsWith('ea1Cht')) {
    digits = String(n).replace(/\d/g, (digit) => String.fromCodePoint(0xff10 + Number(digit)));
    return scheme.endsWith('Period') ? `${digits}．` : digits;
  } else {
    digits = String(n);
  }
  if (scheme.endsWith('ParenBoth')) {
    return `(${digits})`;
  }
  if (scheme.endsWith('ParenR')) {
    return `${digits})`;
  }
  if (scheme.endsWith('Plain')) {
    return digits;
  }
  if (scheme.endsWith('Minus')) {
    return `- ${digits} -`;
  }
  return `${digits}.`;
}

function roman(n: number): string {
  const numerals: [number, string][] = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let rest = Math.min(n, 3999);
  let result = '';
  for (const [amount, letters] of numerals) {
    while (rest >= amount) {
      result += letters;
      rest -= amount;
    }
  }
  return result;
}

function alpha(n: number): string {
  const letter = String.fromCharCode(65 + ((n - 1) % 26));
  return letter.repeat(Math.min(10, Math.floor((n - 1) / 26) + 1));
}

export function setBox(element: HTMLElement | SVGElement, x: number, y: number, w: number, h: number): void {
  element.style.setProperty('left', px(x));
  element.style.setProperty('top', px(y));
  element.style.setProperty('width', px(w));
  element.style.setProperty('height', px(h));
}

export function px(value: number): string {
  return `${Math.round(value * 100) / 100}px`;
}
