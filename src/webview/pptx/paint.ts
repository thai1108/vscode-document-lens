// Fills, outlines and shadows, read from DrawingML and expressed as SVG paint (presentation attributes and <defs>;
// never style attributes, which the webview's CSP refuses).
import { colorIn, toCss, toHex, transformColor, type ColorContext, type Rgba } from './color.ts';
import type { Picture } from './media.ts';
import { relationshipTarget, type RenderContext } from './renderContext.ts';
import type { Theme } from './theme.ts';
import { EMU_PER_PX } from './units.ts';
import { attr, child, children, find, kids, num, relAttr } from './xml.ts';

const SVG_NS = 'http://www.w3.org/2000/svg';

export interface GradientStop {
  readonly offset: number;
  readonly color: Rgba;
}

export interface Insets {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

export type Paint =
  | { readonly kind: 'none' }
  | { readonly kind: 'solid'; readonly color: Rgba }
  | {
      readonly kind: 'gradient';
      readonly stops: readonly GradientStop[];
      /** Degrees, clockwise from the x axis (linear gradients). */
      readonly angle: number;
      readonly radial: boolean;
      /** Center of a radial gradient as fractions of the box. */
      readonly cx: number;
      readonly cy: number;
    }
  | {
      readonly kind: 'image';
      readonly picture: Picture;
      readonly srcRect?: Insets;
      readonly fillRect?: Insets;
      /** Tiling: scale factors and offset in px; absent means stretch. */
      readonly tile?: { readonly sx: number; readonly sy: number; readonly tx: number; readonly ty: number };
      readonly opacity: number;
      readonly grayscale: boolean;
    }
  | { readonly kind: 'pattern'; readonly preset: string; readonly fg: Rgba; readonly bg: Rgba };

export const NO_PAINT: Paint = { kind: 'none' };

export interface Box {
  readonly w: number;
  readonly h: number;
}

let sequence = 0;

/** Document-wide unique ids: thumbnails and the main view render the same slide into one document. */
function uniqueId(prefix: string): string {
  sequence += 1;
  return `pptx-${prefix}-${sequence}`;
}

export function svg<K extends keyof SVGElementTagNameMap>(tag: K, attributes: Record<string, string | number | undefined> = {}): SVGElementTagNameMap[K] {
  const element = document.createElementNS(SVG_NS, tag);
  for (const [name, value] of Object.entries(attributes)) {
    if (value !== undefined) {
      element.setAttribute(name, String(value));
    }
  }
  return element;
}

function insets(element: Element | undefined): Insets | undefined {
  if (!element) {
    return undefined;
  }
  const part = (name: string) => (num(element, name) ?? 0) / 100000;
  return { l: part('l'), t: part('t'), r: part('r'), b: part('b') };
}

/** Reads a fill element (noFill, solidFill, gradFill, blipFill, pattFill); grpFill and unknown fills give undefined. */
export async function readFill(fill: Element | undefined, colors: ColorContext, context: RenderContext): Promise<Paint | undefined> {
  switch (fill?.localName) {
    case 'noFill':
      return NO_PAINT;
    case 'solidFill': {
      const color = colorIn(fill, colors);
      return color ? { kind: 'solid', color } : NO_PAINT;
    }
    case 'gradFill':
      return readGradient(fill, colors);
    case 'blipFill':
      return readBlipFill(fill, context);
    case 'pattFill': {
      const fg = colorIn(child(fill, 'fgClr'), colors) ?? { r: 0, g: 0, b: 0, a: 1 };
      const bg = colorIn(child(fill, 'bgClr'), colors) ?? { r: 255, g: 255, b: 255, a: 1 };
      return { kind: 'pattern', preset: attr(fill, 'prst') ?? 'pct50', fg, bg };
    }
    default:
      return undefined;
  }
}

function readGradient(fill: Element, colors: ColorContext): Paint {
  const stops = children(child(fill, 'gsLst'), 'gs')
    .map((gs) => ({ offset: Math.max(0, Math.min(1, (num(gs, 'pos') ?? 0) / 100000)), color: colorIn(gs, colors) }))
    .filter((stop): stop is GradientStop => stop.color !== undefined)
    .sort((a, b) => a.offset - b.offset);
  if (stops.length === 0) {
    return NO_PAINT;
  }
  const path = child(fill, 'path');
  const focus = insets(child(path, 'fillToRect')) ?? { l: 0.5, t: 0.5, r: 0.5, b: 0.5 };
  return {
    kind: 'gradient',
    stops,
    angle: (num(child(fill, 'lin'), 'ang') ?? 0) / 60000,
    radial: path !== undefined,
    cx: (focus.l + (1 - focus.r)) / 2,
    cy: (focus.t + (1 - focus.b)) / 2,
  };
}

/** A picture fill (blipFill of a shape, picture or background). */
export async function readBlipFill(fill: Element, context: RenderContext): Promise<Paint> {
  const blip = child(fill, 'blip');
  const picture = await pictureOf(blip, context);
  const tile = child(fill, 'tile');
  const amount = num(child(blip, 'alphaModFix'), 'amt');
  return {
    kind: 'image',
    picture,
    srcRect: insets(child(fill, 'srcRect')),
    fillRect: tile ? undefined : (insets(find(fill, 'stretch', 'fillRect')) ?? { l: 0, t: 0, r: 0, b: 0 }),
    tile: tile
      ? {
          sx: (num(tile, 'sx') ?? 100000) / 100000,
          sy: (num(tile, 'sy') ?? 100000) / 100000,
          tx: (num(tile, 'tx') ?? 0) / EMU_PER_PX,
          ty: (num(tile, 'ty') ?? 0) / EMU_PER_PX,
        }
      : undefined,
    opacity: amount === undefined ? 1 : Math.max(0, Math.min(1, amount / 100000)),
    grayscale: child(blip, 'grayscl') !== undefined,
  };
}

/** The picture a:blip points at: its SVG version when there is one, never an external (linked) file. */
async function pictureOf(blip: Element | undefined, context: RenderContext): Promise<Picture> {
  if (!blip) {
    return { unsupported: 'Picture' };
  }
  const svgBlip = kids(child(blip, 'extLst'))
    .map((ext) => child(ext, 'svgBlip'))
    .find((element) => element !== undefined);
  for (const candidate of [svgBlip, blip]) {
    const target = relationshipTarget(context, relAttr(candidate, 'embed'));
    if (target && !target.external) {
      const picture = await context.media.picture(target.target);
      if ('url' in picture || candidate === blip) {
        return picture;
      }
    }
  }
  if (relAttr(blip, 'link')) {
    return { unsupported: 'Linked picture (not loaded)' };
  }
  return { unsupported: 'Missing picture' };
}

/** The theme style a style reference (fillRef, lnRef, effectRef, bgRef) points at, with phClr set to its color. */
export function styleReference(reference: Element | undefined, list: readonly Element[], colors: ColorContext): { element: Element; colors: ColorContext } | undefined {
  const idx = num(reference, 'idx') ?? 0;
  const element = list[idx - 1];
  if (!reference || idx <= 0 || !element) {
    return undefined;
  }
  return { element, colors: { ...colors, placeholder: colorIn(reference, colors) } };
}

/** fillRef / bgRef: 1-999 index the theme's fill styles, 1001+ its background fill styles. */
function themeFillList(theme: Theme, reference: Element | undefined): { list: readonly Element[]; offset: number } {
  const idx = num(reference, 'idx') ?? 0;
  return idx >= 1001 ? { list: theme.bgFillStyles, offset: 1000 } : { list: theme.fillStyles, offset: 0 };
}

export async function styleFill(reference: Element | undefined, theme: Theme, colors: ColorContext, context: RenderContext): Promise<Paint | undefined> {
  if (!reference) {
    return undefined;
  }
  const { list, offset } = themeFillList(theme, reference);
  const idx = (num(reference, 'idx') ?? 0) - offset;
  const style = list[idx - 1];
  if (idx <= 0 || !style) {
    return NO_PAINT;
  }
  return readFill(style, { ...colors, placeholder: colorIn(reference, colors) }, context);
}

/** Path fill modifiers of preset geometries (the shaded faces of a cube, the top of a can). */
export function modifyPaint(paint: Paint, mode: string): Paint {
  const factor: Record<string, [string, number]> = {
    darken: ['shade', 60000],
    darkenLess: ['shade', 80000],
    lighten: ['tint', 60000],
    lightenLess: ['tint', 80000],
  };
  const change = factor[mode];
  if (!change) {
    return paint;
  }
  const transform = (color: Rgba) => transformColor(color, change[0], change[1]);
  if (paint.kind === 'solid') {
    return { kind: 'solid', color: transform(paint.color) };
  }
  if (paint.kind === 'gradient') {
    return { ...paint, stops: paint.stops.map((stop) => ({ offset: stop.offset, color: transform(stop.color) })) };
  }
  return paint;
}

/**
 * Sets `fill` or `stroke` (and its opacity) on an SVG element for `paint` over a box of `box` px, adding gradient,
 * pattern or picture definitions to `defs`.
 */
export function applyPaint(target: SVGElement, property: 'fill' | 'stroke', paint: Paint | undefined, defs: SVGDefsElement, box: Box): void {
  if (!paint || paint.kind === 'none') {
    target.setAttribute(property, 'none');
    return;
  }
  if (paint.kind === 'solid') {
    target.setAttribute(property, toHex(paint.color));
    if (paint.color.a < 1) {
      target.setAttribute(`${property}-opacity`, String(round(paint.color.a)));
    }
    return;
  }
  const id = uniqueId(paint.kind);
  defs.append(paintServer(id, paint, box));
  target.setAttribute(property, `url(#${id})`);
}

function paintServer(id: string, paint: Exclude<Paint, { kind: 'none' | 'solid' }>, box: Box): SVGElement {
  if (paint.kind === 'gradient') {
    const gradient = paint.radial ? radialGradient(id, paint, box) : linearGradient(id, paint.angle, box);
    for (const stop of paint.stops) {
      gradient.append(svg('stop', { offset: round(stop.offset), 'stop-color': toHex(stop.color), 'stop-opacity': stop.color.a < 1 ? round(stop.color.a) : undefined }));
    }
    return gradient;
  }
  if (paint.kind === 'pattern') {
    return patternTile(id, paint.preset, paint.fg, paint.bg);
  }
  return imagePattern(id, paint, box);
}

function linearGradient(id: string, angle: number, box: Box): SVGLinearGradientElement {
  const radians = (angle * Math.PI) / 180;
  const dx = Math.cos(radians);
  const dy = Math.sin(radians);
  const half = (Math.abs(box.w * dx) + Math.abs(box.h * dy)) / 2;
  const cx = box.w / 2;
  const cy = box.h / 2;
  return svg('linearGradient', {
    id,
    gradientUnits: 'userSpaceOnUse',
    x1: round(cx - dx * half),
    y1: round(cy - dy * half),
    x2: round(cx + dx * half),
    y2: round(cy + dy * half),
  });
}

function radialGradient(id: string, paint: { cx: number; cy: number }, box: Box): SVGRadialGradientElement {
  const cx = paint.cx * box.w;
  const cy = paint.cy * box.h;
  const radius = Math.max(...[0, box.w].flatMap((x) => [0, box.h].map((y) => Math.hypot(x - cx, y - cy))));
  return svg('radialGradient', { id, gradientUnits: 'userSpaceOnUse', cx: round(cx), cy: round(cy), r: round(Math.max(radius, 1)) });
}

function imagePattern(id: string, paint: Extract<Paint, { kind: 'image' }>, box: Box): SVGPatternElement {
  if (!('url' in paint.picture)) {
    return svg('pattern', { id, width: 1, height: 1 });
  }
  if (paint.tile) {
    const width = Math.max(1, (paint.picture.width ?? 64) * paint.tile.sx);
    const height = Math.max(1, (paint.picture.height ?? 64) * paint.tile.sy);
    const pattern = svg('pattern', { id, patternUnits: 'userSpaceOnUse', x: round(paint.tile.tx), y: round(paint.tile.ty), width: round(width), height: round(height) });
    pattern.append(svg('image', { href: paint.picture.url, width: round(width), height: round(height), preserveAspectRatio: 'none', opacity: paint.opacity < 1 ? round(paint.opacity) : undefined }));
    return pattern;
  }
  const rect = imageRect(box, paint.srcRect, paint.fillRect);
  const pattern = svg('pattern', { id, patternUnits: 'userSpaceOnUse', x: 0, y: 0, width: round(Math.max(box.w, 0.01)), height: round(Math.max(box.h, 0.01)) });
  pattern.append(svg('image', { href: paint.picture.url, x: round(rect.x), y: round(rect.y), width: round(rect.w), height: round(rect.h), preserveAspectRatio: 'none', opacity: paint.opacity < 1 ? round(paint.opacity) : undefined }));
  return pattern;
}

/**
 * Where the whole picture goes so that its srcRect-cropped part fills the fillRect-inset box (negative values extend
 * beyond it).
 */
function imageRect(box: Box, srcRect: Insets = { l: 0, t: 0, r: 0, b: 0 }, fillRect: Insets = { l: 0, t: 0, r: 0, b: 0 }): { x: number; y: number; w: number; h: number } {
  const tx = box.w * fillRect.l;
  const ty = box.h * fillRect.t;
  const tw = box.w * (1 - fillRect.l - fillRect.r);
  const th = box.h * (1 - fillRect.t - fillRect.b);
  const visibleW = Math.max(0.001, 1 - srcRect.l - srcRect.r);
  const visibleH = Math.max(0.001, 1 - srcRect.t - srcRect.b);
  const w = tw / visibleW;
  const h = th / visibleH;
  return { x: tx - srcRect.l * w, y: ty - srcRect.t * h, w, h };
}

// 8×8 ordered-dither thresholds: the first N cells of this order are set for an N/64 dot pattern.
const BAYER = [0, 32, 8, 40, 2, 34, 10, 42, 48, 16, 56, 24, 50, 18, 58, 26, 12, 44, 4, 36, 14, 46, 6, 38, 60, 28, 52, 20, 62, 30, 54, 22, 3, 35, 11, 43, 1, 33, 9, 41, 51, 19, 59, 27, 49, 17, 57, 25, 15, 47, 7, 39, 13, 45, 5, 37, 63, 31, 55, 23, 61, 29, 53, 21];

function patternCell(preset: string): (x: number, y: number) => boolean {
  const percent = /^pct(\d+)$/.exec(preset);
  if (percent) {
    const count = Math.round((64 * Number(percent[1])) / 100);
    return (x, y) => (BAYER[y * 8 + x] ?? 64) < count;
  }
  const lines: Record<string, (x: number, y: number) => boolean> = {
    horz: (_x, y) => y % 4 === 0,
    ltHorz: (_x, y) => y % 4 === 0,
    narHorz: (_x, y) => y % 2 === 0,
    dkHorz: (_x, y) => y % 4 < 2,
    dashHorz: (x, y) => (y === 0 && x < 4) || (y === 4 && x >= 4),
    vert: (x) => x % 4 === 0,
    ltVert: (x) => x % 4 === 0,
    narVert: (x) => x % 2 === 0,
    dkVert: (x) => x % 4 < 2,
    dashVert: (x, y) => (x === 0 && y < 4) || (x === 4 && y >= 4),
    smGrid: (x, y) => x % 4 === 0 || y % 4 === 0,
    lgGrid: (x, y) => x === 0 || y === 0,
    cross: (x, y) => x === 0 || y === 0,
    dotGrid: (x, y) => (x === 0 && y % 2 === 0) || (y === 0 && x % 2 === 0),
    dnDiag: (x, y) => (x - y + 8) % 4 === 0,
    ltDnDiag: (x, y) => (x - y + 8) % 4 === 0,
    dkDnDiag: (x, y) => (x - y + 8) % 4 < 2,
    wdDnDiag: (x, y) => (x - y + 8) % 8 < 3,
    dashDnDiag: (x, y) => (x - y + 8) % 4 === 0 && y % 4 < 2,
    upDiag: (x, y) => (x + y) % 4 === 0,
    ltUpDiag: (x, y) => (x + y) % 4 === 0,
    dkUpDiag: (x, y) => (x + y) % 4 < 2,
    wdUpDiag: (x, y) => (x + y) % 8 < 3,
    dashUpDiag: (x, y) => (x + y) % 4 === 0 && y % 4 < 2,
    diagCross: (x, y) => (x - y + 8) % 4 === 0 || (x + y) % 4 === 0,
    smCheck: (x, y) => ((x >> 1) + (y >> 1)) % 2 === 0,
    lgCheck: (x, y) => ((x >> 2) + (y >> 2)) % 2 === 0,
    openDmnd: (x, y) => (x - y + 8) % 8 === 0 || (x + y) % 8 === 0,
    solidDmnd: (x, y) => Math.abs(x - 3.5) + Math.abs(y - 3.5) <= 3,
    horzBrick: (x, y) => y % 4 === 0 || (y < 4 ? x === 0 : x === 4),
    zigZag: (x, y) => y === (x < 4 ? x : 7 - x) % 4,
  };
  return lines[preset] ?? ((x, y) => (BAYER[y * 8 + x] ?? 64) < 16);
}

function patternTile(id: string, preset: string, fg: Rgba, bg: Rgba): SVGPatternElement {
  const pattern = svg('pattern', { id, patternUnits: 'userSpaceOnUse', width: 8, height: 8 });
  pattern.append(svg('rect', { width: 8, height: 8, fill: toHex(bg), 'fill-opacity': bg.a < 1 ? round(bg.a) : undefined }));
  const cell = patternCell(preset);
  const cells: string[] = [];
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) {
      if (cell(x, y)) {
        cells.push(`M${x} ${y}h1v1h-1Z`);
      }
    }
  }
  if (cells.length > 0) {
    pattern.append(svg('path', { d: cells.join(''), fill: toHex(fg), 'fill-opacity': fg.a < 1 ? round(fg.a) : undefined }));
  }
  return pattern;
}

export interface LineEnd {
  readonly type: string;
  readonly width: string;
  readonly length: string;
}

export interface LineStyle {
  readonly paint: Paint;
  /** CSS pixels. */
  readonly width: number;
  readonly dash?: readonly number[];
  readonly cap: 'butt' | 'round' | 'square';
  readonly join: 'round' | 'bevel' | 'miter';
  readonly head?: LineEnd;
  readonly tail?: LineEnd;
  /** sng, dbl, thickThin, thinThick or tri. */
  readonly compound?: string;
}

/** One a:ln (from the shape, its placeholder chain or the theme via lnRef) with the colors it resolves against. */
export interface LineSource {
  readonly ln: Element;
  readonly colors: ColorContext;
}

const LINE_FILLS = ['noFill', 'solidFill', 'gradFill', 'pattFill'];

const DASHES: Readonly<Record<string, readonly number[]>> = {
  dot: [1, 1],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** Merges line properties from the most specific source down; no fill anywhere means no line. */
export async function readLine(sources: readonly LineSource[], context: RenderContext): Promise<LineStyle | undefined> {
  const withChild = (names: readonly string[]) => {
    for (const source of sources) {
      const found = kids(source.ln).find((node) => names.includes(node.localName));
      if (found) {
        return { element: found, colors: source.colors };
      }
    }
    return undefined;
  };
  const withAttr = (name: string) => sources.map((source) => attr(source.ln, name)).find((value) => value !== undefined);
  const fill = withChild(LINE_FILLS);
  if (!fill) {
    return undefined;
  }
  let paint = (await readFill(fill.element, fill.colors, context)) ?? NO_PAINT;
  if (paint.kind === 'pattern') {
    paint = { kind: 'solid', color: paint.fg };
  }
  if (paint.kind === 'none') {
    return { paint, width: 0, cap: 'butt', join: 'round' };
  }
  const widthEmu = Number(withAttr('w') ?? 9525);
  const width = Math.max(Number.isFinite(widthEmu) ? widthEmu / EMU_PER_PX : 1, 0.75);
  const dashElement = withChild(['prstDash', 'custDash'])?.element;
  let dash: number[] | undefined;
  if (dashElement?.localName === 'prstDash') {
    dash = DASHES[attr(dashElement, 'val') ?? 'solid']?.map((value) => value * Math.max(width, 1));
  } else if (dashElement?.localName === 'custDash') {
    dash = children(dashElement, 'ds')
      .slice(0, 32)
      .flatMap((ds) => [((num(ds, 'd') ?? 100000) / 100000) * width, ((num(ds, 'sp') ?? 100000) / 100000) * width]);
  }
  const cap = withAttr('cap');
  const join = withChild(['round', 'bevel', 'miter'])?.element.localName;
  const end = (name: string): LineEnd | undefined => {
    const element = withChild([name])?.element;
    const type = attr(element, 'type') ?? 'none';
    return element && type !== 'none' ? { type, width: attr(element, 'w') ?? 'med', length: attr(element, 'len') ?? 'med' } : undefined;
  };
  return {
    paint,
    width,
    dash,
    cap: cap === 'rnd' ? 'round' : cap === 'sq' ? 'square' : 'butt',
    join: join === 'bevel' ? 'bevel' : join === 'miter' ? 'miter' : 'round',
    head: end('headEnd'),
    tail: end('tailEnd'),
    compound: withAttr('cmpd'),
  };
}

// Line and gap widths across a compound line, outside to inside and symmetric (thickThin/thinThick are drawn as
// equal double lines: a mask can only cut bands centered on the path).
const COMPOUND_BANDS: Readonly<Record<string, readonly number[]>> = {
  dbl: [1, 1, 1],
  thickThin: [1, 1, 1],
  thinThick: [1, 1, 1],
  tri: [1, 1, 1, 1, 1],
};
/** Each line of a compound outline stays at least this wide (px), as PowerPoint keeps them visible. */
const MIN_COMPOUND_BAND = 0.75;

export function applyLine(target: SVGElement, line: LineStyle | undefined, defs: SVGDefsElement, box: Box, open: boolean): void {
  if (!line || line.paint.kind === 'none' || line.width <= 0) {
    target.setAttribute('stroke', 'none');
    return;
  }
  applyPaint(target, 'stroke', line.paint, defs, box);
  const bands = line.compound ? COMPOUND_BANDS[line.compound] : undefined;
  const path = target.getAttribute('d');
  if (bands && path) {
    target.setAttribute('stroke-width', String(round(compoundMask(target, path, bands, line.width, defs, box))));
  } else {
    target.setAttribute('stroke-width', String(round(line.width)));
  }
  target.setAttribute('stroke-linecap', line.cap);
  target.setAttribute('stroke-linejoin', line.join);
  if (line.dash) {
    target.setAttribute('stroke-dasharray', line.dash.map(round).join(' '));
  }
  if (open) {
    const color = line.paint.kind === 'solid' ? line.paint.color : line.paint.kind === 'gradient' ? (line.paint.stops[0]?.color ?? BLACK) : BLACK;
    if (line.head) {
      target.setAttribute('marker-start', `url(#${lineEndMarker(defs, line.head, color, line.width, true)})`);
    }
    if (line.tail) {
      target.setAttribute('marker-end', `url(#${lineEndMarker(defs, line.tail, color, line.width, false)})`);
    }
  }
}

/** Masks the gaps out of a stroke as wide as the whole compound line; returns that width. */
function compoundMask(target: SVGElement, path: string, bands: readonly number[], width: number, defs: SVGDefsElement, box: Box): number {
  const units = bands.reduce((sum, band) => sum + band, 0);
  const unit = Math.max(width / units, MIN_COMPOUND_BAND);
  const id = uniqueId('compound');
  const margin = units * unit + 4;
  const mask = svg('mask', { id, maskUnits: 'userSpaceOnUse', x: round(-margin), y: round(-margin), width: round(box.w + margin * 2), height: round(box.h + margin * 2) });
  mask.append(svg('rect', { x: round(-margin), y: round(-margin), width: round(box.w + margin * 2), height: round(box.h + margin * 2), fill: 'white' }));
  let inner = units;
  for (const [index, band] of bands.entries()) {
    inner -= band * 2;
    if (inner <= 0) {
      break;
    }
    mask.append(svg('path', { d: path, fill: 'none', stroke: index % 2 === 0 ? 'black' : 'white', 'stroke-width': round(inner * unit) }));
  }
  defs.append(mask);
  target.setAttribute('mask', `url(#${id})`);
  return units * unit;
}

const BLACK: Rgba = { r: 0, g: 0, b: 0, a: 1 };
const END_SIZES: Readonly<Record<string, number>> = { sm: 2, med: 3, lg: 5 };

function lineEndMarker(defs: SVGDefsElement, end: LineEnd, color: Rgba, lineWidth: number, start: boolean): string {
  const id = uniqueId('arrow');
  // Arrowheads scale with the line, but stay visible on hairlines.
  const base = Math.max(lineWidth, 1.5);
  const width = (END_SIZES[end.width] ?? 3) * base;
  const length = (END_SIZES[end.length] ?? 3) * base;
  const centered = end.type === 'oval' || end.type === 'diamond';
  const marker = svg('marker', {
    id,
    viewBox: '0 0 10 10',
    markerUnits: 'userSpaceOnUse',
    markerWidth: round(length),
    markerHeight: round(width),
    refX: centered ? 5 : 9,
    refY: 5,
    orient: start ? 'auto-start-reverse' : 'auto',
    preserveAspectRatio: 'none',
  });
  const fill = toHex(color);
  const opacity = color.a < 1 ? round(color.a) : undefined;
  switch (end.type) {
    case 'arrow':
      marker.append(svg('path', { d: 'M1 1L9 5L1 9', fill: 'none', stroke: fill, 'stroke-opacity': opacity, 'stroke-width': 1.5, 'stroke-linejoin': 'miter' }));
      break;
    case 'stealth':
      marker.append(svg('path', { d: 'M0 0L10 5L0 10L3 5Z', fill, 'fill-opacity': opacity }));
      break;
    case 'diamond':
      marker.append(svg('path', { d: 'M0 5L5 0L10 5L5 10Z', fill, 'fill-opacity': opacity }));
      break;
    case 'oval':
      marker.append(svg('ellipse', { cx: 5, cy: 5, rx: 5, ry: 5, fill, 'fill-opacity': opacity }));
      break;
    default:
      marker.append(svg('path', { d: 'M0 0L10 5L0 10Z', fill, 'fill-opacity': opacity }));
  }
  defs.append(marker);
  return id;
}

/** An outer shadow as a CSS drop-shadow() filter. */
export function shadowFilter(effectList: Element | undefined, colors: ColorContext): string | undefined {
  const shadow = child(effectList, 'outerShdw');
  if (!shadow) {
    return undefined;
  }
  const color = colorIn(shadow, colors) ?? { r: 0, g: 0, b: 0, a: 0.4 };
  const distance = (num(shadow, 'dist') ?? 0) / EMU_PER_PX;
  const direction = ((num(shadow, 'dir') ?? 0) / 60000) * (Math.PI / 180);
  const blur = (num(shadow, 'blurRad') ?? 0) / EMU_PER_PX;
  return `drop-shadow(${round(distance * Math.cos(direction))}px ${round(distance * Math.sin(direction))}px ${round(blur / 2)}px ${toCss(color)})`;
}

export function round(value: number): number {
  return Math.round(value * 100) / 100;
}
