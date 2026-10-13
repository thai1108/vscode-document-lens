// Effective paragraph and run properties. A property comes from the most specific place that sets it:
//   run rPr → paragraph pPr → the shape's lstStyle → its layout and master placeholders' lstStyles
//   → the shape style's fontRef or the table style's cell text style → the master's title/body/other text style
//   → the presentation's defaultTextStyle.
import { COLOR_ELEMENTS } from './color.ts';
import { attr, child, num, oneOf } from './xml.ts';

const COLOR_CHILDREN: readonly string[] = COLOR_ELEMENTS;

export const FILL_ELEMENTS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'] as const;

/** Text defaults from a shape style (p:style/a:fontRef) or a table style (a:tcTxStyle). */
export interface TextDefaults {
  /** An element whose first color child is the text color. */
  readonly color?: Element;
  readonly typeface?: 'major' | 'minor';
  readonly bold?: boolean;
  readonly italic?: boolean;
}

/** List styles from least to most specific; `base` sits below the defaults, `shape` above them. */
export interface TextLayers {
  readonly base: readonly (Element | undefined)[];
  readonly shape: readonly (Element | undefined)[];
  readonly defaults?: TextDefaults;
}

/** p:style/a:fontRef → text defaults. */
export function fontRefDefaults(fontRef: Element | undefined): TextDefaults | undefined {
  if (!fontRef) {
    return undefined;
  }
  const idx = attr(fontRef, 'idx');
  return {
    color: oneOf(fontRef, COLOR_CHILDREN) ? fontRef : undefined,
    typeface: idx === 'major' || idx === 'minor' ? idx : undefined,
  };
}

/** a:tcTxStyle of a table style region → text defaults. */
export function cellTextDefaults(style: Element | undefined): TextDefaults | undefined {
  if (!style) {
    return undefined;
  }
  const fontRef = child(style, 'fontRef');
  const flag = (name: string) => (attr(style, name) === 'on' ? true : attr(style, name) === 'off' ? false : undefined);
  const idx = attr(fontRef, 'idx');
  return {
    color: oneOf(style, COLOR_CHILDREN) ? style : fontRef && oneOf(fontRef, COLOR_CHILDREN) ? fontRef : undefined,
    typeface: idx === 'major' || idx === 'minor' ? idx : undefined,
    bold: flag('b'),
    italic: flag('i'),
  };
}

export interface Spacing {
  /** Multiple of the line (1 = single spacing) or of the font size for space before/after. */
  readonly pct?: number;
  /** Points. */
  readonly pts?: number;
}

export interface BulletStyle {
  readonly kind: 'none' | 'char' | 'autonum' | 'blip';
  readonly char?: string;
  readonly scheme?: string;
  readonly startAt?: number;
  /** Typeface element (a:buFont); undefined means the text's font. */
  readonly font?: Element;
  /** Color element holder (a:buClr); undefined means the text's color. */
  readonly color?: Element;
  readonly sizePct?: number;
  readonly sizePts?: number;
}

export interface ParagraphStyle {
  readonly level: number;
  readonly align: string;
  /** EMUs. */
  readonly marL: number;
  readonly indent: number;
  readonly rtl: boolean;
  readonly lineSpacing?: Spacing;
  readonly spaceBefore?: Spacing;
  readonly spaceAfter?: Spacing;
  readonly bullet: BulletStyle;
  /** EMUs. */
  readonly defaultTab: number;
}

export interface RunStyle {
  /** Points. */
  readonly size: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: string;
  readonly strike: string;
  /** Percent of the font size (30 = raised by 30%). */
  readonly baseline: number;
  readonly cap: string;
  /** Points (letter spacing). */
  readonly spacing: number;
  /** Fill element for the text (solidFill, gradFill, noFill, ...), or an element holding the default color. */
  readonly fill?: Element;
  readonly highlight?: Element;
  readonly latin?: string;
  readonly ea?: string;
  readonly cs?: string;
  readonly sym?: string;
  readonly lang?: string;
}

const DEFAULT_FONT_SIZE = 18;

export class TextStyleResolver {
  private readonly layers: TextLayers;

  constructor(layers: TextLayers) {
    this.layers = layers;
  }

  /** pPr-like elements for a level, most specific first. */
  private paragraphSources(level: number, pPr: Element | undefined): Element[] {
    const sources: Element[] = [];
    if (pPr) {
      sources.push(pPr);
    }
    const levelName = `lvl${level + 1}pPr`;
    for (const style of [...this.layers.shape].reverse()) {
      pushLevel(sources, style, levelName);
    }
    for (const style of [...this.layers.base].reverse()) {
      pushLevel(sources, style, levelName);
    }
    return sources;
  }

  paragraph(pPr: Element | undefined): ParagraphStyle {
    const level = Math.max(0, Math.min(8, Math.trunc(num(pPr, 'lvl') ?? 0)));
    const sources = this.paragraphSources(level, pPr);
    const value = (name: string) => firstAttr(sources, name);
    const group = (names: readonly string[]) => {
      for (const source of sources) {
        const found = oneOf(source, names);
        if (found) {
          return found;
        }
      }
      return undefined;
    };
    return {
      level,
      align: value('algn') ?? 'l',
      marL: Number(value('marL') ?? 0) || 0,
      indent: Number(value('indent') ?? 0) || 0,
      rtl: value('rtl') === '1',
      lineSpacing: spacing(group(['lnSpc'])),
      spaceBefore: spacing(group(['spcBef'])),
      spaceAfter: spacing(group(['spcAft'])),
      bullet: bullet(group(['buNone', 'buAutoNum', 'buChar', 'buBlip']), group(['buFontTx', 'buFont']), group(['buClrTx', 'buClr']), group(['buSzTx', 'buSzPct', 'buSzPts'])),
      defaultTab: Number(value('defTabSz') ?? 914400) || 914400,
    };
  }

  /** Run properties; `rPr` is the run's own a:rPr (or a:endParaRPr for an empty paragraph). */
  run(level: number, pPr: Element | undefined, rPr: Element | undefined): RunStyle {
    const levelName = `lvl${level + 1}pPr`;
    const shapeSources: Element[] = [];
    if (rPr) {
      shapeSources.push(rPr);
    }
    const ownDefaults = child(pPr, 'defRPr');
    if (ownDefaults) {
      shapeSources.push(ownDefaults);
    }
    for (const style of [...this.layers.shape].reverse()) {
      pushDefaults(shapeSources, style, levelName);
    }
    const baseSources: Element[] = [];
    for (const style of [...this.layers.base].reverse()) {
      pushDefaults(baseSources, style, levelName);
    }
    const all = [...shapeSources, ...baseSources];
    const value = (name: string) => firstAttr(all, name);
    const element = (sources: readonly Element[], names: readonly string[]) => {
      for (const source of sources) {
        const found = oneOf(source, names);
        if (found) {
          return found;
        }
      }
      return undefined;
    };
    const defaults = this.layers.defaults;
    const typeface = (name: 'latin' | 'ea' | 'cs' | 'sym') => {
      const own = attr(element(shapeSources, [name]), 'typeface');
      if (own) {
        return own;
      }
      if (defaults?.typeface && name !== 'sym') {
        return `${defaults.typeface === 'major' ? '+mj' : '+mn'}-${name === 'latin' ? 'lt' : name}`;
      }
      return attr(element(baseSources, [name]), 'typeface');
    };
    const flag = (name: 'b' | 'i', fallback: boolean | undefined) => {
      const raw = firstAttr(shapeSources, name) ?? (fallback === undefined ? firstAttr(baseSources, name) : fallback ? '1' : '0');
      return raw === '1' || raw === 'true';
    };
    return {
      size: (Number(value('sz')) || DEFAULT_FONT_SIZE * 100) / 100,
      bold: flag('b', defaults?.bold),
      italic: flag('i', defaults?.italic),
      underline: value('u') ?? 'none',
      strike: value('strike') ?? 'noStrike',
      baseline: (Number(value('baseline')) || 0) / 1000,
      cap: value('cap') ?? 'none',
      spacing: (Number(value('spc')) || 0) / 100,
      fill: element(shapeSources, FILL_ELEMENTS) ?? defaults?.color ?? element(baseSources, FILL_ELEMENTS),
      highlight: element(all, ['highlight']),
      latin: typeface('latin'),
      ea: typeface('ea'),
      cs: typeface('cs'),
      sym: typeface('sym'),
      lang: value('lang'),
    };
  }
}

function pushLevel(sources: Element[], style: Element | undefined, levelName: string): void {
  const level = child(style, levelName);
  if (level) {
    sources.push(level);
  }
  const defaults = child(style, 'defPPr');
  if (defaults) {
    sources.push(defaults);
  }
}

function pushDefaults(sources: Element[], style: Element | undefined, levelName: string): void {
  for (const name of [levelName, 'defPPr']) {
    const defaults = child(child(style, name), 'defRPr');
    if (defaults) {
      sources.push(defaults);
    }
  }
}

function firstAttr(sources: readonly Element[], name: string): string | undefined {
  for (const source of sources) {
    const value = source.getAttribute(name);
    if (value !== null) {
      return value;
    }
  }
  return undefined;
}

function spacing(element: Element | undefined): Spacing | undefined {
  const pct = child(element, 'spcPct');
  if (pct) {
    return { pct: (num(pct, 'val') ?? 100000) / 100000 };
  }
  const pts = child(element, 'spcPts');
  if (pts) {
    return { pts: (num(pts, 'val') ?? 0) / 100 };
  }
  return undefined;
}

function bullet(type: Element | undefined, font: Element | undefined, color: Element | undefined, size: Element | undefined): BulletStyle {
  const common = {
    font: font?.localName === 'buFont' ? font : undefined,
    color: color?.localName === 'buClr' ? color : undefined,
    sizePct: size?.localName === 'buSzPct' ? (num(size, 'val') ?? 100000) / 100000 : undefined,
    sizePts: size?.localName === 'buSzPts' ? (num(size, 'val') ?? 0) / 100 : undefined,
  };
  switch (type?.localName) {
    case 'buChar':
      return { kind: 'char', char: attr(type, 'char') ?? '•', ...common };
    case 'buAutoNum':
      return { kind: 'autonum', scheme: attr(type, 'type') ?? 'arabicPeriod', startAt: num(type, 'startAt') ?? 1, ...common };
    case 'buBlip':
      return { kind: 'blip', ...common };
    default:
      return { kind: 'none' };
  }
}
