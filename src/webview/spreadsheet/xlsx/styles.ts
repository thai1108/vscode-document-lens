// xl/styles.xml (cell formats: fonts, fills, borders, alignment, number formats) and the theme colors they refer to.
import { SHEET_LIMITS } from '../limits.ts';
import { DEFAULT_THEME, defaultPalette, mixColors, parseHexColor, resolveColor, type ColorContext, type ColorSpec } from '../model/colors.ts';
import { builtinFormat } from '../model/numberFormat.ts';
import { DEFAULT_FONT, defaultFormat, type BorderEdge, type BorderStyle, type Borders, type CellFormat, type Font, type HorizontalAlignment, type VerticalAlignment } from '../model/types.ts';
import { scanXml, type XmlTag } from './xml.ts';

export interface Styles {
  formats: CellFormat[];
  /** The font of the Normal style, which sets the default column width and row height. */
  defaultFont: Font;
  colors: ColorContext;
}

/** Theme colors from xl/theme/theme1.xml, in the order cells refer to them. */
export function parseTheme(xml: string): string[] {
  const scheme: Record<string, string> = {};
  let current: string | undefined;
  let depth = 0;
  scanXml(xml, {
    open(tag) {
      if (tag.name === 'clrScheme') {
        depth = 1;
      } else if (depth === 1 && !tag.selfClosing) {
        current = tag.name;
        depth = 2;
      } else if (depth === 2 && current && !scheme[current]) {
        const value = tag.name === 'srgbClr' ? tag.attr('val') : tag.name === 'sysClr' ? tag.attr('lastClr') : undefined;
        const color = value ? parseHexColor(value) : undefined;
        if (color) {
          scheme[current] = color;
        }
      }
    },
    close(name) {
      if (name === 'clrScheme') {
        depth = 0;
      } else if (depth === 2 && name === current) {
        depth = 1;
        current = undefined;
      }
    },
  });
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return order.map((name, i) => scheme[name] ?? DEFAULT_THEME[i] ?? '#000000');
}

interface FillSpec {
  pattern: string;
  fg?: ColorSpec;
  bg?: ColorSpec;
}

/** How much of the foreground color each pattern shows, to approximate patterns with one color. */
const PATTERN_DENSITY: Record<string, number> = {
  solid: 1,
  darkGray: 0.75,
  mediumGray: 0.5,
  lightGray: 0.25,
  gray125: 0.125,
  gray0625: 0.0625,
  darkHorizontal: 0.5,
  darkVertical: 0.5,
  darkDown: 0.5,
  darkUp: 0.5,
  darkGrid: 0.5,
  darkTrellis: 0.75,
  lightHorizontal: 0.25,
  lightVertical: 0.25,
  lightDown: 0.25,
  lightUp: 0.25,
  lightGrid: 0.25,
  lightTrellis: 0.25,
};

const BORDER_STYLES = new Set<string>(['thin', 'medium', 'thick', 'dashed', 'mediumDashed', 'dotted', 'hair', 'double', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']);
const HORIZONTAL = new Set<string>(['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']);
const VERTICAL = new Set<string>(['top', 'center', 'bottom', 'justify', 'distributed']);

interface XfSpec {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  horizontal: HorizontalAlignment;
  vertical: VerticalAlignment;
  wrap: boolean;
  indent: number;
}

type BorderSide = 'top' | 'right' | 'bottom' | 'left';
interface BorderSpec {
  style: BorderStyle;
  color?: ColorSpec;
}

export function parseStyles(xml: string, theme: readonly string[] = DEFAULT_THEME): Styles {
  const numberFormats = new Map<number, string>();
  const fonts: Partial<Font & { colorSpec: ColorSpec }>[] = [];
  const fills: FillSpec[] = [];
  const borders: Partial<Record<BorderSide, BorderSpec>>[] = [];
  const xfs: XfSpec[] = [];
  const palette = defaultPalette();
  let customPalette: string[] | undefined;

  // Where the scanner is: the list being read and the item inside it.
  let list = '';
  let font: Partial<Font & { colorSpec: ColorSpec }> | undefined;
  let fill: FillSpec | undefined;
  let border: Partial<Record<BorderSide, BorderSpec>> | undefined;
  let side: BorderSide | undefined;
  let xf: XfSpec | undefined;

  scanXml(xml, {
    open(tag) {
      switch (tag.name) {
        case 'numFmts':
        case 'fonts':
        case 'fills':
        case 'borders':
        case 'cellXfs':
        case 'cellStyleXfs':
        case 'dxfs':
        case 'indexedColors':
          list = tag.name;
          return;
      }
      switch (list) {
        case 'numFmts':
          if (tag.name === 'numFmt') {
            const id = intAttr(tag, 'numFmtId');
            const code = tag.attr('formatCode');
            if (id !== undefined && code !== undefined) {
              numberFormats.set(id, code);
            }
          }
          return;
        case 'fonts':
          if (tag.name === 'font') {
            font = {};
          } else if (font) {
            readFontProperty(font, tag);
          }
          return;
        case 'fills':
          if (tag.name === 'fill') {
            fill = { pattern: 'none' };
          } else if (fill && tag.name === 'patternFill') {
            fill.pattern = tag.attr('patternType') ?? 'none';
          } else if (fill && tag.name === 'fgColor') {
            fill.fg = colorSpec(tag);
          } else if (fill && tag.name === 'bgColor') {
            fill.bg = colorSpec(tag);
          } else if (fill && tag.name === 'gradientFill') {
            fill.pattern = 'gradient';
          } else if (fill && tag.name === 'color' && fill.pattern === 'gradient' && !fill.fg) {
            fill.fg = colorSpec(tag);
          }
          return;
        case 'borders':
          if (tag.name === 'border') {
            border = {};
          } else if (border && isBorderSide(tag.name)) {
            const style = tag.attr('style');
            side = style && BORDER_STYLES.has(style) ? borderSide(tag.name) : undefined;
            if (side && style) {
              border[side] = { style: style as BorderStyle };
            }
          } else if (border && side && tag.name === 'color') {
            const edge = border[side];
            if (edge) {
              edge.color = colorSpec(tag);
            }
          }
          return;
        case 'cellXfs':
          if (tag.name === 'xf') {
            xf = {
              numFmtId: intAttr(tag, 'numFmtId') ?? 0,
              fontId: intAttr(tag, 'fontId') ?? 0,
              fillId: intAttr(tag, 'fillId') ?? 0,
              borderId: intAttr(tag, 'borderId') ?? 0,
              horizontal: 'general',
              vertical: 'bottom',
              wrap: false,
              indent: 0,
            };
            if (xfs.length < SHEET_LIMITS.maxStyles) {
              xfs.push(xf);
            }
          } else if (xf && tag.name === 'alignment') {
            const horizontal = tag.attr('horizontal');
            const vertical = tag.attr('vertical');
            if (horizontal && HORIZONTAL.has(horizontal)) {
              xf.horizontal = horizontal as HorizontalAlignment;
            }
            if (vertical && VERTICAL.has(vertical)) {
              xf.vertical = vertical as VerticalAlignment;
            }
            xf.wrap = boolAttr(tag, 'wrapText') ?? false;
            xf.indent = Math.min(250, Math.max(0, intAttr(tag, 'indent') ?? 0));
          }
          return;
        case 'indexedColors':
          if (tag.name === 'rgbColor') {
            const color = parseHexColor(tag.attr('rgb') ?? '');
            (customPalette ??= []).push(color ?? '#000000');
          }
          return;
      }
    },
    close(name) {
      switch (name) {
        case 'font':
          if (list === 'fonts' && font) {
            if (fonts.length < SHEET_LIMITS.maxStyles) {
              fonts.push(font);
            }
            font = undefined;
          }
          return;
        case 'fill':
          if (list === 'fills' && fill) {
            if (fills.length < SHEET_LIMITS.maxStyles) {
              fills.push(fill);
            }
            fill = undefined;
          }
          return;
        case 'border':
          if (list === 'borders' && border) {
            if (borders.length < SHEET_LIMITS.maxStyles) {
              borders.push(border);
            }
            border = undefined;
          }
          return;
        case 'left':
        case 'right':
        case 'top':
        case 'bottom':
        case 'start':
        case 'end':
          side = undefined;
          return;
        case 'xf':
          xf = undefined;
          return;
        case 'numFmts':
        case 'fonts':
        case 'fills':
        case 'borders':
        case 'cellXfs':
        case 'cellStyleXfs':
        case 'dxfs':
        case 'indexedColors':
          list = '';
          return;
      }
    },
  });

  const colors: ColorContext = { palette: customPalette ? [...customPalette, ...palette.slice(customPalette.length)] : palette, theme };
  const resolvedFonts = fonts.map((spec) => resolveFont(spec, colors));
  const defaultFont = resolvedFonts[0] ?? DEFAULT_FONT;
  const resolvedFills = fills.map((spec) => resolveFill(spec, colors));
  const resolvedBorders = borders.map((spec) => resolveBorders(spec, colors));
  const formats = xfs.map((spec): CellFormat => {
    const format: CellFormat = {
      font: resolvedFonts[spec.fontId] ?? defaultFont,
      borders: resolvedBorders[spec.borderId] ?? {},
      horizontal: spec.horizontal,
      vertical: spec.vertical,
      wrap: spec.wrap,
      indent: spec.indent,
      numberFormat: numberFormats.get(spec.numFmtId) ?? builtinFormat(spec.numFmtId),
    };
    const fill = resolvedFills[spec.fillId];
    if (fill) {
      format.fill = fill;
    }
    return format;
  });
  if (!formats.length) {
    formats.push(defaultFormat(defaultFont));
  }
  return { formats, defaultFont, colors };
}

export function colorSpec(tag: XmlTag): ColorSpec {
  const spec: ColorSpec = {};
  const rgb = tag.attr('rgb');
  const theme = intAttr(tag, 'theme');
  const indexed = intAttr(tag, 'indexed');
  const tint = Number(tag.attr('tint'));
  if (rgb !== undefined) {
    spec.rgb = rgb;
  } else if (theme !== undefined) {
    spec.theme = theme;
  } else if (indexed !== undefined) {
    spec.indexed = indexed;
  }
  if (Number.isFinite(tint) && tint !== 0) {
    spec.tint = tint;
  }
  if (boolAttr(tag, 'auto')) {
    spec.auto = true;
  }
  return spec;
}

/** Applies a font child element (`<b/>`, `<sz val="11"/>`, ...) of `<font>` or a rich text run's `<rPr>`. */
export function readFontProperty(font: Partial<Font & { colorSpec: ColorSpec }>, tag: XmlTag): void {
  const on = () => boolAttr(tag, 'val') ?? true;
  switch (tag.name) {
    case 'b':
      font.bold = on();
      return;
    case 'i':
      font.italic = on();
      return;
    case 'strike':
      font.strike = on();
      return;
    case 'u': {
      const value = tag.attr('val') ?? 'single';
      font.underline = value === 'none' ? 'none' : value.startsWith('double') ? 'double' : 'single';
      return;
    }
    case 'sz': {
      const size = Number(tag.attr('val'));
      if (Number.isFinite(size) && size > 0) {
        font.size = Math.min(409, Math.max(1, size));
      }
      return;
    }
    case 'name':
    case 'rFont': {
      const name = tag.attr('val');
      if (name) {
        font.name = name.slice(0, 64);
      }
      return;
    }
    case 'color':
      font.colorSpec = colorSpec(tag);
      return;
    case 'vertAlign': {
      const value = tag.attr('val');
      if (value === 'superscript' || value === 'subscript') {
        font.vertAlign = value;
      }
      return;
    }
  }
}

export function resolveFont(spec: Partial<Font & { colorSpec: ColorSpec }>, colors: ColorContext, base: Font = DEFAULT_FONT): Font {
  const { colorSpec: color, ...rest } = spec;
  const font: Font = { ...base, ...rest };
  const resolved = color && !color.auto ? resolveColor(color, colors) : undefined;
  if (resolved) {
    font.color = resolved;
  } else {
    delete font.color;
  }
  return font;
}

/** Only the properties a rich text run sets, with its color resolved. */
export function resolveRunFont(spec: Partial<Font & { colorSpec: ColorSpec }>, colors: ColorContext): Partial<Font> {
  const { colorSpec: color, ...font } = spec;
  const resolved = color && !color.auto ? resolveColor(color, colors) : undefined;
  return resolved ? { ...font, color: resolved } : font;
}

function resolveFill(spec: FillSpec, colors: ColorContext): string | undefined {
  if (spec.pattern === 'none') {
    return undefined;
  }
  if (spec.pattern === 'gradient') {
    return spec.fg ? resolveColor(spec.fg, colors) : undefined;
  }
  const density = PATTERN_DENSITY[spec.pattern] ?? 0.5;
  // A pattern's foreground defaults to black and its background to white (palette 64/65, the system colors).
  const fg = (spec.fg && resolveColor(spec.fg, colors)) ?? (density === 1 ? undefined : '#000000');
  if (density === 1) {
    return fg;
  }
  const bg = (spec.bg && resolveColor(spec.bg, colors)) ?? '#ffffff';
  return fg ? mixColors(fg, bg, density) : undefined;
}

function resolveBorders(spec: Partial<Record<BorderSide, BorderSpec>>, colors: ColorContext): Borders {
  const result: Borders = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const edge = spec[side];
    if (edge) {
      const resolved: BorderEdge = { style: edge.style, color: (edge.color && !edge.color.auto && resolveColor(edge.color, colors)) || '#000000' };
      result[side] = resolved;
    }
  }
  return result;
}

function isBorderSide(name: string): boolean {
  return name === 'left' || name === 'right' || name === 'top' || name === 'bottom' || name === 'start' || name === 'end';
}

function borderSide(name: string): BorderSide {
  return name === 'start' ? 'left' : name === 'end' ? 'right' : (name as BorderSide);
}

export function intAttr(tag: XmlTag, name: string): number | undefined {
  const text = tag.attr(name);
  if (text === undefined) {
    return undefined;
  }
  const value = Number(text);
  return Number.isInteger(value) && value >= 0 && value <= 0x7fffffff ? value : undefined;
}

export function boolAttr(tag: XmlTag, name: string): boolean | undefined {
  const text = tag.attr(name);
  if (text === undefined) {
    return undefined;
  }
  return text === '1' || text === 'true' || text === 'on';
}
