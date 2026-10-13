// DrawingML colors: the six color kinds, the master's color map (bg1 → lt1, ...) and the color transforms. Transform
// semantics follow what PowerPoint renders: lum/sat/hue work in HSL, tint and shade in linear RGB.
import { attr, kids, num } from './xml.ts';

/** r, g, b in 0..255 (unrounded), a in 0..1. */
export interface Rgba {
  r: number;
  g: number;
  b: number;
  a: number;
}

export interface ColorContext {
  /** Theme color scheme: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink → RRGGBB. */
  readonly scheme: Readonly<Record<string, string>>;
  /** Master color map, possibly overridden by the slide: bg1, tx1, bg2, tx2, accent1.. → scheme names. */
  readonly map: Readonly<Record<string, string>>;
  /** phClr: the color a theme style reference (fillRef, lnRef, ...) puts into the theme's style. */
  readonly placeholder?: Rgba;
}

export const COLOR_ELEMENTS = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'] as const;

export const DEFAULT_COLOR_MAP: Readonly<Record<string, string>> = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

const SYSTEM_COLORS: Record<string, string> = {
  windowText: '000000',
  window: 'FFFFFF',
  btnFace: 'F0F0F0',
  btnText: '000000',
  highlight: '0078D7',
  highlightText: 'FFFFFF',
  grayText: '6D6D6D',
  menu: 'F0F0F0',
  menuText: '000000',
  infoBk: 'FFFFE1',
  infoText: '000000',
  btnShadow: 'A0A0A0',
  btnHighlight: 'FFFFFF',
  '3dDkShadow': '696969',
  '3dLight': 'E3E3E3',
  windowFrame: '646464',
  captionText: '000000',
  scrollBar: 'C8C8C8',
  background: '000000',
  activeCaption: '99B4D1',
  inactiveCaption: 'BFCDDB',
  appWorkspace: 'ABABAB',
  activeBorder: 'B4B4B4',
  inactiveBorder: 'F4F7FC',
  inactiveCaptionText: '000000',
  hotLight: '0066CC',
  gradientActiveCaption: 'B9D1EA',
  gradientInactiveCaption: 'D7E4F2',
  menuBar: 'F0F0F0',
  menuHighlight: '3399FF',
};

// The DrawingML preset colors are the CSS named colors, with dk/lt/med abbreviations for dark/light/medium.
const NAMED_COLORS =
  'aliceblue:f0f8ff,antiquewhite:faebd7,aqua:00ffff,aquamarine:7fffd4,azure:f0ffff,beige:f5f5dc,bisque:ffe4c4,black:000000,' +
  'blanchedalmond:ffebcd,blue:0000ff,blueviolet:8a2be2,brown:a52a2a,burlywood:deb887,cadetblue:5f9ea0,chartreuse:7fff00,' +
  'chocolate:d2691e,coral:ff7f50,cornflowerblue:6495ed,cornsilk:fff8dc,crimson:dc143c,cyan:00ffff,darkblue:00008b,' +
  'darkcyan:008b8b,darkgoldenrod:b8860b,darkgray:a9a9a9,darkgrey:a9a9a9,darkgreen:006400,darkkhaki:bdb76b,darkmagenta:8b008b,' +
  'darkolivegreen:556b2f,darkorange:ff8c00,darkorchid:9932cc,darkred:8b0000,darksalmon:e9967a,darkseagreen:8fbc8f,' +
  'darkslateblue:483d8b,darkslategray:2f4f4f,darkslategrey:2f4f4f,darkturquoise:00ced1,darkviolet:9400d3,deeppink:ff1493,' +
  'deepskyblue:00bfff,dimgray:696969,dimgrey:696969,dodgerblue:1e90ff,firebrick:b22222,floralwhite:fffaf0,forestgreen:228b22,' +
  'fuchsia:ff00ff,gainsboro:dcdcdc,ghostwhite:f8f8ff,gold:ffd700,goldenrod:daa520,gray:808080,grey:808080,green:008000,' +
  'greenyellow:adff2f,honeydew:f0fff0,hotpink:ff69b4,indianred:cd5c5c,indigo:4b0082,ivory:fffff0,khaki:f0e68c,lavender:e6e6fa,' +
  'lavenderblush:fff0f5,lawngreen:7cfc00,lemonchiffon:fffacd,lightblue:add8e6,lightcoral:f08080,lightcyan:e0ffff,' +
  'lightgoldenrodyellow:fafad2,lightgray:d3d3d3,lightgrey:d3d3d3,lightgreen:90ee90,lightpink:ffb6c1,lightsalmon:ffa07a,' +
  'lightseagreen:20b2aa,lightskyblue:87cefa,lightslategray:778899,lightslategrey:778899,lightsteelblue:b0c4de,' +
  'lightyellow:ffffe0,lime:00ff00,limegreen:32cd32,linen:faf0e6,magenta:ff00ff,maroon:800000,mediumaquamarine:66cdaa,' +
  'mediumblue:0000cd,mediumorchid:ba55d3,mediumpurple:9370db,mediumseagreen:3cb371,mediumslateblue:7b68ee,' +
  'mediumspringgreen:00fa9a,mediumturquoise:48d1cc,mediumvioletred:c71585,midnightblue:191970,mintcream:f5fffa,' +
  'mistyrose:ffe4e1,moccasin:ffe4b5,navajowhite:ffdead,navy:000080,oldlace:fdf5e6,olive:808000,olivedrab:6b8e23,' +
  'orange:ffa500,orangered:ff4500,orchid:da70d6,palegoldenrod:eee8aa,palegreen:98fb98,paleturquoise:afeeee,' +
  'palevioletred:db7093,papayawhip:ffefd5,peachpuff:ffdab9,peru:cd853f,pink:ffc0cb,plum:dda0dd,powderblue:b0e0e6,' +
  'purple:800080,red:ff0000,rosybrown:bc8f8f,royalblue:4169e1,saddlebrown:8b4513,salmon:fa8072,sandybrown:f4a460,' +
  'seagreen:2e8b57,seashell:fff5ee,sienna:a0522d,silver:c0c0c0,skyblue:87ceeb,slateblue:6a5acd,slategray:708090,' +
  'slategrey:708090,snow:fffafa,springgreen:00ff7f,steelblue:4682b4,tan:d2b48c,teal:008080,thistle:d8bfd8,tomato:ff6347,' +
  'turquoise:40e0d0,violet:ee82ee,wheat:f5deb3,white:ffffff,whitesmoke:f5f5f5,yellow:ffff00,yellowgreen:9acd32';

let namedColors: Map<string, string> | undefined;

export function presetColor(name: string): string | undefined {
  namedColors ??= new Map(NAMED_COLORS.split(',').map((pair) => pair.split(':') as [string, string]));
  const css = name
    .replace(/^dk(?=[A-Z])/, 'dark')
    .replace(/^lt(?=[A-Z])/, 'light')
    .replace(/^med(?=[A-Z])/, 'medium')
    .toLowerCase();
  return namedColors.get(css);
}

export function hexToRgba(hex: string | undefined): Rgba | undefined {
  if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex)) {
    return undefined;
  }
  const value = Number.parseInt(hex, 16);
  return { r: (value >> 16) & 0xff, g: (value >> 8) & 0xff, b: value & 0xff, a: 1 };
}

/** The color of a color element (srgbClr, schemeClr, ...) with its transforms applied. */
export function readColor(element: Element | undefined, context: ColorContext): Rgba | undefined {
  if (!element) {
    return undefined;
  }
  const base = baseColor(element, context);
  return base && applyTransforms(base, kids(element));
}

/** The first color element directly inside `parent` (e.g. a solidFill), resolved. */
export function colorIn(parent: Element | undefined, context: ColorContext): Rgba | undefined {
  const element = kids(parent).find((node) => (COLOR_ELEMENTS as readonly string[]).includes(node.localName));
  return readColor(element, context);
}

function baseColor(element: Element, context: ColorContext): Rgba | undefined {
  const value = attr(element, 'val') ?? '';
  switch (element.localName) {
    case 'srgbClr':
      return hexToRgba(value);
    case 'schemeClr':
      if (value === 'phClr') {
        return context.placeholder && { ...context.placeholder };
      }
      return hexToRgba(context.scheme[context.map[value] ?? value]);
    case 'sysClr':
      return hexToRgba(attr(element, 'lastClr')) ?? hexToRgba(SYSTEM_COLORS[value]);
    case 'prstClr':
      return hexToRgba(presetColor(value));
    case 'scrgbClr':
      return fromLinear({ r: (num(element, 'r') ?? 0) / 100000, g: (num(element, 'g') ?? 0) / 100000, b: (num(element, 'b') ?? 0) / 100000 });
    case 'hslClr':
      return hslToRgba({ h: (num(element, 'hue') ?? 0) / 60000, s: (num(element, 'sat') ?? 0) / 100000, l: (num(element, 'lum') ?? 0) / 100000 }, 1);
    default:
      return undefined;
  }
}

interface Hsl {
  h: number;
  s: number;
  l: number;
}

interface Linear {
  r: number;
  g: number;
  b: number;
}

const clamp01 = (value: number) => Math.min(1, Math.max(0, value));

/** Applies the transform elements (lumMod, tint, alpha, ...) in document order. */
function applyTransforms(color: Rgba, transforms: readonly Element[]): Rgba {
  return transforms.reduce((result, transform) => transformColor(result, transform.localName, num(transform, 'val')), color);
}

/** One color transform; `raw` is the val attribute (1/100000 for percentages, 60000ths of a degree for hues). */
export function transformColor(color: Rgba, name: string, raw: number | undefined): Rgba {
  let result = { ...color };
  const value = (raw ?? 0) / 100000;
  const hsl = () => rgbaToHsl(result);
  const setHsl = (next: Hsl) => {
    result = hslToRgba(next, result.a);
  };
  const linear = () => toLinear(result);
  const setLinear = (next: Linear) => {
    result = { ...fromLinear(next), a: result.a };
  };
  switch (name) {
    case 'alpha':
      result.a = clamp01(value);
      break;
    case 'alphaMod':
      result.a = clamp01(result.a * value);
      break;
    case 'alphaOff':
      result.a = clamp01(result.a + value);
      break;
    case 'lum':
      setHsl({ ...hsl(), l: clamp01(value) });
      break;
    case 'lumMod': {
      const current = hsl();
      setHsl({ ...current, l: clamp01(current.l * value) });
      break;
    }
    case 'lumOff': {
      const current = hsl();
      setHsl({ ...current, l: clamp01(current.l + value) });
      break;
    }
    case 'sat':
      setHsl({ ...hsl(), s: clamp01(value) });
      break;
    case 'satMod': {
      const current = hsl();
      setHsl({ ...current, s: clamp01(current.s * value) });
      break;
    }
    case 'satOff': {
      const current = hsl();
      setHsl({ ...current, s: clamp01(current.s + value) });
      break;
    }
    case 'hue':
      setHsl({ ...hsl(), h: wrapHue((raw ?? 0) / 60000) });
      break;
    case 'hueMod': {
      const current = hsl();
      setHsl({ ...current, h: wrapHue(current.h * value) });
      break;
    }
    case 'hueOff': {
      const current = hsl();
      setHsl({ ...current, h: wrapHue(current.h + (raw ?? 0) / 60000) });
      break;
    }
    case 'comp': {
      const current = hsl();
      setHsl({ ...current, h: wrapHue(current.h + 180) });
      break;
    }
    case 'tint':
      if (raw !== undefined && value < 1) {
        const c = linear();
        const tint = clamp01(value);
        setLinear({ r: 1 - (1 - c.r) * tint, g: 1 - (1 - c.g) * tint, b: 1 - (1 - c.b) * tint });
      }
      break;
    case 'shade':
      if (raw !== undefined && value < 1) {
        const c = linear();
        const shade = clamp01(value);
        setLinear({ r: c.r * shade, g: c.g * shade, b: c.b * shade });
      }
      break;
    case 'inv':
      result = { r: 255 - result.r, g: 255 - result.g, b: 255 - result.b, a: result.a };
      break;
    case 'gray': {
      const gray = 0.3 * result.r + 0.59 * result.g + 0.11 * result.b;
      result = { r: gray, g: gray, b: gray, a: result.a };
      break;
    }
    case 'red':
    case 'green':
    case 'blue':
    case 'redMod':
    case 'greenMod':
    case 'blueMod':
    case 'redOff':
    case 'greenOff':
    case 'blueOff': {
      const c = linear();
      const key = name.startsWith('red') ? 'r' : name.startsWith('green') ? 'g' : 'b';
      if (name.endsWith('Mod')) {
        c[key] = clamp01(c[key] * value);
      } else if (name.endsWith('Off')) {
        c[key] = clamp01(c[key] + value);
      } else {
        c[key] = clamp01(value);
      }
      setLinear(c);
      break;
    }
    default:
      break;
  }
  return result;
}

function wrapHue(hue: number): number {
  const wrapped = hue % 360;
  return wrapped < 0 ? wrapped + 360 : wrapped;
}

// PowerPoint's tint/shade work on gamma-decoded components; 2.3 is the exponent Office uses for this conversion.
const GAMMA = 2.3;

function toLinear(color: Rgba): Linear {
  return { r: (color.r / 255) ** GAMMA, g: (color.g / 255) ** GAMMA, b: (color.b / 255) ** GAMMA };
}

function fromLinear(color: Linear): Rgba {
  const channel = (value: number) => clamp01(value) ** (1 / GAMMA) * 255;
  return { r: channel(color.r), g: channel(color.g), b: channel(color.b), a: 1 };
}

function rgbaToHsl(color: Rgba): Hsl {
  const r = color.r / 255;
  const g = color.g / 255;
  const b = color.b / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) {
    return { h: 0, s: 0, l };
  }
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) {
    h = (g - b) / d + (g < b ? 6 : 0);
  } else if (max === g) {
    h = (b - r) / d + 2;
  } else {
    h = (r - g) / d + 4;
  }
  return { h: h * 60, s, l };
}

function hslToRgba(hsl: Hsl, a: number): Rgba {
  const { s, l } = hsl;
  const h = wrapHue(hsl.h) / 360;
  if (s === 0) {
    return { r: l * 255, g: l * 255, b: l * 255, a };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channel = (t: number) => {
    const u = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    if (u < 1 / 6) {
      return p + (q - p) * 6 * u;
    }
    if (u < 1 / 2) {
      return q;
    }
    if (u < 2 / 3) {
      return p + (q - p) * (2 / 3 - u) * 6;
    }
    return p;
  };
  return { r: channel(h + 1 / 3) * 255, g: channel(h) * 255, b: channel(h - 1 / 3) * 255, a };
}

/** #rrggbb, ignoring alpha (SVG takes it separately as fill-opacity / stroke-opacity). */
export function toHex(color: Rgba): string {
  const channel = (value: number) =>
    Math.round(Math.min(255, Math.max(0, value)))
      .toString(16)
      .padStart(2, '0');
  return `#${channel(color.r)}${channel(color.g)}${channel(color.b)}`;
}

/** A CSS color: #rrggbb when opaque, rgba() otherwise. */
export function toCss(color: Rgba): string {
  if (color.a >= 1) {
    return toHex(color);
  }
  const channel = (value: number) => Math.round(Math.min(255, Math.max(0, value)));
  return `rgba(${channel(color.r)}, ${channel(color.g)}, ${channel(color.b)}, ${Math.round(clamp01(color.a) * 1000) / 1000})`;
}
