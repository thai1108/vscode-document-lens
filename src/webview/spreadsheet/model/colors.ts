// Spreadsheet colors (RGB, legacy palette indexes, theme colors with tint) resolved to `#rrggbb`.

/** Excel's default 56-color palette, as palette indexes 8-63 (0-7 repeat the first eight). */
const DEFAULT_PALETTE = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

/** Palette indexes 0-63: the eight basic colors, then the 56 customizable ones. */
export function defaultPalette(): string[] {
  return [...DEFAULT_PALETTE.slice(0, 8), ...DEFAULT_PALETTE].map((hex) => `#${hex.toLowerCase()}`);
}

/** Office theme colors in the order cells refer to them (`theme="0"` is lt1, `theme="1"` dk1, ...). */
export const DEFAULT_THEME = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#0563c1', '#954f72'];

export interface ColorContext {
  palette: readonly string[];
  theme: readonly string[];
}

export interface ColorSpec {
  rgb?: string;
  theme?: number;
  indexed?: number;
  tint?: number;
  auto?: boolean;
}

/** `#rrggbb` for a color spec, or undefined for automatic/system colors and anything unreadable. */
export function resolveColor(spec: ColorSpec, context: ColorContext): string | undefined {
  let color: string | undefined;
  if (spec.rgb !== undefined) {
    color = parseHexColor(spec.rgb);
  } else if (spec.theme !== undefined) {
    color = context.theme[spec.theme];
  } else if (spec.indexed !== undefined) {
    color = context.palette[spec.indexed];
  }
  if (!color) {
    return undefined;
  }
  return spec.tint ? applyTint(color, spec.tint) : color;
}

/** `RRGGBB` or `AARRGGBB` (the alpha byte is ignored, as Excel does) to `#rrggbb`. */
export function parseHexColor(text: string): string | undefined {
  const hex = text.trim();
  if (!/^(?:[0-9a-fA-F]{2})?[0-9a-fA-F]{6}$/.test(hex)) {
    return undefined;
  }
  return `#${hex.slice(-6).toLowerCase()}`;
}

/** Lightens (tint > 0) or darkens (tint < 0) a color in HSL space, the way Office applies theme tints. */
export function applyTint(color: string, tint: number): string {
  const amount = Math.max(-1, Math.min(1, tint));
  const [h, s, l] = rgbToHsl(color);
  const lightness = amount < 0 ? l * (1 + amount) : l * (1 - amount) + amount;
  return hslToRgb(h, s, lightness);
}

/** Mixes `color` over `background` at `weight` (0-1) of `color`. */
export function mixColors(color: string, background: string, weight: number): string {
  const a = channels(color);
  const b = channels(background);
  return toHex(a.map((value, i) => value * weight + (b[i] ?? 0) * (1 - weight)));
}

function channels(color: string): number[] {
  const value = parseInt(color.slice(1), 16);
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

function toHex(values: number[]): string {
  return `#${values.map((value) => Math.round(Math.max(0, Math.min(255, value))).toString(16).padStart(2, '0')).join('')}`;
}

function rgbToHsl(color: string): [number, number, number] {
  const [r = 0, g = 0, b = 0] = channels(color).map((value) => value / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) {
    return [0, 0, l];
  }
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): string {
  if (s === 0) {
    return toHex([l * 255, l * 255, l * 255]);
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number) => {
    const x = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    if (x < 1 / 6) {
      return p + (q - p) * 6 * x;
    }
    if (x < 1 / 2) {
      return q;
    }
    if (x < 2 / 3) {
      return p + (q - p) * (2 / 3 - x) * 6;
    }
    return p;
  };
  return toHex([hue(h + 1 / 3) * 255, hue(h) * 255, hue(h - 1 / 3) * 255]);
}
