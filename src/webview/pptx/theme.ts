// The theme part: color scheme, font scheme and the format scheme's style lists that shapes refer to by index
// (fillRef, lnRef, effectRef, bgRef).
import { attr, child, find, kids } from './xml.ts';

export interface FontScheme {
  readonly latin: string;
  readonly ea: string;
  readonly cs: string;
  /** Per-script fonts (Jpan, Hans, Hant, Hang, ...) used when `ea` is empty. */
  readonly scripts: Readonly<Record<string, string>>;
}

export interface Theme {
  readonly colors: Readonly<Record<string, string>>;
  readonly major: FontScheme;
  readonly minor: FontScheme;
  readonly fillStyles: readonly Element[];
  readonly lineStyles: readonly Element[];
  readonly effectStyles: readonly Element[];
  readonly bgFillStyles: readonly Element[];
}

/** The Office theme's colors, for packages whose theme is missing or incomplete. */
const DEFAULT_COLORS: Record<string, string> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

export const DEFAULT_THEME: Theme = {
  colors: DEFAULT_COLORS,
  major: { latin: 'Calibri Light', ea: '', cs: '', scripts: {} },
  minor: { latin: 'Calibri', ea: '', cs: '', scripts: {} },
  fillStyles: [],
  lineStyles: [],
  effectStyles: [],
  bgFillStyles: [],
};

export function parseTheme(root: Element): Theme {
  const elements = child(root, 'themeElements');
  const colors: Record<string, string> = { ...DEFAULT_COLORS };
  for (const entry of kids(child(elements, 'clrScheme'))) {
    const color = kids(entry)[0];
    const value = color?.localName === 'sysClr' ? (attr(color, 'lastClr') ?? (attr(color, 'val') === 'window' ? 'FFFFFF' : '000000')) : attr(color, 'val');
    if (value && /^[0-9a-fA-F]{6}$/.test(value)) {
      colors[entry.localName] = value;
    }
  }
  const fonts = child(elements, 'fontScheme');
  const format = child(elements, 'fmtScheme');
  return {
    colors,
    major: fontScheme(child(fonts, 'majorFont'), DEFAULT_THEME.major),
    minor: fontScheme(child(fonts, 'minorFont'), DEFAULT_THEME.minor),
    fillStyles: kids(child(format, 'fillStyleLst')),
    lineStyles: kids(child(format, 'lnStyleLst')),
    effectStyles: kids(child(format, 'effectStyleLst')),
    bgFillStyles: kids(child(format, 'bgFillStyleLst')),
  };
}

function fontScheme(element: Element | undefined, fallback: FontScheme): FontScheme {
  if (!element) {
    return fallback;
  }
  const scripts: Record<string, string> = {};
  for (const font of kids(element)) {
    const script = attr(font, 'script');
    const typeface = attr(font, 'typeface');
    if (font.localName === 'font' && script && typeface) {
      scripts[script] = typeface;
    }
  }
  return {
    latin: attr(find(element, 'latin'), 'typeface') ?? fallback.latin,
    ea: attr(find(element, 'ea'), 'typeface') ?? '',
    cs: attr(find(element, 'cs'), 'typeface') ?? '',
    scripts,
  };
}
