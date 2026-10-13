// Shape outlines: evaluates DrawingML guide formulas (preset shapes and custGeom alike) and builds SVG path data in
// CSS pixels, plus the text rectangle the geometry defines.
import { EMU_PER_PX } from './units.ts';
import { PPTX_LIMITS } from './limits.ts';
import { POLYGON_CORNERS, PRESET_ALIASES, PRESETS, STAR_POINTS, type PresetSource } from './presets.ts';
import { attr, child, children, kids, num } from './xml.ts';

export interface Formula {
  readonly name: string;
  readonly op: string;
  readonly args: readonly string[];
}

export type FillMode = 'norm' | 'none' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';

export interface PathCommand {
  readonly op: 'M' | 'L' | 'A' | 'Q' | 'C' | 'Z';
  readonly args: readonly string[];
}

export interface PathDef {
  /** Coordinate space of the path; coordinates are in shape EMUs when absent. */
  readonly w?: number;
  readonly h?: number;
  readonly fill: FillMode;
  readonly stroke: boolean;
  readonly commands: readonly PathCommand[];
}

export interface GeometryDef {
  readonly adjust: readonly Formula[];
  readonly guides: readonly Formula[];
  readonly paths: readonly PathDef[];
  readonly textRect?: readonly [string, string, string, string];
}

export interface BuiltPath {
  readonly d: string;
  readonly fill: FillMode;
  readonly stroke: boolean;
}

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface BuiltGeometry {
  readonly paths: readonly BuiltPath[];
  /** Text rectangle in CSS pixels relative to the shape's box. */
  readonly textRect: Rect;
}

const FILL_MODES: readonly FillMode[] = ['norm', 'none', 'lighten', 'lightenLess', 'darken', 'darkenLess'];
const ANGLE_UNIT = Math.PI / 180 / 60000;

const presetCache = new Map<string, GeometryDef>();

export function presetGeometry(name: string): GeometryDef | undefined {
  const key = PRESET_ALIASES[name] ?? name;
  let def = presetCache.get(key);
  if (!def) {
    const source = PRESETS[key];
    const star = STAR_POINTS[key];
    const corners = POLYGON_CORNERS[key];
    def = source ? fromSource(source) : star ? starGeometry(star[0], star[1]) : corners ? polygonGeometry(corners) : undefined;
    if (def) {
      presetCache.set(key, def);
    }
  }
  return def;
}

function fromSource(source: PresetSource): GeometryDef {
  const rect = source.rect?.split(' ');
  return {
    adjust: parseAdjust(source.av),
    guides: parseFormulas(source.gd),
    paths: (typeof source.path === 'string' ? [source.path] : source.path).map(parsePath),
    textRect: rect?.length === 4 ? (rect as [string, string, string, string]) : undefined,
  };
}

function parseAdjust(text: string | undefined): Formula[] {
  const tokens = text?.trim().split(/\s+/).filter(Boolean) ?? [];
  const result: Formula[] = [];
  for (let i = 0; i + 1 < tokens.length; i += 2) {
    result.push({ name: tokens[i] ?? '', op: 'val', args: [tokens[i + 1] ?? '0'] });
  }
  return result;
}

function parseFormulas(text: string | undefined): Formula[] {
  return (text ?? '')
    .split(';')
    .map((entry) => entry.trim().split(/\s+/))
    .filter((tokens) => tokens.length >= 2)
    .map(([name = '', op = 'val', ...args]) => ({ name, op, args }));
}

function parsePath(text: string): PathDef {
  const tokens = text.trim().split(/\s+/);
  let w: number | undefined;
  let h: number | undefined;
  let fill: FillMode = 'norm';
  let stroke = true;
  const commands: PathCommand[] = [];
  const arity: Record<string, number> = { M: 2, L: 2, A: 4, Q: 4, C: 6, Z: 0 };
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i] ?? '';
    const option = /^(fill|stroke|w|h)=(.+)$/.exec(token);
    if (option) {
      const [, key, value = ''] = option;
      if (key === 'fill') {
        fill = FILL_MODES.includes(value as FillMode) ? (value as FillMode) : 'norm';
      } else if (key === 'stroke') {
        stroke = value !== '0';
      } else if (key === 'w') {
        w = Number(value);
      } else {
        h = Number(value);
      }
      continue;
    }
    const count = arity[token];
    if (count === undefined) {
      continue;
    }
    commands.push({ op: token as PathCommand['op'], args: tokens.slice(i + 1, i + 1 + count) });
    i += count;
  }
  return { w, h, fill, stroke, commands };
}

/** Regular n-pointed star inscribed in the box, starting at the top. */
function starGeometry(points: number, adj: number): GeometryDef {
  const guides = ['a pin 0 adj 50000', 'iwd2 */ wd2 a 50000', 'ihd2 */ hd2 a 50000'];
  const commands: string[] = [];
  for (let k = 0; k < points * 2; k++) {
    const angle = Math.round(16200000 + (k * 10800000) / points);
    const [rx, ry] = k % 2 === 0 ? ['wd2', 'hd2'] : ['iwd2', 'ihd2'];
    guides.push(`dx${k} cos ${rx} ${angle}`, `dy${k} sin ${ry} ${angle}`, `x${k} +- hc dx${k} 0`, `y${k} +- vc dy${k} 0`);
    commands.push(`${k === 0 ? 'M' : 'L'} x${k} y${k}`);
  }
  guides.push('ix cos iwd2 2700000', 'iy sin ihd2 2700000', 'il +- hc 0 ix', 'ir +- hc ix 0', 'it +- vc 0 iy', 'ib +- vc iy 0');
  return fromSource({ av: `adj ${adj}`, gd: guides.join('; '), path: `${commands.join(' ')} Z`, rect: 'il it ir ib' });
}

function polygonGeometry(corners: number): GeometryDef {
  const guides: string[] = [];
  const commands: string[] = [];
  for (let k = 0; k < corners; k++) {
    const angle = Math.round(16200000 + (k * 21600000) / corners);
    guides.push(`dx${k} cos wd2 ${angle}`, `dy${k} sin hd2 ${angle}`, `x${k} +- hc dx${k} 0`, `y${k} +- vc dy${k} 0`);
    commands.push(`${k === 0 ? 'M' : 'L'} x${k} y${k}`);
  }
  guides.push('ix cos wd2 2700000', 'iy sin hd2 2700000', 'il +- hc 0 ix', 'ir +- hc ix 0', 'it +- vc 0 iy', 'ib +- vc iy 0');
  return fromSource({ gd: guides.join('; '), path: `${commands.join(' ')} Z`, rect: 'il it ir ib' });
}

/** avLst entries of a prstGeom or custGeom: <a:gd name="adj" fmla="val 25000"/>. */
export function formulasOf(list: Element | undefined): Formula[] {
  return children(list, 'gd')
    .slice(0, PPTX_LIMITS.maxGuides)
    .map((gd) => {
      const [op = 'val', ...args] = (attr(gd, 'fmla') ?? '').trim().split(/\s+/);
      return { name: attr(gd, 'name') ?? '', op, args };
    });
}

export function customGeometry(custGeom: Element): GeometryDef {
  const rect = child(custGeom, 'rect');
  let budget = PPTX_LIMITS.maxPathCommands;
  const paths: PathDef[] = [];
  for (const path of children(child(custGeom, 'pathLst'), 'path')) {
    const commands: PathCommand[] = [];
    for (const node of kids(path)) {
      if (budget-- <= 0) {
        break;
      }
      const points = children(node, 'pt').flatMap((pt) => [attr(pt, 'x') ?? '0', attr(pt, 'y') ?? '0']);
      switch (node.localName) {
        case 'moveTo':
          commands.push({ op: 'M', args: points.slice(0, 2) });
          break;
        case 'lnTo':
          commands.push({ op: 'L', args: points.slice(0, 2) });
          break;
        case 'arcTo':
          commands.push({ op: 'A', args: ['wR', 'hR', 'stAng', 'swAng'].map((name) => attr(node, name) ?? '0') });
          break;
        case 'quadBezTo':
          commands.push({ op: 'Q', args: points.slice(0, 4) });
          break;
        case 'cubicBezTo':
          commands.push({ op: 'C', args: points.slice(0, 6) });
          break;
        case 'close':
          commands.push({ op: 'Z', args: [] });
          break;
        default:
          break;
      }
    }
    const fill = attr(path, 'fill') as FillMode | undefined;
    paths.push({
      w: num(path, 'w'),
      h: num(path, 'h'),
      fill: fill && FILL_MODES.includes(fill) ? fill : 'norm',
      stroke: attr(path, 'stroke') !== '0' && attr(path, 'stroke') !== 'false',
      commands,
    });
  }
  return {
    adjust: formulasOf(child(custGeom, 'avLst')),
    guides: formulasOf(child(custGeom, 'gdLst')),
    paths,
    textRect: rect ? [attr(rect, 'l') ?? 'l', attr(rect, 't') ?? 't', attr(rect, 'r') ?? 'r', attr(rect, 'b') ?? 'b'] : undefined,
  };
}

function builtins(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h);
  const values: [string, number][] = [
    ['w', w],
    ['h', h],
    ['l', 0],
    ['t', 0],
    ['r', w],
    ['b', h],
    ['hc', w / 2],
    ['vc', h / 2],
    ['ss', ss],
    ['ls', Math.max(w, h)],
    ['cd2', 10800000],
    ['cd4', 5400000],
    ['cd8', 2700000],
    ['3cd4', 16200000],
    ['3cd8', 8100000],
    ['5cd8', 13500000],
    ['7cd8', 18900000],
  ];
  for (const divisor of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
    values.push([`wd${divisor}`, w / divisor], [`hd${divisor}`, h / divisor], [`ssd${divisor}`, ss / divisor]);
  }
  return new Map(values);
}

/** Evaluates the adjust values (with the shape's overrides) and guides, in EMUs. */
export function evaluateGuides(def: GeometryDef, w: number, h: number, overrides: readonly Formula[] = []): Map<string, number> {
  const values = builtins(w, h);
  const byName = new Map(overrides.map((formula) => [formula.name, formula]));
  let budget = PPTX_LIMITS.maxGuides;
  for (const formula of def.adjust) {
    if (budget-- <= 0) {
      break;
    }
    values.set(formula.name, calculate(byName.get(formula.name) ?? formula, values));
  }
  for (const formula of def.guides) {
    if (budget-- <= 0) {
      break;
    }
    values.set(formula.name, calculate(formula, values));
  }
  return values;
}

export function calculate(formula: Formula, values: ReadonlyMap<string, number>): number {
  const arg = (index: number) => {
    const token = formula.args[index] ?? '0';
    const known = values.get(token);
    if (known !== undefined) {
      return known;
    }
    const parsed = Number(token);
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const [x, y, z] = [arg(0), arg(1), arg(2)];
  let result: number;
  switch (formula.op) {
    case '*/':
      result = z === 0 ? 0 : (x * y) / z;
      break;
    case '+-':
      result = x + y - z;
      break;
    case '+/':
      result = z === 0 ? 0 : (x + y) / z;
      break;
    case '?:':
      result = x > 0 ? y : z;
      break;
    case 'abs':
      result = Math.abs(x);
      break;
    case 'at2':
      result = Math.atan2(y, x) / ANGLE_UNIT;
      break;
    case 'cat2':
      result = x * Math.cos(Math.atan2(z, y));
      break;
    case 'cos':
      result = x * Math.cos(y * ANGLE_UNIT);
      break;
    case 'max':
      result = Math.max(x, y);
      break;
    case 'min':
      result = Math.min(x, y);
      break;
    case 'mod':
      result = Math.sqrt(x * x + y * y + z * z);
      break;
    case 'pin':
      result = y < x ? x : y > z ? z : y;
      break;
    case 'sat2':
      result = x * Math.sin(Math.atan2(z, y));
      break;
    case 'sin':
      result = x * Math.sin(y * ANGLE_UNIT);
      break;
    case 'sqrt':
      result = Math.sqrt(Math.max(0, x));
      break;
    case 'tan':
      result = x * Math.tan(y * ANGLE_UNIT);
      break;
    case 'val':
      result = x;
      break;
    default:
      result = 0;
  }
  return Number.isFinite(result) ? result : 0;
}

/** Builds the paths of `def` for a shape of `w` × `h` EMUs, in CSS pixels. */
export function buildGeometry(def: GeometryDef, w: number, h: number, overrides: readonly Formula[] = []): BuiltGeometry {
  const values = evaluateGuides(def, w, h, overrides);
  const value = (token: string | undefined) => {
    const known = token === undefined ? 0 : values.get(token);
    if (known !== undefined) {
      return known;
    }
    const parsed = Number(token);
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const widthPx = w / EMU_PER_PX;
  const heightPx = h / EMU_PER_PX;
  const paths = def.paths.map((path) => {
    const sx = path.w ? widthPx / path.w : 1 / EMU_PER_PX;
    const sy = path.h ? heightPx / path.h : 1 / EMU_PER_PX;
    return { d: pathData(path.commands, value, sx, sy), fill: path.fill, stroke: path.stroke };
  });
  const [l = 'l', t = 't', r = 'r', b = 'b'] = def.textRect ?? [];
  const x = value(l) / EMU_PER_PX;
  const y = value(t) / EMU_PER_PX;
  return { paths, textRect: { x, y, w: value(r) / EMU_PER_PX - x, h: value(b) / EMU_PER_PX - y } };
}

function pathData(commands: readonly PathCommand[], value: (token: string | undefined) => number, sx: number, sy: number): string {
  const parts: string[] = [];
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  const point = (args: readonly string[], index: number): [number, number] => [value(args[index]) * sx, value(args[index + 1]) * sy];
  for (const command of commands) {
    switch (command.op) {
      case 'M':
        [x, y] = point(command.args, 0);
        [startX, startY] = [x, y];
        parts.push(`M${fmt(x)} ${fmt(y)}`);
        break;
      case 'L':
        [x, y] = point(command.args, 0);
        parts.push(`L${fmt(x)} ${fmt(y)}`);
        break;
      case 'Q': {
        const [x1, y1] = point(command.args, 0);
        [x, y] = point(command.args, 2);
        parts.push(`Q${fmt(x1)} ${fmt(y1)} ${fmt(x)} ${fmt(y)}`);
        break;
      }
      case 'C': {
        const [x1, y1] = point(command.args, 0);
        const [x2, y2] = point(command.args, 2);
        [x, y] = point(command.args, 4);
        parts.push(`C${fmt(x1)} ${fmt(y1)} ${fmt(x2)} ${fmt(y2)} ${fmt(x)} ${fmt(y)}`);
        break;
      }
      case 'A': {
        const wR = value(command.args[0]) * sx;
        const hR = value(command.args[1]) * sy;
        const arc = arcSegments(x, y, wR, hR, value(command.args[2]), value(command.args[3]));
        parts.push(arc.d);
        [x, y] = [arc.x, arc.y];
        break;
      }
      case 'Z':
        parts.push('Z');
        [x, y] = [startX, startY];
        break;
    }
  }
  return parts.join(' ');
}

/**
 * DrawingML arcTo: the current point lies on the ellipse at visual angle stAng; the arc sweeps swAng (positive is
 * clockwise on screen). Angles are visual, so they are converted to the ellipse's parametric angles.
 */
export function arcSegments(x: number, y: number, wR: number, hR: number, stAng: number, swAng: number): { d: string; x: number; y: number } {
  if (wR <= 0 || hR <= 0 || swAng === 0) {
    return { d: '', x, y };
  }
  const parametric = (angle: number) => {
    const radians = angle * ANGLE_UNIT;
    return Math.atan2(wR * Math.sin(radians), hR * Math.cos(radians));
  };
  const start = parametric(stAng);
  const cx = x - wR * Math.cos(start);
  const cy = y - hR * Math.sin(start);
  // A full turn (or more) is drawn as two halves: SVG cannot draw an arc back to its own start point.
  const pieces = Math.abs(swAng) >= 21600000 ? 2 : 1;
  const sweep = Math.max(-21600000, Math.min(21600000, swAng)) / pieces;
  const segments: string[] = [];
  let endX = x;
  let endY = y;
  for (let i = 1; i <= pieces; i++) {
    const end = parametric(stAng + sweep * i);
    endX = cx + wR * Math.cos(end);
    endY = cy + hR * Math.sin(end);
    const large = Math.abs(sweep) > 10800000 ? 1 : 0;
    const clockwise = sweep > 0 ? 1 : 0;
    segments.push(`A${fmt(wR)} ${fmt(hR)} 0 ${large} ${clockwise} ${fmt(endX)} ${fmt(endY)}`);
  }
  return { d: segments.join(' '), x: endX, y: endY };
}

function fmt(value: number): string {
  return String(Math.round(value * 100) / 100);
}
