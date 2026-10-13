// DrawingML measures: EMUs (914400 per inch), 60000ths of a degree, 1/100000 fractions; slides are laid out in CSS
// pixels at 96 per inch before the viewer zooms them.
import { bool, child, num } from './xml.ts';

export const EMU_PER_PX = 9525;
export const PX_PER_PT = 96 / 72;

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

/** Position and size in EMUs, rotation in degrees (clockwise). */
export interface Xfrm {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly rot: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

/** A group's frame plus the child coordinate space its members are positioned in. */
export interface GroupXfrm extends Xfrm {
  readonly chX: number;
  readonly chY: number;
  readonly chW: number;
  readonly chH: number;
}

// Positions far outside any slide only produce enormous layers; PowerPoint itself caps coordinates near this.
const MAX_COORDINATE = 51206400 * 4;

const coordinate = (value: number | undefined) => Math.max(-MAX_COORDINATE, Math.min(MAX_COORDINATE, value ?? 0));
const extent = (value: number | undefined) => Math.max(0, Math.min(MAX_COORDINATE, value ?? 0));

/** Reads a:xfrm / p:xfrm; undefined when the element or its offset/extent is missing. */
export function readXfrm(xfrm: Element | undefined): Xfrm | undefined {
  const off = child(xfrm, 'off');
  const ext = child(xfrm, 'ext');
  if (!xfrm || !off || !ext) {
    return undefined;
  }
  return {
    x: coordinate(num(off, 'x')),
    y: coordinate(num(off, 'y')),
    w: extent(num(ext, 'cx')),
    h: extent(num(ext, 'cy')),
    rot: ((num(xfrm, 'rot') ?? 0) / 60000) % 360,
    flipH: bool(xfrm, 'flipH') ?? false,
    flipV: bool(xfrm, 'flipV') ?? false,
  };
}

export function readGroupXfrm(xfrm: Element | undefined): GroupXfrm | undefined {
  const base = readXfrm(xfrm);
  if (!base) {
    return undefined;
  }
  const chOff = child(xfrm, 'chOff');
  const chExt = child(xfrm, 'chExt');
  return {
    ...base,
    chX: chOff ? coordinate(num(chOff, 'x')) : base.x,
    chY: chOff ? coordinate(num(chOff, 'y')) : base.y,
    chW: chExt ? extent(num(chExt, 'cx')) : base.w,
    chH: chExt ? extent(num(chExt, 'cy')) : base.h,
  };
}

/**
 * Maps a member's frame from its group's child coordinate space into the group's own box (both in EMUs, relative to
 * the group's top-left corner). Members keep their own rotation; the group's scale stretches positions and sizes.
 */
export function mapIntoGroup(member: Xfrm, group: GroupXfrm): Xfrm {
  const sx = group.chW > 0 ? group.w / group.chW : 1;
  const sy = group.chH > 0 ? group.h / group.chH : 1;
  return { ...member, x: (member.x - group.chX) * sx, y: (member.y - group.chY) * sy, w: member.w * sx, h: member.h * sy };
}
