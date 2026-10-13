// Which rows or columns of the scrolling part of the grid are on screen. Along each axis the grid is: a header band,
// the frozen entries (always shown), then the scrolling entries; positions are header + zoom × axis offset.
import type { Axis } from './axis.ts';

export interface Span {
  first: number;
  last: number;
}

export interface AxisView {
  axis: Axis;
  /** Entries in the grid along this axis (used range plus margin). */
  count: number;
  frozen: number;
  /** Header band size in screen pixels (column header height, or row header width). */
  header: number;
  zoom: number;
}

/** Screen position of the start of entry `index` within the grid canvas. */
export function position(view: AxisView, index: number): number {
  return view.header + Math.round(view.axis.start(index) * view.zoom);
}

/** Size of the header plus the frozen entries, in screen pixels: what stays put along this axis. */
export function frozenExtent(view: AxisView): number {
  return position(view, view.frozen);
}

/** Total canvas size along this axis. */
export function canvasExtent(view: AxisView): number {
  return position(view, view.count);
}

/**
 * The scrolling entries visible at `scroll` in a viewport of `viewport` pixels, widened by `overscan` pixels on each
 * side. Never starts before the first scrolling entry; empty spans have last < first.
 */
export function scrollingSpan(view: AxisView, scroll: number, viewport: number, overscan = 0): Span {
  const { axis, count, frozen, zoom } = view;
  if (frozen >= count) {
    return { first: count, last: count - 1 };
  }
  const from = (scroll - overscan) / zoom + axis.start(frozen);
  const to = (scroll + viewport - frozenExtent(view) + overscan) / zoom + axis.start(frozen);
  if (to <= from) {
    return { first: frozen, last: frozen - 1 };
  }
  return {
    first: Math.max(frozen, axis.indexAt(from, count)),
    last: Math.max(frozen, axis.indexAt(Math.max(from, to - 1e-6), count)),
  };
}

/**
 * How many of the first `count` entries fit on a canvas at most `maxPx` long (header included) at `zoom`: all of
 * them, or those that end before the limit.
 */
export function countThatFits(axis: Axis, count: number, header: number, zoom: number, maxPx: number): number {
  if (header + axis.start(count) * zoom <= maxPx) {
    return count;
  }
  return axis.indexAt((maxPx - header) / zoom, count);
}

export function contains(outer: Span, inner: Span): boolean {
  return inner.last < inner.first || (outer.first <= inner.first && outer.last >= inner.last);
}

/** The scroll position that brings entry `index` into view (unchanged if it already is); frozen entries never move. */
export function scrollToShow(view: AxisView, index: number, scroll: number, viewport: number): number {
  if (index < view.frozen) {
    return scroll;
  }
  const visibleStart = scroll + frozenExtent(view);
  const start = position(view, index);
  const end = position(view, index + 1);
  if (start < visibleStart) {
    return Math.max(0, start - frozenExtent(view));
  }
  if (end > scroll + viewport) {
    return Math.max(0, Math.min(start - frozenExtent(view), end - viewport));
  }
  return scroll;
}
