// Cell ranges (merged cells, links) looked up by position. Each range is listed in the 64 × 64-cell tiles it covers;
// a range covering more tiles than that is kept apart and checked on every lookup, so memory stays proportional to
// the number of ranges whatever their size.
import type { CellRange } from './address.ts';

const TILE = 64;
/** Ranges covering more tiles than this are "large". */
const MAX_TILES = 16;
/** Large ranges kept; a real sheet has a handful. */
const MAX_LARGE = 1024;
/** A lookup over more tiles than this checks every range instead. */
const MAX_LOOKUP_TILES = 1024;

export class RangeIndex<T> {
  private readonly items: T[] = [];
  private readonly ranges: CellRange[] = [];
  private readonly tiles = new Map<number, number[]>();
  private readonly large: number[] = [];
  private readonly rangeOf: (item: T) => CellRange;

  constructor(rangeOf: (item: T) => CellRange, items: Iterable<T> = []) {
    this.rangeOf = rangeOf;
    for (const item of items) {
      this.add(item);
    }
  }

  get size(): number {
    return this.items.length;
  }

  /** Adds an item; false when it is a large range and the index already holds as many as it keeps. */
  add(item: T): boolean {
    const range = this.rangeOf(item);
    const tileRows = tile(range.bottom) - tile(range.top) + 1;
    const tileCols = tile(range.right) - tile(range.left) + 1;
    if (tileRows * tileCols > MAX_TILES && this.large.length >= MAX_LARGE) {
      return false;
    }
    const index = this.items.push(item) - 1;
    this.ranges.push(range);
    if (tileRows * tileCols > MAX_TILES) {
      this.large.push(index);
      return true;
    }
    for (let r = tile(range.top); r <= tile(range.bottom); r++) {
      for (let c = tile(range.left); c <= tile(range.right); c++) {
        const key = r * 65_536 + c;
        let list = this.tiles.get(key);
        if (!list) {
          list = [];
          this.tiles.set(key, list);
        }
        list.push(index);
      }
    }
    return true;
  }

  /** Adds an item unless its range overlaps one already added, as Excel drops overlapping merged cells. */
  addIfFree(item: T): boolean {
    const range = this.rangeOf(item);
    return !this.overlaps(range.top, range.bottom, range.left, range.right) && this.add(item);
  }

  /** Items whose range overlaps rows top..bottom and columns left..right, in the order they were added. */
  query(top: number, bottom: number, left: number, right: number): T[] {
    return this.indicesOverlapping(top, bottom, left, right).map((i) => this.items[i] as T);
  }

  /** The item added last whose range contains the cell. */
  at(row: number, col: number): T | undefined {
    const found = this.indicesOverlapping(row, row, col, col);
    const last = found[found.length - 1];
    return last === undefined ? undefined : this.items[last];
  }

  private overlaps(top: number, bottom: number, left: number, right: number): boolean {
    return this.indicesOverlapping(top, bottom, left, right, true).length > 0;
  }

  private indicesOverlapping(top: number, bottom: number, left: number, right: number, firstOnly = false): number[] {
    const hits = (i: number) => {
      const range = this.ranges[i];
      return range !== undefined && range.bottom >= top && range.top <= bottom && range.right >= left && range.left <= right;
    };
    const tileCount = (tile(bottom) - tile(top) + 1) * (tile(right) - tile(left) + 1);
    if (tileCount > MAX_LOOKUP_TILES) {
      const found: number[] = [];
      for (let i = 0; i < this.ranges.length && !(firstOnly && found.length); i++) {
        if (hits(i)) {
          found.push(i);
        }
      }
      return found;
    }
    const found = new Set<number>();
    for (const i of this.large) {
      if (hits(i)) {
        found.add(i);
        if (firstOnly) {
          return [i];
        }
      }
    }
    for (let r = tile(top); r <= tile(bottom); r++) {
      for (let c = tile(left); c <= tile(right); c++) {
        for (const i of this.tiles.get(r * 65_536 + c) ?? []) {
          if (!found.has(i) && hits(i)) {
            found.add(i);
            if (firstOnly) {
              return [i];
            }
          }
        }
      }
    }
    return [...found].sort((a, b) => a - b);
  }
}

function tile(index: number): number {
  return Math.floor(index / TILE);
}
