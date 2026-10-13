// Positions along one axis of the grid (rows or columns) at 100% zoom: the entries the sheet sizes itself, then any
// number of default-sized ones, so the grid can extend past the used range without storing anything for it.
import type { AxisSizes } from '../model/types.ts';

export class Axis {
  readonly defaultSize: number;
  /** offsets[i] is the start of entry i, for i up to the last entry with its own size. */
  private readonly offsets: Float64Array;

  /** Entries before `hiddenBefore` have no size: rows or columns scrolled out of view above or left of a frozen pane. */
  constructor(sizes: AxisSizes, limit: number, hiddenBefore = 0) {
    this.defaultSize = Math.max(0, sizes.defaultSize);
    const hidden = Math.max(0, Math.min(hiddenBefore, limit));
    let last = hidden - 1;
    for (const index of sizes.sizes.keys()) {
      if (index < limit && index > last) {
        last = index;
      }
    }
    this.offsets = new Float64Array(last + 2);
    for (let i = 0; i <= last; i++) {
      const size = i < hidden ? 0 : sizes.sizes.get(i);
      this.offsets[i + 1] = (this.offsets[i] ?? 0) + (size === undefined ? this.defaultSize : Math.max(0, size));
    }
  }

  /** Start of entry `index`. */
  start(index: number): number {
    const explicit = this.offsets.length - 1;
    if (index <= explicit) {
      return this.offsets[Math.max(0, index)] ?? 0;
    }
    return (this.offsets[explicit] ?? 0) + (index - explicit) * this.defaultSize;
  }

  size(index: number): number {
    return this.start(index + 1) - this.start(index);
  }

  /**
   * The entry containing `offset` (the first one whose end is past it), skipping zero-size (hidden) entries; offsets
   * beyond `count` give `count - 1`.
   */
  indexAt(offset: number, count: number): number {
    if (count <= 0) {
      return 0;
    }
    const explicit = this.offsets.length - 1;
    let index: number;
    if (offset < (this.offsets[explicit] ?? 0)) {
      // Largest i with offsets[i] <= offset, then past any hidden entries that also start there.
      let low = 0;
      let high = explicit;
      while (low < high) {
        const mid = (low + high + 1) >>> 1;
        if ((this.offsets[mid] ?? 0) <= offset) {
          low = mid;
        } else {
          high = mid - 1;
        }
      }
      index = low;
      while (index < explicit && this.size(index) === 0) {
        index++;
      }
    } else {
      index = this.defaultSize > 0 ? explicit + Math.floor((offset - (this.offsets[explicit] ?? 0)) / this.defaultSize) : explicit;
    }
    return Math.max(0, Math.min(count - 1, index));
  }

  /** How many entries from `first` on start within `length` (a page for PageUp/PageDown); at least 1. */
  countUntil(first: number, length: number, count: number): number {
    const end = this.indexAt(this.start(first) + length - 1e-6, count);
    return Math.max(1, end - first + 1);
  }
}
