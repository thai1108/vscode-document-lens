// Merged ranges looked up by row band, so drawing a window of the grid only visits the merges near it.
import type { CellRange } from '../model/address.ts';

const BAND = 64;
/** Merges taller than this many bands are checked on every query instead of being listed in each band. */
const MAX_BANDS = 256;

export class MergeIndex {
  private readonly merges: readonly CellRange[];
  private readonly bands = new Map<number, number[]>();
  private readonly tall: number[] = [];

  constructor(merges: readonly CellRange[]) {
    this.merges = merges;
    merges.forEach((merge, i) => {
      const first = Math.floor(merge.top / BAND);
      const last = Math.floor(merge.bottom / BAND);
      if (last - first >= MAX_BANDS) {
        this.tall.push(i);
        return;
      }
      for (let band = first; band <= last; band++) {
        let list = this.bands.get(band);
        if (!list) {
          list = [];
          this.bands.set(band, list);
        }
        list.push(i);
      }
    });
  }

  /** Merges overlapping rows top..bottom and columns left..right, in the order the sheet lists them. */
  query(top: number, bottom: number, left: number, right: number): CellRange[] {
    const found = new Set<number>();
    const visit = (i: number) => {
      const merge = this.merges[i];
      if (merge && merge.bottom >= top && merge.top <= bottom && merge.right >= left && merge.left <= right) {
        found.add(i);
      }
    };
    for (let band = Math.floor(top / BAND); band <= Math.floor(bottom / BAND); band++) {
      this.bands.get(band)?.forEach(visit);
    }
    this.tall.forEach(visit);
    return [...found].sort((a, b) => a - b).map((i) => this.merges[i] as CellRange);
  }

  /** The merge containing a cell, if any. */
  at(row: number, col: number): CellRange | undefined {
    return this.query(row, row, col, col)[0];
  }
}
