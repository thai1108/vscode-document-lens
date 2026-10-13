// Compact cell storage. A sheet can hold millions of cells, so instead of one object per cell the values live in
// parallel typed arrays ordered by row, then column, with an index of where each row starts.

export const CellKind = {
  /** No value; kept only because its format shows (fill or border). */
  Blank: 0,
  Number: 1,
  /** Index into the workbook's shared strings. */
  SharedString: 2,
  /** Index into the sheet's own strings (inline strings, formula text results). */
  SheetString: 3,
  Boolean: 4,
  /** Index into the sheet's own strings, holding the error text such as `#DIV/0!`. */
  Error: 5,
} as const;
export type CellKind = (typeof CellKind)[keyof typeof CellKind];

/** Formula reference of a cell without one. */
export const NO_FORMULA = -1;

export class CellStore {
  /** Number of stored cells. */
  readonly size: number;
  /** One past the last row that has a stored cell. */
  readonly rowCount: number;
  private readonly rowStart: Int32Array;
  private readonly cols: Uint16Array;
  private readonly kinds: Uint8Array;
  private readonly values: Float64Array;
  private readonly styles: Uint16Array;
  private readonly formulas: Int32Array | null;

  constructor(rowStart: Int32Array, cols: Uint16Array, kinds: Uint8Array, values: Float64Array, styles: Uint16Array, formulas: Int32Array | null) {
    this.rowStart = rowStart;
    this.rowCount = rowStart.length - 1;
    this.size = cols.length;
    this.cols = cols;
    this.kinds = kinds;
    this.values = values;
    this.styles = styles;
    this.formulas = formulas;
  }

  static empty(): CellStore {
    return new CellStore(new Int32Array(1), new Uint16Array(0), new Uint8Array(0), new Float64Array(0), new Uint16Array(0), null);
  }

  /** Index of the first cell of `row` (inclusive) and of the next row (exclusive). */
  rowStartIndex(row: number): number {
    return row < this.rowCount ? (this.rowStart[row] ?? this.size) : this.size;
  }

  rowEndIndex(row: number): number {
    return row < this.rowCount ? (this.rowStart[row + 1] ?? this.size) : this.size;
  }

  /** Index of the first cell in `row` whose column is at least `col` (the row's end index if none). */
  lowerBound(row: number, col: number): number {
    let low = this.rowStartIndex(row);
    let high = this.rowEndIndex(row);
    while (low < high) {
      const mid = (low + high) >>> 1;
      if ((this.cols[mid] ?? 0) < col) {
        low = mid + 1;
      } else {
        high = mid;
      }
    }
    return low;
  }

  /** Index of the cell at (row, col), or -1. */
  find(row: number, col: number): number {
    const index = this.lowerBound(row, col);
    return index < this.rowEndIndex(row) && this.cols[index] === col ? index : -1;
  }

  col(index: number): number {
    return this.cols[index] ?? 0;
  }

  kind(index: number): CellKind {
    return (this.kinds[index] ?? CellKind.Blank) as CellKind;
  }

  value(index: number): number {
    return this.values[index] ?? 0;
  }

  style(index: number): number {
    return this.styles[index] ?? 0;
  }

  formula(index: number): number {
    return this.formulas?.[index] ?? NO_FORMULA;
  }
}

/**
 * Collects cells in any order (rows ascending, then columns, is the cheap path every Excel writer takes) and builds
 * a CellStore. Stops accepting cells at `maxCells`.
 */
export class CellStoreBuilder {
  private readonly maxCells: number;
  private count = 0;
  private capacity: number;
  private rows: Uint32Array;
  private cols: Uint16Array;
  private kinds: Uint8Array;
  private values: Float64Array;
  private styles: Uint16Array;
  private formulas: Int32Array | null = null;
  private sorted = true;
  private lastRow = -1;
  private lastCol = -1;

  constructor(maxCells: number, expected = 1024) {
    this.maxCells = maxCells;
    this.capacity = Math.max(16, Math.min(maxCells, expected));
    this.rows = new Uint32Array(this.capacity);
    this.cols = new Uint16Array(this.capacity);
    this.kinds = new Uint8Array(this.capacity);
    this.values = new Float64Array(this.capacity);
    this.styles = new Uint16Array(this.capacity);
  }

  get size(): number {
    return this.count;
  }

  get full(): boolean {
    return this.count >= this.maxCells;
  }

  /** Adds a cell; false once the cap is reached. */
  add(row: number, col: number, kind: CellKind, value: number, style: number, formula = NO_FORMULA): boolean {
    if (this.count >= this.maxCells) {
      return false;
    }
    if (this.count === this.capacity) {
      this.grow();
    }
    const i = this.count++;
    if (row < this.lastRow || (row === this.lastRow && col <= this.lastCol)) {
      this.sorted = false;
    }
    this.lastRow = row;
    this.lastCol = col;
    this.rows[i] = row;
    this.cols[i] = col;
    this.kinds[i] = kind;
    this.values[i] = value;
    this.styles[i] = style;
    if (formula !== NO_FORMULA) {
      if (!this.formulas) {
        this.formulas = new Int32Array(this.capacity).fill(NO_FORMULA);
      }
      this.formulas[i] = formula;
    }
    return true;
  }

  build(): CellStore {
    const n = this.count;
    let order: Uint32Array | undefined;
    if (!this.sorted) {
      order = this.sortedOrder();
    }
    const pick = <T extends Uint32Array | Uint16Array | Uint8Array | Float64Array | Int32Array>(source: T, create: (length: number) => T): T => {
      if (!order) {
        return source.subarray(0, n) as T;
      }
      const target = create(order.length);
      for (let i = 0; i < order.length; i++) {
        target[i] = source[order[i] ?? 0] ?? 0;
      }
      return target;
    };
    const rows = pick(this.rows, (length) => new Uint32Array(length));
    const cols = pick(this.cols, (length) => new Uint16Array(length));
    const kinds = pick(this.kinds, (length) => new Uint8Array(length));
    const values = pick(this.values, (length) => new Float64Array(length));
    const styles = pick(this.styles, (length) => new Uint16Array(length));
    const formulas = this.formulas ? pick(this.formulas, (length) => new Int32Array(length)) : null;
    const size = cols.length;
    const rowCount = size ? (rows[size - 1] ?? 0) + 1 : 0;
    const rowStart = new Int32Array(rowCount + 1);
    for (let i = 0; i < size; i++) {
      const next = (rows[i] ?? 0) + 1;
      rowStart[next] = (rowStart[next] ?? 0) + 1;
    }
    for (let r = 0; r < rowCount; r++) {
      rowStart[r + 1] = (rowStart[r + 1] ?? 0) + (rowStart[r] ?? 0);
    }
    return new CellStore(rowStart, cols, kinds, values, styles, formulas);
  }

  private grow(): void {
    const capacity = Math.min(this.maxCells, this.capacity * 2);
    const grow = <T extends Uint32Array | Uint16Array | Uint8Array | Float64Array | Int32Array>(source: T, create: (length: number) => T): T => {
      const target = create(capacity);
      target.set(source);
      return target;
    };
    this.rows = grow(this.rows, (length) => new Uint32Array(length));
    this.cols = grow(this.cols, (length) => new Uint16Array(length));
    this.kinds = grow(this.kinds, (length) => new Uint8Array(length));
    this.values = grow(this.values, (length) => new Float64Array(length));
    this.styles = grow(this.styles, (length) => new Uint16Array(length));
    if (this.formulas) {
      const formulas = new Int32Array(capacity).fill(NO_FORMULA);
      formulas.set(this.formulas);
      this.formulas = formulas;
    }
    this.capacity = capacity;
  }

  /** Cell indices sorted by (row, col); of cells written twice to the same place, the last one wins. */
  private sortedOrder(): Uint32Array {
    const { rows, cols } = this;
    const all = new Uint32Array(this.count);
    for (let i = 0; i < all.length; i++) {
      all[i] = i;
    }
    all.sort((a, b) => (rows[a] ?? 0) - (rows[b] ?? 0) || (cols[a] ?? 0) - (cols[b] ?? 0) || a - b);
    const order: number[] = [];
    for (let i = 0; i < all.length; i++) {
      const cell = all[i] ?? 0;
      const next = all[i + 1];
      if (next !== undefined && rows[next] === rows[cell] && cols[next] === cols[cell]) {
        continue;
      }
      order.push(cell);
    }
    return Uint32Array.from(order);
  }
}
