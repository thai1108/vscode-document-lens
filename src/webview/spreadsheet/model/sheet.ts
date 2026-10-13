// Builds the format-neutral Sheet the grid shows. Readers (.xlsx, .xls, and any later one such as CSV) only decide
// what each cell holds; the builder stores it compactly, applies the caps and works out the used range and notices.
import { SHEET_LIMITS } from '../limits.ts';
import type { CellRange, CellRef } from './address.ts';
import { CellKind, CellStoreBuilder, NO_FORMULA, type CellStore } from './cells.ts';
import { shiftFormula } from './formula.ts';
import type { AxisSizes, CellFormat, Hyperlink, RichRun, Sheet, SheetKind } from './types.ts';

export interface SheetBuilderOptions {
  /** Cells kept before the rest of the sheet is cut off. */
  maxCells?: number;
  /** Likely number of cells, to size the storage up front. */
  expectedCells?: number;
}

/** The parts of a Sheet that are not cells. */
export interface SheetLayout {
  rows: AxisSizes;
  cols: AxisSizes;
  frozen?: { rows: number; cols: number };
  topLeft?: CellRef;
  activeCell?: CellRef;
  showGridLines?: boolean;
  /** Where the sheet came from, e.g. `xl/worksheets/sheet2.xml`. */
  sourcePart?: string;
}

export class SheetBuilder {
  readonly name: string;
  readonly kind: SheetKind;
  readonly merges: CellRange[] = [];
  readonly links: Hyperlink[] = [];
  /** Created with the first cell, so a size hint can still arrive before it. */
  private cells: CellStoreBuilder | undefined;
  private expectedCells: number | undefined;
  private readonly maxCells: number;
  /** Formats whose blank cells still show (a fill or a border). */
  private readonly visibleBlank: boolean[];
  private readonly strings: string[] = [];
  private readonly rich = new Map<number, RichRun[]>();
  private readonly formulas: string[] = [];
  private readonly sharedFormulas = new Map<number, { formula: string; row: number; col: number }>();
  private readonly notices: string[] = [];
  private usedRows = 0;
  private usedCols = 0;
  private cutAt: number | undefined;
  private hasDrawings = false;

  constructor(name: string, kind: SheetKind, formats: readonly CellFormat[], options: SheetBuilderOptions = {}) {
    this.name = name;
    this.kind = kind;
    this.maxCells = options.maxCells ?? SHEET_LIMITS.maxCells;
    this.expectedCells = options.expectedCells;
    this.visibleBlank = formats.map((format) => Boolean(format.fill || format.borders.top || format.borders.right || format.borders.bottom || format.borders.left));
  }

  /** Whether the cell cap has been reached; later cells are ignored. */
  get cutOff(): boolean {
    return this.cutAt !== undefined;
  }

  /** How many cells are likely to come (a hint from the file, so it is bounded); only before the first cell. */
  expectCells(count: number): void {
    this.expectedCells ??= Math.max(0, Math.min(count, 1 << 20));
  }

  /** A string the sheet's own cells refer to (CellKind.SheetString or CellKind.Error); its index. */
  addString(text: string, runs?: RichRun[]): number {
    if (runs) {
      this.rich.set(this.strings.length, runs);
    }
    return this.strings.push(text.length > SHEET_LIMITS.maxStringLength ? text.slice(0, SHEET_LIMITS.maxStringLength) : text) - 1;
  }

  /** A formula's text (without `=`), as the formula reference of a cell; NO_FORMULA past the cap. */
  addFormula(text: string): number {
    return text && this.formulas.length < SHEET_LIMITS.maxFormulas ? this.formulas.push(text) - 1 : NO_FORMULA;
  }

  /** The formula `si` shared by several cells, written once on the cell at (row, col). */
  defineSharedFormula(si: number, text: string, row: number, col: number): void {
    if (!this.sharedFormulas.has(si)) {
      this.sharedFormulas.set(si, { formula: text, row, col });
    }
  }

  /** The formula reference of a cell that uses shared formula `si` (moved to the cell when read back). */
  sharedFormula(si: number): number {
    return this.sharedFormulas.has(si) ? -2 - si : NO_FORMULA;
  }

  /**
   * Adds a cell. Blank cells are kept only when their format shows or they have a formula. Returns false once the
   * sheet has been cut off at the cell cap.
   */
  add(row: number, col: number, kind: CellKind, value: number, style: number, formula = NO_FORMULA): boolean {
    if (this.cutAt !== undefined) {
      return false;
    }
    const shows = kind !== CellKind.Blank || this.visibleBlank[style] === true;
    if (!shows && formula === NO_FORMULA) {
      return true;
    }
    this.cells ??= new CellStoreBuilder(this.maxCells, this.expectedCells);
    if (!this.cells.add(row, col, kind, value, style, formula)) {
      this.cutAt = row;
      this.notices.push(`Only the first ${row.toLocaleString('en-US')} rows of “${this.name}” are shown: it has more than ${this.maxCells.toLocaleString('en-US')} cells.`);
      return false;
    }
    if (shows) {
      this.usedRows = Math.max(this.usedRows, row + 1);
      this.usedCols = Math.max(this.usedCols, col + 1);
    }
    return true;
  }

  addMerge(range: CellRange): void {
    if (this.merges.length < SHEET_LIMITS.maxMerges && (range.top !== range.bottom || range.left !== range.right)) {
      this.merges.push(range);
    }
  }

  addLink(link: Hyperlink): void {
    if (this.links.length < SHEET_LIMITS.maxLinks) {
      this.links.push(link);
    }
  }

  /** The sheet has pictures, charts or shapes, which the preview does not show. */
  noteDrawings(): void {
    this.hasDrawings = true;
  }

  notice(text: string): void {
    this.notices.push(text);
  }

  build(layout: SheetLayout): Sheet {
    const cells = (this.cells ?? new CellStoreBuilder(0, 0)).build();
    let usedRows = this.usedRows;
    let usedCols = this.usedCols;
    for (const merge of this.merges) {
      usedRows = Math.max(usedRows, merge.bottom + 1);
      usedCols = Math.max(usedCols, merge.right + 1);
    }
    const notices = [...this.notices];
    if (this.hasDrawings) {
      notices.push(`Pictures, charts and shapes in “${this.name}” are not shown.`);
    }
    return {
      name: this.name,
      kind: this.kind,
      cells,
      strings: { texts: this.strings, rich: this.rich },
      formulaText: formulaLookup(cells, this.formulas, this.sharedFormulas),
      rows: layout.rows,
      cols: layout.cols,
      used: { rows: usedRows, cols: usedCols },
      merges: this.merges,
      links: this.links,
      frozen: layout.frozen ?? { rows: 0, cols: 0 },
      ...(layout.topLeft ? { topLeft: layout.topLeft } : {}),
      ...(layout.activeCell ? { activeCell: layout.activeCell } : {}),
      showGridLines: layout.showGridLines ?? true,
      ...(layout.sourcePart ? { sourcePart: layout.sourcePart } : {}),
      notices,
    };
  }
}

/** A sheet with no cells, e.g. a chart sheet, showing `notice`. */
export function emptySheet(name: string, kind: SheetKind, defaultRow: number, defaultCol: number, notice: string): Sheet {
  const builder = new SheetBuilder(name, kind, []);
  builder.notice(notice);
  return builder.build({ rows: { defaultSize: defaultRow, sizes: new Map() }, cols: { defaultSize: defaultCol, sizes: new Map() } });
}

function formulaLookup(
  cells: CellStore,
  formulas: readonly string[],
  shared: ReadonlyMap<number, { formula: string; row: number; col: number }>,
): (index: number, row: number, col: number) => string | undefined {
  return (index, row, col) => {
    const ref = cells.formula(index);
    if (ref === NO_FORMULA) {
      return undefined;
    }
    if (ref >= 0) {
      return formulas[ref];
    }
    const master = shared.get(-2 - ref);
    return master ? shiftFormula(master.formula, row - master.row, col - master.col) : undefined;
  };
}
