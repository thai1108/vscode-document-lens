// The workbook model both readers (.xlsx and .xls) produce and the grid renders. Sizes are CSS pixels at 100% zoom;
// colors are `#rrggbb`.
import type { CellRange, CellRef } from './address.ts';
import type { CellStore } from './cells.ts';

export interface Font {
  name: string;
  /** Points. */
  size: number;
  bold: boolean;
  italic: boolean;
  underline: 'none' | 'single' | 'double';
  strike: boolean;
  color?: string;
  vertAlign?: 'superscript' | 'subscript';
}

export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'mediumDashed'
  | 'dotted'
  | 'hair'
  | 'double'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export interface BorderEdge {
  style: BorderStyle;
  color: string;
}

export interface Borders {
  top?: BorderEdge;
  right?: BorderEdge;
  bottom?: BorderEdge;
  left?: BorderEdge;
}

export type HorizontalAlignment = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

/** One cell format (an `xf` in the file), shared by every cell that uses it. */
export interface CellFormat {
  font: Font;
  fill?: string;
  borders: Borders;
  horizontal: HorizontalAlignment;
  vertical: VerticalAlignment;
  wrap: boolean;
  indent: number;
  numberFormat: string;
}

/** A piece of a rich text string; `font` holds only what the run sets, the rest comes from the cell. */
export interface RichRun {
  text: string;
  font?: Partial<Font>;
}

export interface StringTable {
  readonly texts: readonly string[];
  readonly rich: ReadonlyMap<number, readonly RichRun[]>;
}

export interface Hyperlink {
  range: CellRange;
  /** External http(s) target; other schemes are dropped when read. */
  url?: string;
  /** Place in the workbook, e.g. `Sheet2!A1`. */
  location?: string;
  tooltip?: string;
}

/** Column widths or row heights: a default size plus the entries that differ (0 for hidden). */
export interface AxisSizes {
  defaultSize: number;
  sizes: ReadonlyMap<number, number>;
}

/**
 * Frozen panes: `rows` rows from row `top` and `cols` columns from column `left` stay on screen. Rows above `top` and
 * columns left of `left` were scrolled out of the frozen pane in Excel and are not shown: Excel counts the frozen rows
 * from the sheet view's top-left cell, as LibreOffice's import also does.
 */
export interface FrozenPanes {
  rows: number;
  cols: number;
  top: number;
  left: number;
}

export const NO_FROZEN_PANES: FrozenPanes = { rows: 0, cols: 0, top: 0, left: 0 };

/** Frozen panes of `rows` rows and `cols` columns in a view whose top-left cell is `viewTopLeft`. */
export function frozenPanes(rows: number, cols: number, viewTopLeft: CellRef | undefined, maxRows: number, maxCols: number): FrozenPanes {
  const top = rows > 0 ? (viewTopLeft?.row ?? 0) : 0;
  const left = cols > 0 ? (viewTopLeft?.col ?? 0) : 0;
  const rowsFit = rows > 0 && top + rows <= maxRows;
  const colsFit = cols > 0 && left + cols <= maxCols;
  return { rows: rowsFit ? rows : 0, cols: colsFit ? cols : 0, top: rowsFit ? top : 0, left: colsFit ? left : 0 };
}

export type SheetKind = 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';

export interface SheetEntry {
  name: string;
  hidden: boolean;
  kind: SheetKind;
}

export interface Sheet {
  readonly name: string;
  readonly kind: SheetKind;
  readonly cells: CellStore;
  /** Strings the sheet's cells refer to with CellKind.SheetString or CellKind.Error. */
  readonly strings: StringTable;
  /** Formula texts the cells refer to (CellStore.formula), without the leading `=`. */
  formulaText(index: number, row: number, col: number): string | undefined;
  readonly rows: AxisSizes;
  readonly cols: AxisSizes;
  /** Rows and columns that hold something visible: values, formatted blanks, merges. */
  readonly used: { rows: number; cols: number };
  readonly merges: readonly CellRange[];
  readonly links: readonly Hyperlink[];
  readonly frozen: FrozenPanes;
  /** First cell shown in the scrolling part of the sheet when it opens. */
  readonly topLeft?: CellRef;
  readonly activeCell?: CellRef;
  readonly showGridLines: boolean;
  /**
   * The package part the sheet was read from (.xlsx), e.g. `xl/worksheets/sheet2.xml`. A cell in it is addressed by
   * its A1 reference, formatCellRef(row, col).
   */
  readonly sourcePart?: string;
  /** What the preview leaves out or cut off, for the notice banner. */
  readonly notices: readonly string[];
}

export interface Workbook {
  readonly sheets: readonly SheetEntry[];
  /** Index into `sheets` of the sheet Excel shows first. */
  readonly activeSheet: number;
  readonly date1904: boolean;
  /** Indexed by the style number stored with each cell; formats[0] is the default. */
  readonly formats: readonly CellFormat[];
  readonly sharedStrings: StringTable;
  readonly notices: readonly string[];
  loadSheet(index: number, signal?: AbortSignal): Promise<Sheet>;
}

export const DEFAULT_FONT: Font = { name: 'Calibri', size: 11, bold: false, italic: false, underline: 'none', strike: false };

export function defaultFormat(font: Font = DEFAULT_FONT): CellFormat {
  return { font, borders: {}, horizontal: 'general', vertical: 'bottom', wrap: false, indent: 0, numberFormat: 'General' };
}

export const EMPTY_STRINGS: StringTable = { texts: [], rich: new Map() };
