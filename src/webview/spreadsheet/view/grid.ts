// The sheet grid. Only the cells in and around the viewport exist in the DOM; they are redrawn when scrolling moves
// past them. The grid is split in four parts so frozen panes stay put with native scrolling:
//
//   canvas ─┬─ top band (sticky top: column headers, frozen rows) ── corner (sticky left: frozen rows × frozen cols)
//           └─ body (scrolling cells) ── left band (sticky left: row headers, frozen columns)
//
// Positions are integers in screen pixels: header + round(zoom × offset at 100%). Every cell box starts one pixel
// before its column/row so its left/top border lands on the gridline it shares with the neighbor.
import { externalUrl } from '../../../protocol.ts';
import { SHEET_LIMITS } from '../limits.ts';
import { columnName, formatRange, rangeContains, type CellRange, type CellRef } from '../model/address.ts';
import { CellKind } from '../model/cells.ts';
import { cellDisplay, readCell, type CellDisplay, type TextAlign } from '../model/display.ts';
import type { CellFormat, Hyperlink, Sheet, Workbook } from '../model/types.ts';
import { RangeIndex } from '../model/rangeIndex.ts';
import { Axis } from './axis.ts';
import { CellSelection, type NavigationBounds } from './selection.ts';
import { applyFont, borderCss, type StyleRules } from './styleRules.ts';
import { canvasExtent, contains, countThatFits, frozenExtent, position, scrollingSpan, scrollToShow, type AxisView, type Span } from './viewport.ts';

/** Column header height and row header digit width at 100%. */
const HEADER_HEIGHT = 20;
const HEADER_DIGIT_WIDTH = 7;
/** Empty rows and columns shown past the used range. */
const EXTRA_ROWS = 20;
const EXTRA_COLS = 5;
/** How far text may run over empty neighbors. */
const OVERFLOW_MAX_COLS = 64;
/** A frozen pane taking more of the viewport than this is ignored, or nothing could scroll. */
const MAX_FROZEN_SHARE = 0.8;
/** Narrowest a number is squeezed to fit its cell before it is shown as ####. */
const MIN_NUMBER_SQUEEZE = 0.8;
const EMPTY: Span = { first: 0, last: -1 };

type Quadrant = 'corner' | 'top' | 'left' | 'body';

/** Scroll position at 100% zoom and the selected cell, as saved in the webview state. */
export interface GridPosition {
  top: number;
  left: number;
  row?: number;
  col?: number;
}

/** What the grid needs from the viewer. Web links are `<a href>` elements the shell opens through the host. */
export interface GridHost {
  readonly workbook: Workbook;
  readonly rules: StyleRules;
  /** A link to a place in the workbook such as `Sheet2!A1`. */
  goTo(location: string): void;
  selectionChanged(): void;
  viewChanged(): void;
}

export class Grid {
  /** The scrolling element; the viewer puts it in the page. */
  readonly element: HTMLDivElement;
  private readonly host: GridHost;
  private readonly sheet: Sheet;
  private readonly formats: readonly CellFormat[];
  private readonly canvas: HTMLDivElement;
  private readonly topBand: HTMLDivElement;
  private readonly corner: HTMLDivElement;
  private readonly body: HTMLDivElement;
  private readonly leftBand: HTMLDivElement;
  private readonly rows: Axis;
  private readonly cols: Axis;
  private readonly merges: RangeIndex<CellRange>;
  private readonly links: RangeIndex<Hyperlink>;
  private zoom: number;
  private rowView: AxisView;
  private colView: AxisView;
  private renderedRows: Span = EMPTY;
  private renderedCols: Span = EMPTY;
  /** The selected cell, by sheet row and column; drawn by markSelection(). */
  readonly selection: CellSelection;
  private selectionBox: HTMLElement | undefined;
  private resizeObserver: ResizeObserver | undefined;
  /** Number texts drawn since the last fitNumbers(). */
  private numbers: HTMLElement[] = [];
  /** Rows of the used range left out because the grid would be taller than the browser can lay out. */
  cutOff = { rows: false, cols: false };
  /** First frozen row and column: those above and left of them are scrolled out of view, as in Excel. */
  private readonly frozenTop: number;
  private readonly frozenLeft: number;

  constructor(document: Document, host: GridHost, sheet: Sheet, zoom: number) {
    this.host = host;
    this.sheet = sheet;
    this.formats = host.workbook.formats;
    this.zoom = zoom;
    this.frozenTop = sheet.frozen.rows > 0 ? sheet.frozen.top : 0;
    this.frozenLeft = sheet.frozen.cols > 0 ? sheet.frozen.left : 0;
    this.rows = new Axis(sheet.rows, SHEET_LIMITS.maxRows, this.frozenTop);
    this.cols = new Axis(sheet.cols, SHEET_LIMITS.maxCols, this.frozenLeft);
    this.merges = new RangeIndex((range: CellRange) => range, sheet.merges);
    this.selection = new CellSelection(this.merges);
    this.links = new RangeIndex((link: Hyperlink) => link.range, sheet.links);
    const div = (className: string) => {
      const node = document.createElement('div');
      node.className = className;
      return node;
    };
    this.element = div('dl-grid');
    this.element.tabIndex = 0;
    // Not an ARIA grid: only the cells near the viewport exist. The cell bar announces the selected cell.
    this.element.setAttribute('role', 'region');
    this.element.setAttribute('aria-label', sheet.name);
    if (!sheet.showGridLines) {
      this.element.classList.add('dl-no-gridlines');
    }
    this.canvas = div('dl-canvas');
    this.topBand = div('dl-band-top');
    this.corner = div('dl-band-corner');
    this.body = div('dl-body');
    this.leftBand = div('dl-band-left');
    this.topBand.append(this.corner);
    this.body.append(this.leftBand);
    this.canvas.append(this.topBand, this.body);
    this.element.append(this.canvas);
    this.rowView = { axis: this.rows, count: 1, frozen: 0, header: 0, zoom };
    this.colView = { axis: this.cols, count: 1, frozen: 0, header: 0, zoom };

    this.element.addEventListener('scroll', this.onScroll);
    this.element.addEventListener('mousedown', this.onMouseDown);
    this.element.addEventListener('click', this.onClick);
    this.element.addEventListener('keydown', this.onKeyDown);
  }

  /** Lays the grid out once it is in the page, at a saved position or where the file says it was left. */
  start(saved?: GridPosition): void {
    this.layout();
    const ResizeObserverClass = this.element.ownerDocument.defaultView?.ResizeObserver;
    if (ResizeObserverClass) {
      let size = this.sizeKey();
      // The first callback comes right away, for the size just laid out.
      this.resizeObserver = new ResizeObserverClass(() => {
        if (this.sizeKey() !== size) {
          size = this.sizeKey();
          this.refresh();
        }
      });
      this.resizeObserver.observe(this.element);
    }
    if (saved) {
      if (saved.row !== undefined && saved.col !== undefined) {
        this.selection.select(saved.row, saved.col);
      }
      this.scrollTo(saved.top * this.zoom, saved.left * this.zoom);
    } else {
      const active = this.sheet.activeCell;
      if (active) {
        this.selection.select(active.row, active.col);
      }
      const topLeft = this.sheet.topLeft;
      this.scrollTo(
        topLeft && topLeft.row >= this.rowView.frozen ? position(this.rowView, topLeft.row) - frozenExtent(this.rowView) : 0,
        topLeft && topLeft.col >= this.colView.frozen ? position(this.colView, topLeft.col) - frozenExtent(this.colView) : 0,
      );
    }
    this.renderAll();
  }

  /** Scroll position at 100% zoom plus the selected cell. */
  savedPosition(): GridPosition {
    const saved: GridPosition = { top: this.element.scrollTop / this.zoom, left: this.element.scrollLeft / this.zoom };
    const active = this.selection.active;
    if (active) {
      saved.row = active.row;
      saved.col = active.col;
    }
    return saved;
  }

  setZoom(zoom: number): void {
    if (zoom === this.zoom) {
      return;
    }
    const { top, left } = this.savedPosition();
    this.zoom = zoom;
    this.layout();
    this.scrollTo(top * zoom, left * zoom);
    this.renderAll();
  }

  /** Re-lays out after the viewport changed size. */
  refresh(): void {
    this.layout();
    this.renderAll();
  }

  /** The selected cell's address (or merged range) and its formula or raw value, for the cell bar. */
  selectedInfo(): { address: string; text: string } | undefined {
    const range = this.selection.range;
    return range ? { address: formatRange(range), text: readCell(this.host.workbook, this.sheet, range.top, range.left)?.raw ?? '' } : undefined;
  }

  /** The selected cell's text as shown, for copying. */
  selectedText(): string | undefined {
    const range = this.selection.range;
    return range ? (readCell(this.host.workbook, this.sheet, range.top, range.left)?.text ?? '') : undefined;
  }

  /** Selects a cell and scrolls it into view. */
  goToCell(row: number, col: number): void {
    this.select(row, col);
    this.element.focus({ preventScroll: true });
  }

  focus(): void {
    this.element.focus({ preventScroll: true });
  }

  dispose(): void {
    this.resizeObserver?.disconnect();
    this.element.removeEventListener('scroll', this.onScroll);
    this.element.removeEventListener('mousedown', this.onMouseDown);
    this.element.removeEventListener('click', this.onClick);
    this.element.removeEventListener('keydown', this.onKeyDown);
    this.element.remove();
  }

  // Layout ------------------------------------------------------------------------------------------------------------

  /** The scrolling area's size; without layout (tests) the window's. */
  private viewport(): { width: number; height: number } {
    const window = this.element.ownerDocument.defaultView;
    return { width: this.element.clientWidth || window?.innerWidth || 800, height: this.element.clientHeight || window?.innerHeight || 600 };
  }

  private sizeKey(): string {
    const { width, height } = this.viewport();
    return `${width}x${height}`;
  }

  private layout(): void {
    const zoom = this.zoom;
    const { width, height } = this.viewport();
    const { used, frozen } = this.sheet;
    const headerHeight = Math.round(HEADER_HEIGHT * zoom);
    const fill = (axis: Axis, length: number, header: number, limit: number) => axis.indexAt((length - header) / zoom, limit) + 2;

    const frozenRows = frozen.rows > 0 ? this.frozenTop + frozen.rows : 0;
    const frozenCols = frozen.cols > 0 ? this.frozenLeft + frozen.cols : 0;
    // The grid cannot be larger than the browser lays out; past that, rows or columns are left out (with a notice).
    const wantedRows = Math.min(SHEET_LIMITS.maxRows, Math.max(used.rows + EXTRA_ROWS, frozenRows + 1, fill(this.rows, height, headerHeight, SHEET_LIMITS.maxRows)));
    const rowCount = countThatFits(this.rows, wantedRows, headerHeight, zoom, SHEET_LIMITS.maxCanvasPx);
    const rowHeaderWidth = Math.round((String(rowCount).length * HEADER_DIGIT_WIDTH + 12) * zoom);
    const wantedCols = Math.min(SHEET_LIMITS.maxCols, Math.max(used.cols + EXTRA_COLS, frozenCols + 1, fill(this.cols, width, rowHeaderWidth, SHEET_LIMITS.maxCols)));
    const colCount = countThatFits(this.cols, wantedCols, rowHeaderWidth, zoom, SHEET_LIMITS.maxCanvasPx);
    this.cutOff = { rows: rowCount < used.rows, cols: colCount < used.cols };

    this.rowView = { axis: this.rows, count: rowCount, frozen: Math.min(frozenRows, rowCount - 1), header: headerHeight, zoom };
    this.colView = { axis: this.cols, count: colCount, frozen: Math.min(frozenCols, colCount - 1), header: rowHeaderWidth, zoom };
    if (frozenExtent(this.rowView) > height * MAX_FROZEN_SHARE) {
      this.rowView.frozen = 0;
    }
    if (frozenExtent(this.colView) > width * MAX_FROZEN_SHARE) {
      this.colView.frozen = 0;
    }

    const canvasWidth = canvasExtent(this.colView);
    const canvasHeight = canvasExtent(this.rowView);
    const topHeight = frozenExtent(this.rowView);
    const leftWidth = frozenExtent(this.colView);
    this.element.style.setProperty('--dl-zoom', String(zoom));
    this.element.classList.toggle('dl-frozen-rows', this.rowView.frozen > 0);
    this.element.classList.toggle('dl-frozen-cols', this.colView.frozen > 0);
    setBox(this.canvas, undefined, undefined, canvasWidth, canvasHeight);
    setBox(this.topBand, undefined, undefined, undefined, topHeight);
    setBox(this.corner, undefined, undefined, leftWidth, topHeight);
    setBox(this.body, undefined, undefined, undefined, canvasHeight - topHeight);
    setBox(this.leftBand, undefined, undefined, leftWidth, canvasHeight - topHeight);
  }

  private scrollTo(top: number, left: number): void {
    this.element.scrollTop = Math.max(0, top);
    this.element.scrollLeft = Math.max(0, left);
  }

  // Rendering ---------------------------------------------------------------------------------------------------------

  private readonly onScroll = () => {
    this.update(false);
    this.host.viewChanged();
  };

  private visibleSpans(overscan: boolean): { rows: Span; cols: Span } {
    const { width, height } = this.viewport();
    return {
      rows: scrollingSpan(this.rowView, this.element.scrollTop, height, overscan ? height : 0),
      cols: scrollingSpan(this.colView, this.element.scrollLeft, width, overscan ? width / 2 : 0),
    };
  }

  private update(force: boolean): void {
    const visible = this.visibleSpans(false);
    const rowsChanged = force || !contains(this.renderedRows, visible.rows);
    const colsChanged = force || !contains(this.renderedCols, visible.cols);
    if (!rowsChanged && !colsChanged) {
      return;
    }
    const wanted = this.visibleSpans(true);
    if (rowsChanged) {
      this.renderedRows = wanted.rows;
    }
    if (colsChanged) {
      this.renderedCols = wanted.cols;
    }
    if (colsChanged) {
      this.renderTop();
    }
    if (rowsChanged) {
      this.renderLeft();
    }
    this.renderBody();
    this.fitNumbers();
    this.markSelection();
  }

  /**
   * Draws the selection box and highlights its row and column headers. It never replaces cells, so a click that
   * selects a cell still reaches the link in it.
   */
  private markSelection(): void {
    this.selectionBox?.remove();
    this.selectionBox = undefined;
    const selection = this.selection.range;
    for (const header of this.element.querySelectorAll<HTMLElement>('.dl-col-header, .dl-row-header')) {
      const index = Number(header.dataset['index']);
      const isColumn = header.classList.contains('dl-col-header');
      const selected = selection !== undefined && (isColumn ? index >= selection.left && index <= selection.right : index >= selection.top && index <= selection.bottom);
      header.classList.toggle('dl-selected', selected);
    }
    if (!selection) {
      return;
    }
    const quadrant = this.quadrantOf(selection.top, selection.left);
    const container = { corner: this.corner, top: this.topBand, left: this.leftBand, body: this.body }[quadrant];
    const rect = this.rangeRect(selection, quadrant === 'left' || quadrant === 'body' ? frozenExtent(this.rowView) : 0);
    this.selectionBox = this.box('dl-selection', rect.left, rect.top, rect.width, rect.height);
    container.append(this.selectionBox);
  }

  private renderAll(): void {
    this.renderedRows = EMPTY;
    this.renderedCols = EMPTY;
    this.renderCorner();
    this.update(true);
  }

  private frozenRows(): Span {
    return { first: this.frozenTop, last: this.rowView.frozen - 1 };
  }

  private frozenCols(): Span {
    return { first: this.frozenLeft, last: this.colView.frozen - 1 };
  }

  private renderCorner(): void {
    const parts: Node[] = [];
    const corner = this.box('dl-header dl-corner-header', 0, 0, this.colView.header, this.rowView.header);
    parts.push(corner);
    this.renderColumnHeaders(parts, this.frozenCols());
    this.renderRowHeaders(parts, this.frozenRows(), 0);
    this.renderCells(parts, 'corner', this.frozenRows(), this.frozenCols());
    this.corner.replaceChildren(...parts);
  }

  private renderTop(): void {
    const parts: Node[] = [this.corner];
    this.renderColumnHeaders(parts, this.renderedCols);
    this.renderCells(parts, 'top', this.frozenRows(), this.renderedCols);
    this.topBand.replaceChildren(...parts);
  }

  private renderLeft(): void {
    const parts: Node[] = [];
    this.renderRowHeaders(parts, this.renderedRows, frozenExtent(this.rowView));
    this.renderCells(parts, 'left', this.renderedRows, this.frozenCols());
    this.leftBand.replaceChildren(...parts);
  }

  private renderBody(): void {
    const parts: Node[] = [this.leftBand];
    this.renderCells(parts, 'body', this.renderedRows, this.renderedCols);
    this.body.replaceChildren(...parts);
  }

  /**
   * Numbers and dates that do not fit their cell: Excel shows ####, never a cut-off number. Fonts often differ from
   * the author's (Calibri is missing on many systems), so a number only slightly too wide is narrowed instead.
   */
  private fitNumbers(): void {
    const padding = 2 * Math.round(3 * this.zoom);
    // All measuring first, then all changes, so the browser lays out once.
    const measured = this.numbers.map((text): [HTMLElement, number, number] => [text, (text.parentElement?.clientWidth ?? 0) - padding, text.offsetWidth]);
    this.numbers = [];
    for (const [text, available, width] of measured) {
      if (width <= available + 0.5 || width === 0 || available <= 0) {
        continue;
      }
      const ratio = available / width;
      if (ratio >= MIN_NUMBER_SQUEEZE) {
        text.style.setProperty('transform', `scaleX(${ratio.toFixed(3)})`);
      } else {
        text.textContent = '#'.repeat(Math.max(1, Math.floor((text.textContent ?? '').length * ratio)));
      }
    }
  }

  private renderColumnHeaders(parts: Node[], cols: Span): void {
    for (let col = cols.first; col <= cols.last; col++) {
      const left = position(this.colView, col);
      const width = position(this.colView, col + 1) - left;
      if (width <= 0) {
        continue;
      }
      const header = this.box('dl-header dl-col-header', left, 0, width, this.rowView.header);
      header.textContent = columnName(col);
      header.dataset['index'] = String(col);
      parts.push(header);
    }
  }

  private renderRowHeaders(parts: Node[], rows: Span, offset: number): void {
    for (let row = rows.first; row <= rows.last; row++) {
      const top = position(this.rowView, row);
      const height = position(this.rowView, row + 1) - top;
      if (height <= 0) {
        continue;
      }
      const header = this.box('dl-header dl-row-header', 0, top - offset, this.colView.header, height);
      header.textContent = String(row + 1);
      header.dataset['index'] = String(row);
      parts.push(header);
    }
  }

  /** Gridlines, merged cells and cells of one part of the grid. */
  private renderCells(parts: Node[], quadrant: Quadrant, rows: Span, cols: Span): void {
    if (rows.last < rows.first || cols.last < cols.first) {
      return;
    }
    const offset = quadrant === 'left' || quadrant === 'body' ? frozenExtent(this.rowView) : 0;
    const colLimit: Span = quadrant === 'left' || quadrant === 'corner' ? this.frozenCols() : { first: this.colView.frozen, last: this.colView.count - 1 };
    this.renderGridlines(parts, rows, cols, offset);

    // A merged cell crossing a frozen pane's edge is drawn whole in each part it reaches, each clipped to its part.
    const merges = this.merges.query(rows.first, rows.last, cols.first, cols.last);
    for (const merge of merges) {
      const element = this.renderMerge(merge, offset);
      if (element) {
        parts.push(element);
      }
    }
    // Merged cells cannot overlap (the sheet drops those that do), so each window row has only a few.
    const mergesByRow = new Map<number, CellRange[]>();
    for (const merge of merges) {
      for (let row = Math.max(rows.first, merge.top); row <= Math.min(rows.last, merge.bottom); row++) {
        const list = mergesByRow.get(row);
        if (list) {
          list.push(merge);
        } else {
          mergesByRow.set(row, [merge]);
        }
      }
    }
    const covered = (row: number, col: number) => mergesByRow.get(row)?.some((merge) => rangeContains(merge, row, col)) ?? false;

    const cells = this.sheet.cells;
    for (let row = rows.first; row <= rows.last; row++) {
      if (this.rows.size(row) <= 0 || row >= cells.rowCount) {
        continue;
      }
      const rowStart = cells.rowStartIndex(row);
      const end = cells.rowEndIndex(row);
      let index = cells.lowerBound(row, cols.first);
      // Text from a cell left of this window may run into it, over formatted blanks in between.
      let before = index - 1;
      while (before >= rowStart && cells.kind(before) === CellKind.Blank && cells.col(before) > cols.first - OVERFLOW_MAX_COLS) {
        before--;
      }
      if (before >= rowStart && cells.col(before) >= colLimit.first && cells.col(before) >= cols.first - OVERFLOW_MAX_COLS) {
        const element = this.renderCell(before, row, offset, colLimit, covered, cols.first);
        if (element) {
          parts.push(element);
        }
      }
      for (; index < end; index++) {
        const col = cells.col(index);
        if (col > cols.last) {
          break;
        }
        const element = this.renderCell(index, row, offset, colLimit, covered);
        if (element) {
          parts.push(element);
        }
      }
    }
  }

  private renderGridlines(parts: Node[], rows: Span, cols: Span, offset: number): void {
    const left = position(this.colView, cols.first);
    const right = position(this.colView, cols.last + 1);
    const top = position(this.rowView, rows.first) - offset;
    const bottom = position(this.rowView, rows.last + 1) - offset;
    for (let row = rows.first; row <= rows.last; row++) {
      const y = position(this.rowView, row + 1) - offset - 1;
      if (this.rows.size(row) > 0) {
        parts.push(this.box('dl-gridline', left, y, right - left, 1));
      }
    }
    for (let col = cols.first; col <= cols.last; col++) {
      const x = position(this.colView, col + 1) - 1;
      if (this.cols.size(col) > 0) {
        parts.push(this.box('dl-gridline', x, top, 1, bottom - top));
      }
    }
  }

  /**
   * One cell. With `onlyIfReaching` set, the cell lies left of the window and is drawn only if its text runs into the
   * window at that column.
   */
  private renderCell(
    index: number,
    row: number,
    offset: number,
    colLimit: Span,
    covered: (row: number, col: number) => boolean,
    onlyIfReaching?: number,
  ): HTMLElement | undefined {
    const cells = this.sheet.cells;
    const col = cells.col(index);
    if (this.cols.size(col) <= 0 || covered(row, col)) {
      return undefined;
    }
    const format = this.formatOf(index);
    const display = this.display(index);
    const text = format.wrap ? display.text : display.text.replace(/[\r\n]+/g, ' ');
    const spill = text && display.overflows ? this.overflow(index, row, col, display.align, colLimit) : undefined;
    if (onlyIfReaching !== undefined && (!spill || position(this.colView, col + 1) + spill.right <= position(this.colView, onlyIfReaching))) {
      return undefined;
    }
    const rect = this.rangeRect({ top: row, left: col, bottom: row, right: col }, offset);
    const element = this.box(`dl-cell ${this.host.rules.classFor(cells.style(index))} dl-h-${display.align}`, rect.left, rect.top, rect.width, rect.height);
    if (display.color) {
      element.style.setProperty('color', display.color);
    }
    const content = this.content(text, display, this.links.at(row, col), format.wrap);
    if (spill && (spill.left || spill.right)) {
      element.classList.add('dl-spill');
      const wrapper = element.ownerDocument.createElement('span');
      wrapper.className = 'dl-spill-text';
      wrapper.style.setProperty('left', `${-spill.left}px`);
      wrapper.style.setProperty('width', `${rect.width + spill.left + spill.right}px`);
      wrapper.append(content);
      element.append(wrapper);
    } else {
      element.append(content);
    }
    return element;
  }

  private renderMerge(merge: CellRange, offset: number): HTMLElement | undefined {
    const rect = this.rangeRect(merge, offset);
    if (rect.width <= 1 || rect.height <= 1) {
      return undefined;
    }
    const cells = this.sheet.cells;
    const index = cells.find(merge.top, merge.left);
    const style = index < 0 ? 0 : cells.style(index);
    const element = this.box(`dl-cell dl-merge ${this.host.rules.classFor(style)}`, rect.left, rect.top, rect.width, rect.height);
    // The merged area's right and bottom edges show the borders of the cells along them.
    const right = this.formatAt(merge.top, merge.right).borders.right;
    const bottom = this.formatAt(merge.bottom, merge.left).borders.bottom;
    if (right) {
      element.style.setProperty('border-right', borderCss(right));
    }
    if (bottom) {
      element.style.setProperty('border-bottom', borderCss(bottom));
    }
    if (index >= 0) {
      const format = this.formatOf(index);
      const display = this.display(index);
      element.classList.add(`dl-h-${display.align}`);
      if (display.color) {
        element.style.setProperty('color', display.color);
      }
      const text = format.wrap ? display.text : display.text.replace(/[\r\n]+/g, ' ');
      element.append(this.content(text, display, this.links.at(merge.top, merge.left), format.wrap));
    }
    return element;
  }

  /** The cell's text (or rich text runs), in a link if it has one, as one node. */
  private content(text: string, display: CellDisplay, link: Hyperlink | undefined, wrap: boolean): Node {
    const document = this.element.ownerDocument;
    let node: Node = document.createTextNode(text);
    if (display.runs) {
      const span = document.createElement('span');
      span.className = 'dl-text';
      for (const run of display.runs) {
        const piece = document.createElement('span');
        piece.textContent = wrap ? run.text : run.text.replace(/[\r\n]+/g, ' ');
        if (run.font) {
          applyFont(piece.style, run.font);
        }
        span.append(piece);
      }
      node = span;
    }
    if (link) {
      const url = link.url ? externalUrl(link.url) : null;
      const anchor = document.createElement(url ? 'a' : 'span');
      anchor.className = 'dl-text dl-link';
      if (url) {
        anchor.setAttribute('href', url);
      } else if (link.location) {
        anchor.dataset['location'] = link.location;
      }
      anchor.title = link.tooltip ?? url ?? link.location ?? '';
      anchor.append(node);
      node = anchor;
    }
    if (display.numeric && text) {
      const span = document.createElement('span');
      span.className = 'dl-num';
      span.append(node);
      this.numbers.push(span);
      node = span;
    }
    return node;
  }

  /** How far (in pixels) a cell's text may run over empty neighbors on each side: up to a value or a merged cell. */
  private overflow(index: number, row: number, col: number, align: TextAlign, limit: Span): { left: number; right: number } {
    const cells = this.sheet.cells;
    let right = 0;
    let left = 0;
    if (align !== 'right') {
      let end = Math.min(limit.last + 1, col + 1 + OVERFLOW_MAX_COLS);
      for (let i = index + 1; i < cells.rowEndIndex(row); i++) {
        if (cells.kind(i) !== CellKind.Blank) {
          end = Math.min(end, cells.col(i));
          break;
        }
      }
      for (const merge of end > col + 1 ? this.merges.query(row, row, col + 1, end - 1) : []) {
        end = Math.min(end, merge.left);
      }
      right = position(this.colView, end) - position(this.colView, col + 1);
    }
    if (align !== 'left') {
      // `start` is the first column the text may cover.
      let start = Math.max(limit.first, col - OVERFLOW_MAX_COLS);
      for (let i = index - 1; i >= cells.rowStartIndex(row); i--) {
        if (cells.kind(i) !== CellKind.Blank) {
          start = Math.max(start, cells.col(i) + 1);
          break;
        }
      }
      for (const merge of start < col ? this.merges.query(row, row, start, col - 1) : []) {
        start = Math.max(start, merge.right + 1);
      }
      left = position(this.colView, col) - position(this.colView, start);
    }
    if (align === 'center') {
      const both = Math.min(left, right);
      return { left: both, right: both };
    }
    return { left, right };
  }

  private display(index: number): CellDisplay {
    return cellDisplay(this.host.workbook, this.sheet, index, this.formatOf(index));
  }

  private formatOf(index: number): CellFormat {
    return this.formats[this.sheet.cells.style(index)] ?? (this.formats[0] as CellFormat);
  }

  private formatAt(row: number, col: number): CellFormat {
    const index = this.sheet.cells.find(row, col);
    return index < 0 ? (this.formats[0] as CellFormat) : this.formatOf(index);
  }

  private quadrantOf(row: number, col: number): Quadrant {
    const frozenRow = row < this.rowView.frozen;
    const frozenCol = col < this.colView.frozen;
    return frozenRow ? (frozenCol ? 'corner' : 'top') : frozenCol ? 'left' : 'body';
  }

  /** Box of a cell range, including the gridline pixel before it. */
  private rangeRect(range: CellRange, offset: number): { left: number; top: number; width: number; height: number } {
    const left = position(this.colView, range.left) - 1;
    const top = position(this.rowView, range.top) - 1 - offset;
    return { left, top, width: position(this.colView, range.right + 1) - left, height: position(this.rowView, range.bottom + 1) - offset - top };
  }

  private box(className: string, left: number, top: number, width: number, height: number): HTMLDivElement {
    const node = this.element.ownerDocument.createElement('div');
    node.className = className;
    setBox(node, left, top, width, height);
    return node;
  }

  // Selection and input -----------------------------------------------------------------------------------------------

  private select(row: number, col: number, reveal = true): void {
    const range = this.selection.select(row, col);
    if (reveal) {
      const { width, height } = this.viewport();
      this.scrollTo(scrollToShow(this.rowView, range.top, this.element.scrollTop, height), scrollToShow(this.colView, range.left, this.element.scrollLeft, width));
      // Draw the new window now rather than on the scroll event, a frame later.
      this.update(false);
    }
    this.markSelection();
    this.host.selectionChanged();
  }

  private cellAtPoint(clientX: number, clientY: number): CellRef | undefined {
    const bounds = this.element.getBoundingClientRect();
    const x = clientX - bounds.left - this.element.clientLeft;
    const y = clientY - bounds.top - this.element.clientTop;
    const { width, height } = this.viewport();
    if (x < 0 || y < 0 || x >= width || y >= height) {
      return undefined;
    }
    const canvasX = x < frozenExtent(this.colView) ? x : x + this.element.scrollLeft;
    const canvasY = y < frozenExtent(this.rowView) ? y : y + this.element.scrollTop;
    if (canvasX < this.colView.header || canvasY < this.rowView.header) {
      return undefined;
    }
    return {
      row: this.rows.indexAt((canvasY - this.rowView.header) / this.zoom, this.rowView.count),
      col: this.cols.indexAt((canvasX - this.colView.header) / this.zoom, this.colView.count),
    };
  }

  private readonly onMouseDown = (event: MouseEvent) => {
    if (event.button !== 0) {
      return;
    }
    const cell = this.cellAtPoint(event.clientX, event.clientY);
    if (cell) {
      this.element.focus({ preventScroll: true });
      this.select(cell.row, cell.col, false);
    }
  };

  private readonly onClick = (event: MouseEvent) => {
    const target = (event.target as Element | null)?.closest?.('[data-location]');
    const location = target?.getAttribute('data-location');
    if (location) {
      event.preventDefault();
      this.host.goTo(location);
    }
  };

  private readonly onKeyDown = (event: KeyboardEvent) => {
    if (event.altKey) {
      return;
    }
    const next = this.selection.target(event.key, event.ctrlKey || event.metaKey, this.navigationBounds());
    if (next) {
      event.preventDefault();
      this.select(next.row, next.col);
    }
  };

  private navigationBounds(): NavigationBounds {
    const current = this.selection.range?.top ?? this.rowView.frozen;
    const pageHeight = (this.viewport().height - frozenExtent(this.rowView)) / this.zoom;
    return {
      rows: this.rows,
      cols: this.cols,
      rowCount: this.rowView.count,
      colCount: this.colView.count,
      lastRow: Math.max(0, this.sheet.used.rows - 1),
      lastCol: Math.max(0, this.sheet.used.cols - 1),
      page: Math.max(1, this.rows.countUntil(current, pageHeight, this.rowView.count) - 1),
      origin: { row: this.rowView.frozen, col: this.colView.frozen },
    };
  }
}

function setBox(node: HTMLElement, left: number | undefined, top: number | undefined, width: number | undefined, height: number | undefined): void {
  if (left !== undefined) {
    node.style.setProperty('left', `${left}px`);
  }
  if (top !== undefined) {
    node.style.setProperty('top', `${top}px`);
  }
  if (width !== undefined) {
    node.style.setProperty('width', `${Math.max(0, width)}px`);
  }
  if (height !== undefined) {
    node.style.setProperty('height', `${Math.max(0, height)}px`);
  }
}
