// A worksheet part (xl/worksheets/sheetN.xml), read as a stream: cells go straight into a SheetBuilder, so a sheet
// with millions of cells never exists as one string or one object per cell.
import { externalUrl } from '../../../protocol.ts';
import { SHEET_LIMITS } from '../limits.ts';
import { parseCellRef, parseRange, type CellRef } from '../model/address.ts';
import { CellKind, NO_FORMULA } from '../model/cells.ts';
import type { ColorContext } from '../model/colors.ts';
import { SheetBuilder } from '../model/sheet.ts';
import { baseColumnWidthPx, columnWidthPx, defaultRowHeightPt, maxDigitWidth, pointsToPx } from '../model/sizes.ts';
import type { CellFormat, Font, Hyperlink, Sheet, SheetKind } from '../model/types.ts';
import type { Relationship } from './relationships.ts';
import { RichTextReader } from './strings.ts';
import { boolAttr, intAttr } from './styles.ts';
import { XmlScanner, type XmlTag } from './xml.ts';

export interface WorksheetContext {
  name: string;
  kind: SheetKind;
  /** The part being read, e.g. `xl/worksheets/sheet2.xml`. */
  part?: string;
  formats: readonly CellFormat[];
  sharedStringCount: number;
  date1904: boolean;
  defaultFont: Font;
  colors: ColorContext;
  /** The sheet part's relationships by id (hyperlink targets, drawings). */
  relationships: ReadonlyMap<string, Relationship>;
  /** Cells kept before the rest of the sheet is cut off (SHEET_LIMITS.maxCells unless a test says otherwise). */
  maxCells?: number;
}

type TextTarget = 'none' | 'value' | 'formula';

const MS_PER_DAY = 86_400_000;

export async function readWorksheet(chunks: AsyncIterable<string>, context: WorksheetContext, signal?: AbortSignal): Promise<Sheet> {
  const { formats } = context;
  const sheet = new SheetBuilder(context.name, context.kind, formats, context.maxCells === undefined ? {} : { maxCells: context.maxCells });
  const digitWidth = maxDigitWidth(context.defaultFont);
  const rowSizes = new Map<number, number>();
  const colSizes = new Map<number, number>();

  let defaultRowHeight: number | undefined;
  let rowsHiddenByDefault = false;
  let defaultColWidth: number | undefined;
  let baseColWidth = 8;
  let showGridLines = true;
  let frozen = { rows: 0, cols: 0 };
  let viewTopLeft: CellRef | undefined;
  let paneTopLeft: CellRef | undefined;
  let activePane: string | undefined;
  let activeCell: CellRef | undefined;

  // Position in the document.
  let section = '';
  let sheetViews = 0;
  let row = -1;
  let nextCol = 0;
  // The cell being read.
  let cellRow = 0;
  let cellCol = 0;
  let cellStyle = 0;
  let cellType = 'n';
  let value = '';
  let formula = '';
  let formulaType: string | undefined;
  let formulaShared: number | undefined;
  let inlineString = false;
  let textTarget: TextTarget = 'none';

  const scanner = new XmlScanner({
    open(tag) {
      if (inlineString) {
        inline.open(tag);
        return;
      }
      switch (tag.name) {
        case 'sheetData':
          section = 'sheetData';
          return;
        case 'row':
          if (section === 'sheetData') {
            openRow(tag);
          }
          return;
        case 'c':
          if (section === 'sheetData') {
            openCell(tag);
          }
          return;
        case 'v':
          if (section === 'sheetData' && !tag.selfClosing) {
            setText('value');
          }
          return;
        case 'f':
          if (section === 'sheetData') {
            formulaType = tag.attr('t');
            formulaShared = intAttr(tag, 'si');
            if (!tag.selfClosing) {
              setText('formula');
            }
          }
          return;
        case 'is':
          if (section === 'sheetData' && !tag.selfClosing) {
            inlineString = true;
            inline.reset();
          }
          return;
        case 'dimension': {
          const range = parseRange(tag.attr('ref') ?? '');
          if (range) {
            sheet.expectCells((range.bottom - range.top + 1) * (range.right - range.left + 1));
          }
          return;
        }
        case 'sheetFormatPr':
          readFormat(tag);
          return;
        case 'cols':
        case 'mergeCells':
        case 'hyperlinks':
          section = tag.name;
          return;
        case 'col':
          if (section === 'cols') {
            readColumn(tag);
          }
          return;
        case 'sheetView':
          sheetViews++;
          if (sheetViews === 1) {
            section = 'sheetView';
            showGridLines = boolAttr(tag, 'showGridLines') ?? true;
            viewTopLeft = cellAttr(tag, 'topLeftCell');
          }
          return;
        case 'pane':
          if (section === 'sheetView') {
            readPane(tag);
          }
          return;
        case 'selection':
          if (section === 'sheetView' && (!activeCell || tag.attr('pane') === activePane)) {
            activeCell = cellAttr(tag, 'activeCell') ?? activeCell;
          }
          return;
        case 'mergeCell':
          if (section === 'mergeCells') {
            const range = parseRange(tag.attr('ref') ?? '');
            if (range) {
              sheet.addMerge(range);
            }
          }
          return;
        case 'hyperlink':
          if (section === 'hyperlinks') {
            readHyperlink(tag);
          }
          return;
        case 'drawing':
          sheet.noteDrawings();
          return;
      }
    },
    close(name) {
      if (inlineString) {
        if (name === 'is') {
          inlineString = false;
          return;
        }
        if (name !== 'c') {
          inline.close(name);
          return;
        }
        // A cell ended inside its inline string: malformed, but keep reading cells.
        inlineString = false;
      }
      switch (name) {
        case 'c':
          if (section === 'sheetData') {
            commitCell();
          }
          return;
        case 'v':
        case 'f':
          setText('none');
          return;
        case 'sheetData':
        case 'cols':
        case 'sheetView':
        case 'mergeCells':
        case 'hyperlinks':
          section = '';
          return;
      }
    },
    text(text) {
      if (textTarget === 'value') {
        value += text;
      } else if (textTarget === 'formula') {
        formula += text;
      } else if (inlineString) {
        inline.onText(text);
      }
    },
  });
  const inline = new RichTextReader(scanner, context.colors);

  function setText(target: TextTarget): void {
    textTarget = target;
    scanner.wantText = target !== 'none';
  }

  function openRow(tag: XmlTag): void {
    const r = intAttr(tag, 'r');
    row = r !== undefined && r >= 1 && r <= SHEET_LIMITS.maxRows ? r - 1 : Math.min(row + 1, SHEET_LIMITS.maxRows - 1);
    nextCol = 0;
    const hidden = boolAttr(tag, 'hidden');
    const height = Number(tag.attr('ht'));
    if (rowSizes.size < SHEET_LIMITS.maxRowEntries) {
      if (hidden) {
        rowSizes.set(row, 0);
      } else if (Number.isFinite(height) && height >= 0 && height <= 409.5) {
        rowSizes.set(row, pointsToPx(height));
      } else if (rowsHiddenByDefault && hidden === false) {
        rowSizes.set(row, pointsToPx(defaultRowHeight ?? defaultRowHeightPt(context.defaultFont)));
      }
    }
  }

  function openCell(tag: XmlTag): void {
    const reference = tag.attr('r');
    const parsed = reference ? parseCellRef(reference) : undefined;
    cellRow = parsed ? parsed.row : Math.max(0, row);
    cellCol = parsed ? parsed.col : Math.min(nextCol, SHEET_LIMITS.maxCols - 1);
    nextCol = cellCol + 1;
    const style = intAttr(tag, 's') ?? 0;
    cellStyle = style < formats.length ? style : 0;
    cellType = tag.attr('t') ?? 'n';
    value = '';
    formula = '';
    formulaType = undefined;
    formulaShared = undefined;
  }

  function commitCell(): void {
    setText('none');
    if (sheet.cutOff) {
      return;
    }
    let kind: CellKind = CellKind.Blank;
    let content = 0;
    switch (cellType) {
      case 's': {
        const index = Number(value);
        if (value !== '' && Number.isInteger(index) && index >= 0 && index < context.sharedStringCount) {
          kind = CellKind.SharedString;
          content = index;
        }
        break;
      }
      case 'inlineStr': {
        const { text, runs } = inline.result();
        kind = CellKind.SheetString;
        content = sheet.addString(text, runs);
        break;
      }
      case 'str':
        kind = CellKind.SheetString;
        content = sheet.addString(value);
        break;
      case 'b':
        if (value !== '') {
          kind = CellKind.Boolean;
          content = value === '1' || value === 'true' ? 1 : 0;
        }
        break;
      case 'e':
        kind = CellKind.Error;
        content = sheet.addString(value);
        break;
      case 'd': {
        const serial = serialDate(value, context.date1904);
        if (serial !== undefined) {
          kind = CellKind.Number;
          content = serial;
        } else if (value !== '') {
          kind = CellKind.SheetString;
          content = sheet.addString(value);
        }
        break;
      }
      default:
        if (value !== '') {
          const number = Number(value);
          if (Number.isFinite(number)) {
            kind = CellKind.Number;
            content = number;
          } else {
            kind = CellKind.SheetString;
            content = sheet.addString(value);
          }
        }
    }
    sheet.add(cellRow, cellCol, kind, content, cellStyle, readFormula());
  }

  /** The cell's formula reference; the first cell of a shared formula holds its text for the others. */
  function readFormula(): number {
    if (formulaType === 'shared' && formulaShared !== undefined) {
      if (!formula) {
        return sheet.sharedFormula(formulaShared);
      }
      sheet.defineSharedFormula(formulaShared, formula, cellRow, cellCol);
    }
    return formula ? sheet.addFormula(formula) : NO_FORMULA;
  }

  function readFormat(tag: XmlTag): void {
    const height = Number(tag.attr('defaultRowHeight'));
    if (Number.isFinite(height) && height > 0 && height <= 409.5) {
      defaultRowHeight = height;
    }
    rowsHiddenByDefault = boolAttr(tag, 'zeroHeight') ?? false;
    const width = Number(tag.attr('defaultColWidth'));
    if (Number.isFinite(width) && width >= 0 && width <= 255) {
      defaultColWidth = width;
    }
    const base = intAttr(tag, 'baseColWidth');
    if (base !== undefined && base <= 255) {
      baseColWidth = base;
    }
  }

  function readColumn(tag: XmlTag): void {
    const min = intAttr(tag, 'min');
    const max = intAttr(tag, 'max');
    if (min === undefined || max === undefined || min < 1 || max < min) {
      return;
    }
    const width = Number(tag.attr('width'));
    const hidden = boolAttr(tag, 'hidden') ?? false;
    if (!hidden && !(Number.isFinite(width) && width >= 0 && width <= 255)) {
      return;
    }
    const size = hidden ? 0 : columnWidthPx(width, digitWidth);
    for (let col = min - 1; col < Math.min(max, SHEET_LIMITS.maxCols); col++) {
      colSizes.set(col, size);
    }
  }

  function readPane(tag: XmlTag): void {
    const state = tag.attr('state');
    activePane = tag.attr('activePane');
    if (state === 'frozen' || state === 'frozenSplit') {
      const cols = Number(tag.attr('xSplit') ?? 0);
      const rows = Number(tag.attr('ySplit') ?? 0);
      frozen = {
        rows: Number.isInteger(rows) && rows > 0 ? Math.min(rows, SHEET_LIMITS.maxRows) : 0,
        cols: Number.isInteger(cols) && cols > 0 ? Math.min(cols, SHEET_LIMITS.maxCols) : 0,
      };
      paneTopLeft = cellAttr(tag, 'topLeftCell');
    }
  }

  function readHyperlink(tag: XmlTag): void {
    const range = parseRange(tag.attr('ref') ?? '');
    if (!range) {
      return;
    }
    const id = tag.attr('id');
    const target = id ? context.relationships.get(id) : undefined;
    const location = tag.attr('location');
    const tooltip = tag.attr('tooltip');
    const link: Hyperlink = { range };
    const url = target?.external ? externalUrl(location ? `${target.target}#${location}` : target.target) : null;
    if (url) {
      link.url = url;
    } else if (location && !target) {
      link.location = location;
    } else {
      return;
    }
    if (tooltip) {
      link.tooltip = tooltip.slice(0, 1024);
    }
    sheet.addLink(link);
  }

  for await (const chunk of chunks) {
    signal?.throwIfAborted();
    scanner.write(chunk);
  }
  scanner.end();

  const frozenPanes = frozen.rows > 0 || frozen.cols > 0;
  const topLeft = frozenPanes ? paneTopLeft : viewTopLeft;
  return sheet.build({
    rows: { defaultSize: rowsHiddenByDefault ? 0 : pointsToPx(defaultRowHeight ?? defaultRowHeightPt(context.defaultFont)), sizes: rowSizes },
    cols: { defaultSize: defaultColWidth !== undefined ? columnWidthPx(defaultColWidth, digitWidth) : baseColumnWidthPx(baseColWidth, digitWidth), sizes: colSizes },
    frozen,
    showGridLines,
    ...(topLeft ? { topLeft } : {}),
    ...(activeCell ? { activeCell } : {}),
    ...(context.part ? { sourcePart: context.part } : {}),
  });
}

function cellAttr(tag: XmlTag, name: string): CellRef | undefined {
  const text = tag.attr(name);
  return text ? parseCellRef(text.split(/\s/)[0] ?? '') : undefined;
}

/** An ISO 8601 date (`t="d"` cells) as a serial date number. */
function serialDate(text: string, date1904: boolean): number | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(text.trim());
  if (!match) {
    return undefined;
  }
  const [, y, m, d, hh = '0', mm = '0', ss = '0'] = match;
  const ms = Date.UTC(Number(y), Number(m) - 1, Number(d), Number(hh), Number(mm), 0) + Number(ss) * 1000;
  if (!Number.isFinite(ms)) {
    return undefined;
  }
  // Serial 25569 is 1970-01-01 in the 1900 system (which counts a 29 Feb 1900 that never was); 24107 in the 1904 one.
  const serial = ms / MS_PER_DAY + (date1904 ? 24_107 : 25_569);
  return serial >= 0 ? serial : undefined;
}
