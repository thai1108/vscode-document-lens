// What a cell shows: its formatted text, color and alignment, and what the cell bar shows for it.
import { CellKind } from './cells.ts';
import { formatNumber, formatText, generalNumber, isDateFormat, isoDate } from './numberFormat.ts';
import type { CellFormat, RichRun, Sheet, Workbook } from './types.ts';

export type TextAlign = 'left' | 'center' | 'right';

export interface CellDisplay {
  text: string;
  /** Color from the number format (e.g. `[Red]`), which wins over the font color. */
  color?: string;
  align: TextAlign;
  /** Text that may run over empty neighbors (numbers and dates never do in Excel). */
  overflows: boolean;
  /** A number or date: shown as #### rather than cut off when it does not fit. */
  numeric: boolean;
  runs?: readonly RichRun[];
}

export function cellDisplay(workbook: Workbook, sheet: Sheet, index: number, format: CellFormat): CellDisplay {
  const cells = sheet.cells;
  const kind = cells.kind(index);
  const value = cells.value(index);
  switch (kind) {
    case CellKind.Number: {
      const formatted = formatNumber(value, format.numberFormat, workbook.date1904);
      return { ...formatted, align: align(format, 'right'), overflows: false, numeric: true };
    }
    case CellKind.SharedString:
    case CellKind.SheetString: {
      const table = kind === CellKind.SharedString ? workbook.sharedStrings : sheet.strings;
      const text = table.texts[value] ?? '';
      const runs = table.rich.get(value);
      const formatted = format.numberFormat === 'General' ? { text } : formatText(text, format.numberFormat);
      return { ...formatted, align: align(format, 'left'), overflows: !format.wrap, numeric: false, ...(runs && formatted.text === text ? { runs } : {}) };
    }
    case CellKind.Boolean:
      return { text: value ? 'TRUE' : 'FALSE', align: align(format, 'center'), overflows: false, numeric: false };
    case CellKind.Error:
      return { text: sheet.strings.texts[value] ?? '#VALUE!', align: align(format, 'center'), overflows: false, numeric: false };
    default:
      return { text: '', align: align(format, 'left'), overflows: false, numeric: false };
  }
}

/** One cell of a sheet as a plain object, for the cell bar and for anything else working on single cells. */
export interface CellView {
  row: number;
  col: number;
  kind: CellKind;
  /** The value itself; undefined for a blank cell. */
  value: number | string | boolean | undefined;
  /** As the grid shows it, number format applied. */
  text: string;
  /** As the cell bar shows it: `=` and the formula, or the value unformatted. */
  raw: string;
  /** Without the leading `=`. */
  formula?: string;
  /** Index into Workbook.formats. */
  style: number;
}

/** The cell at (row, col), or undefined if the sheet stores nothing there. */
export function readCell(workbook: Workbook, sheet: Sheet, row: number, col: number): CellView | undefined {
  const cells = sheet.cells;
  const index = cells.find(row, col);
  if (index < 0) {
    return undefined;
  }
  const style = cells.style(index);
  const format = workbook.formats[style] ?? workbook.formats[0];
  if (!format) {
    return undefined;
  }
  const kind = cells.kind(index);
  const number = cells.value(index);
  let value: CellView['value'];
  switch (kind) {
    case CellKind.Number:
      value = number;
      break;
    case CellKind.Boolean:
      value = number !== 0;
      break;
    case CellKind.SharedString:
      value = workbook.sharedStrings.texts[number] ?? '';
      break;
    case CellKind.SheetString:
    case CellKind.Error:
      value = sheet.strings.texts[number] ?? '';
      break;
  }
  const formula = sheet.formulaText(index, row, col);
  return {
    row,
    col,
    kind,
    value,
    text: cellDisplay(workbook, sheet, index, format).text,
    raw: cellRawText(workbook, sheet, index, row, col, format),
    ...(formula === undefined ? {} : { formula }),
    style,
  };
}

/** The cell bar: the formula if there is one, else the value unformatted (dates as ISO dates). */
export function cellRawText(workbook: Workbook, sheet: Sheet, index: number, row: number, col: number, format: CellFormat): string {
  const formula = sheet.formulaText(index, row, col);
  if (formula !== undefined) {
    return `=${formula}`;
  }
  const cells = sheet.cells;
  const value = cells.value(index);
  switch (cells.kind(index)) {
    case CellKind.Number:
      return (isDateFormat(format.numberFormat) && isoDate(value, workbook.date1904)) || generalRaw(value);
    case CellKind.SharedString:
      return workbook.sharedStrings.texts[value] ?? '';
    case CellKind.SheetString:
    case CellKind.Error:
      return sheet.strings.texts[value] ?? '';
    case CellKind.Boolean:
      return value ? 'TRUE' : 'FALSE';
    default:
      return '';
  }
}

/** Full precision up to the 15 significant digits Excel keeps. */
function generalRaw(value: number): string {
  const text = String(Number(value.toPrecision(15)));
  return text.includes('e') ? generalNumber(value) : text;
}

function align(format: CellFormat, general: TextAlign): TextAlign {
  switch (format.horizontal) {
    case 'left':
    case 'fill':
    case 'justify':
      return 'left';
    case 'center':
    case 'centerContinuous':
    case 'distributed':
      return 'center';
    case 'right':
      return 'right';
    default:
      return general;
  }
}
