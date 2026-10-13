// Tables (a:tbl in a graphic frame) as HTML tables: grid column widths, minimum row heights, merged cells, and cell
// fills, borders and text from the cell itself or the table style regions it falls in (header row, banded rows, ...).
import { toCss, type Rgba } from './color.ts';
import { PPTX_LIMITS } from './limits.ts';
import { readFill, readLine, type LineStyle, type Paint } from './paint.ts';
import type { RenderContext } from './renderContext.ts';
import { px, readBodyProperties, renderParagraphs } from './text.ts';
import { cellTextDefaults, TextStyleResolver } from './textStyle.ts';
import { EMU_PER_PX } from './units.ts';
import { attr, bool, child, children, num, oneOf, parseXml } from './xml.ts';

const DEFAULT_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

// PowerPoint's default table style ("Medium Style 2 - Accent 1"). PowerPoint knows its built-in styles by id and
// files may refer to this one without defining it in tableStyles.xml.
const LT1_LINE = (width: number) => `<a:ln w="${width}"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>`;
const ACCENT_FILL = (tint?: number) => `<a:fill><a:solidFill><a:schemeClr val="accent1">${tint ? `<a:tint val="${tint}"/>` : ''}</a:schemeClr></a:solidFill></a:fill>`;
const BOLD_LIGHT_TEXT = '<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle>';
const DEFAULT_STYLE = `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="${DEFAULT_STYLE_ID}">
<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle>
<a:tcStyle><a:tcBdr>${['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map((side) => `<a:${side}>${LT1_LINE(12700)}</a:${side}>`).join('')}</a:tcBdr>${ACCENT_FILL(20000)}</a:tcStyle></a:wholeTbl>
<a:band1H><a:tcStyle><a:tcBdr/>${ACCENT_FILL(40000)}</a:tcStyle></a:band1H>
<a:band1V><a:tcStyle><a:tcBdr/>${ACCENT_FILL(40000)}</a:tcStyle></a:band1V>
<a:lastCol>${BOLD_LIGHT_TEXT}<a:tcStyle><a:tcBdr/>${ACCENT_FILL()}</a:tcStyle></a:lastCol>
<a:firstCol>${BOLD_LIGHT_TEXT}<a:tcStyle><a:tcBdr/>${ACCENT_FILL()}</a:tcStyle></a:firstCol>
<a:lastRow>${BOLD_LIGHT_TEXT}<a:tcStyle><a:tcBdr><a:top>${LT1_LINE(38100)}</a:top></a:tcBdr>${ACCENT_FILL()}</a:tcStyle></a:lastRow>
<a:firstRow>${BOLD_LIGHT_TEXT}<a:tcStyle><a:tcBdr><a:bottom>${LT1_LINE(38100)}</a:bottom></a:tcBdr>${ACCENT_FILL()}</a:tcStyle></a:firstRow>
</a:tblStyle>`;

let defaultStyle: Element | undefined;

/** Region names from least to most specific. */
type Region = 'wholeTbl' | 'band1H' | 'band2H' | 'band1V' | 'band2V' | 'lastCol' | 'firstCol' | 'lastRow' | 'firstRow' | 'neCell' | 'nwCell' | 'seCell' | 'swCell';
type Side = 'left' | 'right' | 'top' | 'bottom';

interface Flags {
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstCol: boolean;
  readonly lastCol: boolean;
  readonly bandRow: boolean;
  readonly bandCol: boolean;
}

interface Cell {
  readonly row: number;
  readonly column: number;
  readonly rows: number;
  readonly columns: number;
}

export async function renderTable(tbl: Element, context: RenderContext): Promise<HTMLElement> {
  const properties = child(tbl, 'tblPr');
  const flags: Flags = {
    firstRow: bool(properties, 'firstRow') ?? false,
    lastRow: bool(properties, 'lastRow') ?? false,
    firstCol: bool(properties, 'firstCol') ?? false,
    lastCol: bool(properties, 'lastCol') ?? false,
    bandRow: bool(properties, 'bandRow') ?? false,
    bandCol: bool(properties, 'bandCol') ?? false,
  };
  const style = await tableStyle(child(properties, 'tableStyleId')?.textContent?.trim(), context);
  const widths = children(child(tbl, 'tblGrid'), 'gridCol').map((column) => Math.max(0, (num(column, 'w') ?? 0) / EMU_PER_PX));
  const rows = children(tbl, 'tr');
  const rowCount = rows.length;
  const columnCount = widths.length;

  const table = document.createElement('table');
  table.className = 'pptx-table';
  table.style.setProperty('width', px(widths.reduce((sum, width) => sum + width, 0)));
  const tableFill = await readFill(oneOf(properties, ['noFill', 'solidFill', 'gradFill', 'pattFill']), context.colors, context);
  if (tableFill) {
    const color = paintColor(tableFill);
    if (color) {
      table.style.setProperty('background-color', color);
    }
  }
  const columns = document.createElement('colgroup');
  for (const width of widths) {
    const column = document.createElement('col');
    column.style.setProperty('width', px(width));
    columns.append(column);
  }
  table.append(columns);
  const body = document.createElement('tbody');
  table.append(body);

  let cells = 0;
  for (const [rowIndex, row] of rows.entries()) {
    const tr = document.createElement('tr');
    tr.style.setProperty('height', px((num(row, 'h') ?? 0) / EMU_PER_PX));
    let columnIndex = 0;
    for (const tc of children(row, 'tc')) {
      const position = columnIndex;
      columnIndex += 1;
      if (bool(tc, 'hMerge') || bool(tc, 'vMerge')) {
        continue;
      }
      if (++cells > PPTX_LIMITS.maxTableCells) {
        context.budget.truncated = true;
        body.append(tr);
        return table;
      }
      const span = Math.max(1, Math.min(columnCount - position, Math.trunc(num(tc, 'gridSpan') ?? 1)));
      const cell: Cell = { row: rowIndex, column: position, rows: Math.max(1, Math.min(rowCount - rowIndex, Math.trunc(num(tc, 'rowSpan') ?? 1))), columns: span };
      tr.append(await renderCell(tc, cell, regionsOf(cell, flags, rowCount, columnCount), style, rowCount, columnCount, context));
    }
    body.append(tr);
  }
  return table;
}

async function tableStyle(id: string | undefined, context: RenderContext): Promise<Element | undefined> {
  if (!id) {
    return undefined;
  }
  const styles = await context.deck.tableStyles();
  const defined = children(styles, 'tblStyle').find((candidate) => attr(candidate, 'styleId') === id);
  if (defined) {
    return defined;
  }
  if (id === DEFAULT_STYLE_ID) {
    defaultStyle ??= parseXml(DEFAULT_STYLE, 'default table style');
    return defaultStyle;
  }
  return undefined;
}

/** The style regions a cell falls in, least specific first. */
function regionsOf(cell: Cell, flags: Flags, rowCount: number, columnCount: number): Region[] {
  const regions: Region[] = ['wholeTbl'];
  const lastRowIndex = rowCount - 1;
  const lastColumnIndex = columnCount - 1;
  const top = cell.row === 0;
  const bottom = cell.row + cell.rows - 1 === lastRowIndex;
  const left = cell.column === 0;
  const right = cell.column + cell.columns - 1 === lastColumnIndex;
  if (flags.bandCol && !(flags.firstCol && left) && !(flags.lastCol && right)) {
    const band = cell.column - (flags.firstCol ? 1 : 0);
    regions.push(band % 2 === 0 ? 'band1V' : 'band2V');
  }
  if (flags.bandRow && !(flags.firstRow && top) && !(flags.lastRow && bottom)) {
    const band = cell.row - (flags.firstRow ? 1 : 0);
    regions.push(band % 2 === 0 ? 'band1H' : 'band2H');
  }
  if (flags.lastCol && right) {
    regions.push('lastCol');
  }
  if (flags.firstCol && left) {
    regions.push('firstCol');
  }
  if (flags.lastRow && bottom) {
    regions.push('lastRow');
  }
  if (flags.firstRow && top) {
    regions.push('firstRow');
  }
  if (flags.firstRow && flags.lastCol && top && right) {
    regions.push('neCell');
  }
  if (flags.firstRow && flags.firstCol && top && left) {
    regions.push('nwCell');
  }
  if (flags.lastRow && flags.lastCol && bottom && right) {
    regions.push('seCell');
  }
  if (flags.lastRow && flags.firstCol && bottom && left) {
    regions.push('swCell');
  }
  return regions;
}

async function renderCell(tc: Element, cell: Cell, regions: readonly Region[], style: Element | undefined, rowCount: number, columnCount: number, context: RenderContext): Promise<HTMLTableCellElement> {
  const td = document.createElement('td');
  if (cell.columns > 1) {
    td.colSpan = cell.columns;
  }
  if (cell.rows > 1) {
    td.rowSpan = cell.rows;
  }
  const properties = child(tc, 'tcPr');
  const css = td.style;
  const margin = (name: string, fallback: number) => px((num(properties, name) ?? fallback) / EMU_PER_PX);
  css.setProperty('padding', `${margin('marT', 45720)} ${margin('marR', 91440)} ${margin('marB', 45720)} ${margin('marL', 91440)}`);
  const anchor = attr(properties, 'anchor');
  css.setProperty('vertical-align', anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top');

  const specific = [...regions].reverse().map((region) => child(style, region));
  const fill = (await readFill(oneOf(properties, ['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill']), context.colors, context)) ?? (await styleCellFill(specific, context));
  const background = fill && paintColor(fill);
  if (background) {
    css.setProperty('background-color', background);
  }
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const line = await cellBorder(properties, side, specific, cell, rowCount, columnCount, context);
    if (line) {
      css.setProperty(`border-${side}`, line);
    }
  }

  const txBody = child(tc, 'txBody');
  if (txBody) {
    const textStyle = specific.map((region) => child(region, 'tcTxStyle')).find((item) => item !== undefined);
    const styles = new TextStyleResolver({
      base: [context.slide.master?.textStyles.other, context.deck.defaultTextStyle],
      shape: [child(txBody, 'lstStyle')],
      defaults: cellTextDefaults(textStyle),
    });
    td.append(...renderParagraphs(txBody, { body: readBodyProperties([child(txBody, 'bodyPr')]), styles, context }));
  }
  return td;
}

async function styleCellFill(regions: readonly (Element | undefined)[], context: RenderContext): Promise<Paint | undefined> {
  for (const region of regions) {
    const fill = child(child(region, 'tcStyle'), 'fill');
    const paint = await readFill(oneOf(fill, ['noFill', 'solidFill', 'gradFill', 'pattFill']), context.colors, context);
    if (paint) {
      return paint;
    }
  }
  return undefined;
}

const CELL_LINES: Readonly<Record<Side, string>> = { left: 'lnL', right: 'lnR', top: 'lnT', bottom: 'lnB' };

/** The CSS border for one side: the cell's own line, else the style's edge or inside line for that side. */
async function cellBorder(properties: Element | undefined, side: Side, regions: readonly (Element | undefined)[], cell: Cell, rowCount: number, columnCount: number, context: RenderContext): Promise<string | undefined> {
  const own = child(properties, CELL_LINES[side]);
  if (own) {
    return cssBorder(await readLine([{ ln: own, colors: context.colors }], context));
  }
  const atEdge: Record<Side, boolean> = {
    left: cell.column === 0,
    right: cell.column + cell.columns >= columnCount,
    top: cell.row === 0,
    bottom: cell.row + cell.rows >= rowCount,
  };
  for (const region of regions) {
    const borders = child(child(region, 'tcStyle'), 'tcBdr');
    if (!borders) {
      continue;
    }
    const name = region?.localName ?? '';
    const horizontal = side === 'top' || side === 'bottom';
    // Row regions own their top and bottom edges, column regions their left and right ones.
    const ownsEdge = name === 'wholeTbl' ? atEdge[side] : /Row|H$/.test(name) ? horizontal || atEdge[side] : /Col|V$/.test(name) ? !horizontal || atEdge[side] : true;
    const element = child(borders, ownsEdge ? side : horizontal ? 'insideH' : 'insideV');
    const ln = child(element, 'ln');
    if (ln) {
      return cssBorder(await readLine([{ ln, colors: context.colors }], context));
    }
  }
  return undefined;
}

function cssBorder(line: LineStyle | undefined): string | undefined {
  if (!line) {
    return undefined;
  }
  const color = paintColor(line.paint);
  if (!color || line.width <= 0) {
    return 'none';
  }
  if (line.compound && line.compound !== 'sng') {
    return `${px(Math.max(line.width, 3))} double ${color}`;
  }
  const dashed = line.dash && line.dash.length > 0 ? (line.dash[0] !== undefined && line.dash[0] <= line.width ? 'dotted' : 'dashed') : 'solid';
  return `${px(line.width)} ${dashed} ${color}`;
}

function paintColor(paint: Paint): string | undefined {
  const color: Rgba | undefined = paint.kind === 'solid' ? paint.color : paint.kind === 'gradient' ? paint.stops[0]?.color : paint.kind === 'pattern' ? paint.fg : undefined;
  return color && toCss(color);
}
