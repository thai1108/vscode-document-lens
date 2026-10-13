// Assembles .xlsx packages for tests from SpreadsheetML fragments: the caller writes the interesting XML (sheet
// bodies, styles), this adds the workbook, relationships and content types around it.
import { ooxmlPackage, type OoxmlPart } from './ooxml.ts';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT = 'application/vnd.openxmlformats-officedocument';

export interface XlsxSheet {
  name: string;
  state?: 'hidden' | 'veryHidden';
  /** Children of `<worksheet>` (or `<chartsheet>`): sheetViews, sheetFormatPr, cols, sheetData, mergeCells, ... */
  body: string;
  chart?: boolean;
  /** Relationships of the sheet part, e.g. hyperlink targets. */
  relationships?: { id: string; type: string; target: string; external?: boolean }[];
}

export interface XlsxSpec {
  sheets: XlsxSheet[];
  /** `<si>` items of the shared string table, as XML. */
  sharedStrings?: string[];
  /** Children of `<styleSheet>`. */
  styles?: string;
  /** A whole theme part. */
  theme?: string;
  date1904?: boolean;
  activeTab?: number;
}

export function xlsxPackage(spec: XlsxSpec): Buffer {
  const workbookRels: string[] = [];
  const parts: OoxmlPart[] = [];
  spec.sheets.forEach((sheet, i) => {
    const kind = sheet.chart ? 'chartsheet' : 'worksheet';
    const name = `xl/${kind}s/sheet${i + 1}.xml`;
    workbookRels.push(`<Relationship Id="rId${i + 1}" Type="${R}/${kind}" Target="${kind}s/sheet${i + 1}.xml"/>`);
    parts.push({ name, contentType: `${CONTENT}.spreadsheetml.${kind}+xml`, xml: `<${kind} xmlns="${MAIN}" xmlns:r="${R}">${sheet.body}</${kind}>` });
    if (sheet.relationships?.length) {
      const items = sheet.relationships.map(
        (relationship) =>
          `<Relationship Id="${relationship.id}" Type="${R}/${relationship.type}" Target="${escapeXml(relationship.target)}"${relationship.external ? ' TargetMode="External"' : ''}/>`,
      );
      parts.push({ name: `xl/${kind}s/_rels/sheet${i + 1}.xml.rels`, xml: `<Relationships xmlns="${RELS}">${items.join('')}</Relationships>` });
    }
  });
  let count = spec.sheets.length;
  const nextId = () => `rId${++count}`;
  if (spec.styles !== undefined) {
    workbookRels.push(`<Relationship Id="${nextId()}" Type="${R}/styles" Target="styles.xml"/>`);
    parts.push({ name: 'xl/styles.xml', contentType: `${CONTENT}.spreadsheetml.styles+xml`, xml: `<styleSheet xmlns="${MAIN}">${spec.styles}</styleSheet>` });
  }
  if (spec.sharedStrings) {
    workbookRels.push(`<Relationship Id="${nextId()}" Type="${R}/sharedStrings" Target="sharedStrings.xml"/>`);
    const count = spec.sharedStrings.length;
    parts.push({
      name: 'xl/sharedStrings.xml',
      contentType: `${CONTENT}.spreadsheetml.sharedStrings+xml`,
      xml: `<sst xmlns="${MAIN}" count="${count}" uniqueCount="${count}">${spec.sharedStrings.join('')}</sst>`,
    });
  }
  if (spec.theme) {
    workbookRels.push(`<Relationship Id="${nextId()}" Type="${R}/theme" Target="theme/theme1.xml"/>`);
    parts.push({ name: 'xl/theme/theme1.xml', contentType: `${CONTENT}.theme+xml`, xml: spec.theme });
  }
  parts.push({ name: 'xl/_rels/workbook.xml.rels', xml: `<Relationships xmlns="${RELS}">${workbookRels.join('')}</Relationships>` });
  const sheets = spec.sheets
    .map((sheet, i) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${i + 1}"${sheet.state ? ` state="${sheet.state}"` : ''} r:id="rId${i + 1}"/>`)
    .join('');
  const workbook = `<workbook xmlns="${MAIN}" xmlns:r="${R}"><workbookPr${spec.date1904 ? ' date1904="1"' : ''}/><bookViews><workbookView activeTab="${spec.activeTab ?? 0}"/></bookViews><sheets>${sheets}</sheets></workbook>`;
  return ooxmlPackage({ name: 'xl/workbook.xml', contentType: `${CONTENT}.spreadsheetml.sheet.main+xml`, xml: workbook }, parts);
}

/** Value of the cell at (row, col) in largeSheetXlsx: readable back from the screen, e.g. 10000049 for AX100000. */
export function largeSheetValue(row: number, col: number): number {
  return (row + 1) * 100 + col;
}

/**
 * A plain sheet of `rows` × `cols` numbers with a frozen first row and column, written the compact way some
 * generators do (no `r` on cells), for virtual scrolling tests.
 */
export function largeSheetXlsx(rows: number, cols: number): Buffer {
  const lines: string[] = [];
  for (let row = 0; row < rows; row++) {
    let line = `<row r="${row + 1}">`;
    for (let col = 0; col < cols; col++) {
      line += `<c><v>${largeSheetValue(row, col)}</v></c>`;
    }
    lines.push(`${line}</row>`);
  }
  const view = '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>';
  return xlsxPackage({ sheets: [{ name: 'Large', body: `${view}<sheetData>${lines.join('')}</sheetData>` }] });
}

export function escapeXml(text: string): string {
  return text.replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[ch] ?? ch);
}

/** A `<c>` element: numbers as is, strings as inline strings unless `type` says otherwise. */
export function cell(ref: string, value: number | string | boolean | null, options: { style?: number; type?: string; formula?: string } = {}): string {
  const style = options.style ? ` s="${options.style}"` : '';
  const formula = options.formula ? `<f>${escapeXml(options.formula)}</f>` : '';
  if (value === null) {
    return `<c r="${ref}"${style}>${formula}</c>`;
  }
  if (typeof value === 'boolean') {
    return `<c r="${ref}"${style} t="b">${formula}<v>${value ? 1 : 0}</v></c>`;
  }
  if (typeof value === 'number') {
    return `<c r="${ref}"${style}${options.type ? ` t="${options.type}"` : ''}>${formula}<v>${value}</v></c>`;
  }
  if (options.type) {
    return `<c r="${ref}"${style} t="${options.type}">${formula}<v>${escapeXml(value)}</v></c>`;
  }
  return `<c r="${ref}"${style} t="inlineStr">${formula}<is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}
