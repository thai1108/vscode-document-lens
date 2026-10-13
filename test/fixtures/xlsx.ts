// sample.xlsx: a styled workbook exercising what the spreadsheet viewer reads. Every value and format here is
// asserted somewhere in test/webview/spreadsheet or test/e2e/spreadsheet.spec.ts.
import type { Sample } from '../support/samples.ts';
import { cell, escapeXml, xlsxPackage } from '../support/xlsxWriter.ts';

export const GITHUB_URL = 'https://github.com/thai1108/vscode-document-lens';

/** Shared strings, by index. */
const STRINGS = [
  'Item',
  'Amount',
  'Date',
  'Share',
  'Note',
  'Coffee beans',
  'Green tea 緑茶',
  'Rice crackers',
  'Shipping',
  'Discount',
  'Total',
  // 11: rich text, three runs.
  '<r><t xml:space="preserve">Plain, </t></r><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/></rPr><t>bold red</t></r><r><rPr><i/><sz val="11"/><rFont val="Calibri"/></rPr><t xml:space="preserve">, italic</t></r>',
  // 12: with a phonetic guide, which is not part of the text.
  '<t>東京本社</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/>',
  'Quarterly report 四半期報告',
  'Document Lens on GitHub',
  'This note is long enough to wrap onto a second line inside its cell. 長いメモは折り返して表示されます。',
  'This title runs across the empty cells to its right',
  'Go to the Data sheet',
  'secret',
  'ID',
  'Name',
  'Qty',
  'Price',
  'Flag',
  'Code',
];
const s = (text: string) => STRINGS.indexOf(text);
const RICH_STRING = 11;
const PHONETIC_STRING = 12;

const STYLES = [
  '<numFmts count="3">',
  '<numFmt numFmtId="164" formatCode="#,##0.00;[Red]-#,##0.00"/>',
  '<numFmt numFmtId="165" formatCode="yyyy-mm-dd"/>',
  '<numFmt numFmtId="166" formatCode="[$-411]ggge&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/>',
  '</numFmts>',
  '<fonts count="6">',
  '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
  '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',
  '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>',
  '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font>',
  '<font><i/><sz val="11"/><color theme="5"/><name val="Calibri"/></font>',
  '<font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><charset val="128"/></font>',
  '</fonts>',
  '<fills count="5">',
  '<fill><patternFill patternType="none"/></fill>',
  '<fill><patternFill patternType="gray125"/></fill>',
  '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>',
  '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>',
  '<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.59999389629810485"/><bgColor indexed="64"/></patternFill></fill>',
  '</fills>',
  '<borders count="4">',
  '<border><left/><right/><top/><bottom/><diagonal/></border>',
  '<border><left/><right/><top/><bottom style="medium"><color auto="1"/></bottom><diagonal/></border>',
  '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>',
  '<border><left/><right/><top style="thin"><color rgb="FF000000"/></top><bottom style="double"><color rgb="FF000000"/></bottom><diagonal/></border>',
  '</borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="15">',
  /* 0 default */ '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
  /* 1 title */ '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>',
  /* 2 header */ '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>',
  /* 3 amount */ '<xf numFmtId="164" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1"/>',
  /* 4 date */ '<xf numFmtId="165" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1"/>',
  /* 5 percent */ '<xf numFmtId="10" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1"/>',
  /* 6 note */ '<xf numFmtId="0" fontId="5" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>',
  /* 7 total */ '<xf numFmtId="164" fontId="2" fillId="0" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>',
  /* 8 link */ '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
  /* 9 era */ '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
  /* 10 accent */ '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
  /* 11 data header */ '<xf numFmtId="0" fontId="2" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>',
  /* 12 text */ '<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>',
  /* 13 integer */ '<xf numFmtId="3" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1"/>',
  /* 14 boolean */ '<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>',
  '</cellXfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
].join('');

/** Office 2007 colors, so theme lookups are told apart from the viewer's built-in defaults (accent2 is C0504D). */
const THEME =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office">' +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3>' +
  '<a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>' +
  '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme></a:themeElements></a:theme>';

const AMOUNTS = [1234.5, 980.25, 312, 45.8, -120];
const TOTAL = AMOUNTS.reduce((sum, amount) => sum + amount, 0);
/** 2024-01-15, 2024-02-15, ... as serial dates (1900 system). */
const DATES = [45306, 45337, 45366, 45397, 45427];
const ITEMS = ['Coffee beans', 'Green tea 緑茶', 'Rice crackers', 'Shipping', 'Discount'];

function summarySheet(): string {
  const rows: string[] = [];
  rows.push(`<row r="1" ht="30" customHeight="1">${cell('A1', s('Quarterly report 四半期報告'), { style: 1, type: 's' })}<c r="B1" s="1"/><c r="C1" s="1"/><c r="D1" s="1"/></row>`);
  rows.push(`<row r="2">${['Item', 'Amount', 'Date', 'Share', 'Note'].map((text, i) => cell(`${'ABCDE'[i]}2`, s(text), { style: 2, type: 's' })).join('')}</row>`);
  ITEMS.forEach((item, i) => {
    const r = i + 3;
    const share = `<c r="D${r}" s="5">${i === 0 ? `<f t="shared" ref="D3:D7" si="0">B3/$B$8</f>` : '<f t="shared" si="0"/>'}<v>${(AMOUNTS[i] ?? 0) / TOTAL}</v></c>`;
    const notes: Record<number, number> = { 0: PHONETIC_STRING, 1: s('This note is long enough to wrap onto a second line inside its cell. 長いメモは折り返して表示されます。') };
    const note = notes[i] === undefined ? '' : cell(`E${r}`, notes[i] ?? 0, { style: 6, type: 's' });
    const secret = i === 0 ? cell('F3', s('secret'), { type: 's' }) : '';
    rows.push(
      `<row r="${r}"${i === 1 ? ' ht="45" customHeight="1"' : ''}>${cell(`A${r}`, s(item), { style: 12, type: 's' })}${cell(`B${r}`, AMOUNTS[i] ?? 0, { style: 3 })}${cell(`C${r}`, DATES[i] ?? 0, { style: 4 })}${share}${note}${secret}</row>`,
    );
  });
  rows.push(`<row r="8">${cell('A8', s('Total'), { style: 7, type: 's' })}${cell('B8', TOTAL, { style: 7, formula: 'SUM(B3:B7)' })}</row>`);
  rows.push(`<row r="9">${cell('E9', s('Document Lens on GitHub'), { style: 8, type: 's' })}</row>`);
  rows.push(`<row r="10">${cell('A10', RICH_STRING, { type: 's' })}</row>`);
  rows.push(`<row r="11" hidden="1">${cell('A11', 'hidden row')}</row>`);
  rows.push(`<row r="12">${cell('B12', s('This title runs across the empty cells to its right'), { type: 's' })}</row>`);
  rows.push(`<row r="13">${cell('A13', 45306, { style: 9 })}${cell('B13', true, { style: 14 })}<c r="C13" t="e"><v>#DIV/0!</v></c></row>`);
  rows.push(`<row r="14">${cell('A14', s('Go to the Data sheet'), { style: 8, type: 's' })}</row>`);
  rows.push(`<row r="15">${cell('A15', 'Accent color', { style: 10 })}</row>`);
  return [
    '<dimension ref="A1:F15"/>',
    '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>',
    '<selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="B3" sqref="B3"/></sheetView></sheetViews>',
    '<sheetFormatPr defaultRowHeight="15"/>',
    '<cols><col min="1" max="1" width="24.7109375" customWidth="1"/><col min="2" max="2" width="14.7109375" customWidth="1"/>',
    '<col min="3" max="3" width="12.7109375" customWidth="1"/><col min="4" max="4" width="10.7109375" customWidth="1"/>',
    '<col min="5" max="5" width="34.7109375" customWidth="1"/><col min="6" max="6" width="9.140625" hidden="1"/></cols>',
    `<sheetData>${rows.join('')}</sheetData>`,
    '<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>',
    '<hyperlinks><hyperlink ref="E9" r:id="rId1"/><hyperlink ref="A14" location="Data!A1" display="Go to the Data sheet"/></hyperlinks>',
  ].join('');
}

export const DATA_ROWS = 300;

function dataSheet(): string {
  const rows: string[] = [];
  rows.push(`<row r="1">${['ID', 'Name', 'Qty', 'Price', 'Date', 'Flag', 'Code'].map((text, i) => cell(`${'ABCDEFG'[i]}1`, s(text), { style: 11, type: 's' })).join('')}</row>`);
  for (let i = 1; i <= DATA_ROWS; i++) {
    const r = i + 1;
    const code = i <= 3 ? `<c r="G${r}" s="12" t="str"><f>A${r}&amp;"-"&amp;B${r}</f><v>${i}-Product ${i}</v></c>` : '';
    rows.push(
      `<row r="${r}">${cell(`A${r}`, i, { style: 13 })}${cell(`B${r}`, `Product ${i}`, { style: 12 })}${cell(`C${r}`, (i * 7) % 50, { style: 13 })}` +
        `${cell(`D${r}`, i * 1.25, { style: 3 })}${cell(`E${r}`, 45292 + i, { style: 4 })}${cell(`F${r}`, i % 2 === 0, { style: 14 })}${code}</row>`,
    );
  }
  return [
    `<dimension ref="A1:G${DATA_ROWS + 1}"/>`,
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>',
    '<sheetFormatPr defaultRowHeight="15"/>',
    '<cols><col min="2" max="2" width="16.7109375" customWidth="1"/><col min="5" max="5" width="12.7109375" customWidth="1"/></cols>',
    `<sheetData>${rows.join('')}</sheetData>`,
  ].join('');
}

export function sampleXlsx(): Buffer {
  return xlsxPackage({
    sheets: [
      { name: 'Summary', body: summarySheet(), relationships: [{ id: 'rId1', type: 'hyperlink', target: GITHUB_URL, external: true }] },
      { name: 'Data', body: dataSheet() },
      { name: 'Hidden notes', state: 'hidden', body: `<sheetData><row r="1">${cell('A1', 'hidden sheet')}</row></sheetData>` },
      {
        name: 'メモ',
        body: `<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><sheetData><row r="1">${cell('A1', '日本語のメモ')}${cell('B1', 42)}</row></sheetData>`,
      },
    ],
    sharedStrings: STRINGS.map((text) => (text.startsWith('<') ? `<si>${text}</si>` : `<si><t xml:space="preserve">${escapeXml(text)}</t></si>`)),
    styles: STYLES,
    theme: THEME,
  });
}

export const samples: Sample[] = [{ fileName: 'sample.xlsx', build: sampleXlsx }];
