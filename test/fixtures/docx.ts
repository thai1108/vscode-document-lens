import { png } from '../support/png.ts';
import type { Sample } from '../support/samples.ts';
import { CONTENT_TYPE, escapeXml, inlinePicture, paragraph, RELATIONSHIP, run, WORD_NAMESPACES, wordPackage } from '../support/wordPackage.ts';

/** The external link in sample.docx. */
export const SAMPLE_LINK = 'https://github.com/thai1108/vscode-document-lens';
/** EMU per CSS pixel (914400 per inch, 96 px per inch). */
const EMU_PER_PX = 9525;

const STYLES = `<w:styles ${WORD_NAMESPACES}>
<w:docDefaults>
  <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Yu Mincho" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:sz w:val="56"/><w:color w:val="1F3864"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="2F5496"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/><w:ind w:left="720"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders>
  <w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/>
  <w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/>
  <w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/>
</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`;

const NUMBERING = `<w:numbering ${WORD_NAMESPACES}>
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const FOOTNOTES = `<w:footnotes ${WORD_NAMESPACES}>
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="1">${paragraph(`<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${run(' A footnote at the foot of the page.')}`)}</w:footnote>
</w:footnotes>`;

const heading = (text: string, extra = '') => paragraph(`${extra}${run(text)}`, '<w:pStyle w:val="Heading1"/>');
const listItem = (numId: number, text: string) =>
  paragraph(run(text), `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`);
const cell = (content: string, properties = '') => `<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ''}${paragraph(content)}</w:tc>`;
const link = (attributes: string, text: string) => `<w:hyperlink ${attributes}>${run(text, '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>`;

/** A font embedded in sample.docx (but used by no text, so the browser never decodes it). */
export const EMBEDDED_FONT = 'Lens Embedded';
const FONT_TABLE = `<w:fonts ${WORD_NAMESPACES}>
<w:font w:name="Calibri"><w:family w:val="swiss"/></w:font>
<w:font w:name="${EMBEDDED_FONT}"><w:embedRegular r:id="rIdFont1" w:fontKey="{00000000-0000-0000-0000-000000000000}"/></w:font>
</w:fonts>`;

/** A horizontal gradient, so a decoded picture is easy to tell from a placeholder. */
const gradient = (width: number, height: number) => png(width, height, (x) => [Math.round((255 * x) / (width - 1)), 120, 215, 255]);

function sampleDocx(): Uint8Array {
  const body = [
    paragraph(run('Document Lens sample'), '<w:pStyle w:val="Title"/>'),
    heading('Formatting'),
    paragraph(
      run('Plain, ') +
        run('bold', '<w:b/>') +
        run(', ') +
        run('italic', '<w:i/>') +
        run(' and ') +
        run('red', '<w:color w:val="C00000"/>') +
        run(' text with a footnote.') +
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>',
    ),
    heading('日本語'),
    paragraph(run('これは日本語のテキストです。文書をそのまま表示します。', '<w:lang w:eastAsia="ja-JP"/>')),
    heading('Lists'),
    listItem(1, 'First bullet'),
    listItem(1, 'Second bullet'),
    paragraph(''),
    listItem(2, 'First step'),
    listItem(2, 'Second step'),
    listItem(2, 'Third step'),
    heading('Table'),
    '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell(run('Merged across'), '<w:gridSpan w:val="2"/>')}${cell(run('C1'))}</w:tr>` +
      `<w:tr>${cell(run('Merged down'), '<w:vMerge w:val="restart"/>')}${cell(run('B2'))}${cell(run('C2'))}</w:tr>` +
      `<w:tr>${cell('', '<w:vMerge/>')}${cell(run('B3'))}${cell(run('C3'))}</w:tr>` +
      '</w:tbl>',
    heading('Picture'),
    paragraph(inlinePicture('rIdImage', 200 * EMU_PER_PX, 100 * EMU_PER_PX, 'Gradient')),
    heading('Links'),
    paragraph(link('r:id="rIdLink"', 'Document Lens on GitHub')),
    paragraph(link('w:anchor="details"', 'Jump to the details')),
    paragraph('<w:r><w:br w:type="page"/></w:r>'),
    heading('Details', '<w:bookmarkStart w:id="0" w:name="details"/><w:bookmarkEnd w:id="0"/>'),
    paragraph(run('The second page, reached through the bookmark link.')),
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/>' +
      '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>',
  ].join('');
  return wordPackage({
    body,
    relationships: [{ id: 'rIdLink', type: RELATIONSHIP.hyperlink, target: SAMPLE_LINK, external: true }],
    parts: [
      { name: 'word/styles.xml', contentType: CONTENT_TYPE.styles, xml: STYLES, relationship: { id: 'rIdStyles', type: RELATIONSHIP.styles } },
      { name: 'word/numbering.xml', contentType: CONTENT_TYPE.numbering, xml: NUMBERING, relationship: { id: 'rIdNumbering', type: RELATIONSHIP.numbering } },
      { name: 'word/footnotes.xml', contentType: CONTENT_TYPE.footnotes, xml: FOOTNOTES, relationship: { id: 'rIdFootnotes', type: RELATIONSHIP.footnotes } },
      {
        name: 'word/header1.xml',
        contentType: CONTENT_TYPE.header,
        xml: `<w:hdr ${WORD_NAMESPACES}>${paragraph(run('Document Lens sample · header'), '<w:jc w:val="right"/>')}</w:hdr>`,
        relationship: { id: 'rIdHeader', type: RELATIONSHIP.header },
      },
      {
        name: 'word/footer1.xml',
        contentType: CONTENT_TYPE.footer,
        xml: `<w:ftr ${WORD_NAMESPACES}>${paragraph(run('Footer · フッター'), '<w:jc w:val="center"/>')}</w:ftr>`,
        relationship: { id: 'rIdFooter', type: RELATIONSHIP.footer },
      },
      {
        name: 'word/fontTable.xml',
        contentType: CONTENT_TYPE.fontTable,
        xml: FONT_TABLE,
        relationship: { id: 'rIdFontTable', type: RELATIONSHIP.fontTable },
        relationships: [{ id: 'rIdFont1', type: RELATIONSHIP.font, target: 'fonts/font1.odttf' }],
      },
    ],
    media: [
      { name: 'word/media/image1.png', contentType: 'image/png', data: gradient(200, 100), relationship: { id: 'rIdImage', type: RELATIONSHIP.image } },
      // With an all-zero key, deobfuscating leaves the bytes as they are; they are not a real font.
      { name: 'word/fonts/font1.odttf', contentType: CONTENT_TYPE.obfuscatedFont, data: Buffer.alloc(64, 0x2a) },
    ],
  });
}

/** Lengths far beyond anything Word writes (twips, half-points and EMU). */
const HUGE = '99999999';
const HUGE_EMU = '99999999999';

/** Theme font name that injects declarations: var() substitutions that would load a picture or make a huge length. */
const THEME_INJECTION =
  'x; --u: "https://example.invalid/beacon.png"; --s: 30; --k: 1000000; border: 10px solid; ' +
  'border-image: url(https://example.invalid/literal.png) var(--s); border-image-source: image-set(var(--u) 1x); ' +
  'margin-left: calc(1pt * var(--k)); font-size: calc(1pt * var(--k))';
/** Embedded font name that would turn its @font-face into one for a font the viewer's own UI uses. */
export const FONT_INJECTION = 'X; font-family: BlinkMacSystemFont; size-adjust: 1000%';

/**
 * A document that tries what a hostile file might: script and local-file links, pictures from the web, VML and style
 * values that try to escape the page, injected CSS rules and var() substitutions, an embedded font posing as a UI
 * font, absurd sizes, an embedded HTML chunk and bookmarks named like the viewer's own elements.
 */
function hostileDocx(): Uint8Array {
  const external = (id: string, type: string, target: string) => ({ id, type, target, external: true });
  const styles = `<w:styles ${WORD_NAMESPACES}>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Injected"><w:name w:val="Injected"/><w:rPr>
  <w:rFonts w:ascii="x;}body,.toolbar,#error{display:none!important}.y{color:red"/>
  <w:color w:val="000000; background-image: url(https://example.invalid/css.png)"/>
</w:rPr></w:style>
<w:style w:type="character" w:styleId="Escaping"><w:name w:val="Escaping"/><w:rPr>
  <w:color w:val="000000} :scope, .dl-docx, .dl-docx-frame { position: fixed !important; inset: 0 !important; z-index: 2147483647 !important; background: red !important; contain: none !important } .y { color: red"/>
</w:rPr></w:style>
</w:styles>`;
  const theme =
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>' +
    '<a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1></a:clrScheme>' +
    `<a:fontScheme name="F"><a:majorFont><a:latin typeface="${escapeXml(THEME_INJECTION)}"/></a:majorFont>` +
    '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>';
  const fontTable =
    `<w:fonts ${WORD_NAMESPACES}><w:font w:name="${escapeXml(FONT_INJECTION)}">` +
    '<w:embedRegular r:id="rIdFont1" w:fontKey="{00000000-0000-0000-0000-000000000000}"/></w:font></w:fonts>';
  const vml =
    '<w:r><w:pict><v:rect style="position:fixed;left:0;top:0;width:99999px;height:99999px;z-index:2147483647;' +
    'background-image:url(https://example.invalid/vml-style.png)" fillcolor="url(https://example.invalid/fill.png)">' +
    '<v:imagedata r:id="rIdVmlRemote"/></v:rect></w:pict></w:r>';
  const body = [
    paragraph(run('Hostile sample'), '<w:pStyle w:val="Injected"/>'),
    paragraph(`<w:hyperlink r:id="rIdScript">${run('script link')}</w:hyperlink>`),
    paragraph(`<w:hyperlink r:id="rIdVbscript">${run('vbscript link')}</w:hyperlink>`),
    paragraph(`<w:hyperlink r:id="rIdDataHtml">${run('data link')}</w:hyperlink>`),
    paragraph(`<w:hyperlink r:id="rIdFile">${run('file link')}</w:hyperlink>`),
    paragraph(`<w:hyperlink r:id="rIdCredentials">${run('link with credentials')}</w:hyperlink>`),
    paragraph(`<w:hyperlink w:anchor="error">${run('bookmark named error')}</w:hyperlink>`),
    paragraph(`<w:bookmarkStart w:id="0" w:name="error"/><w:bookmarkStart w:id="1" w:name="content"/>${run('Bookmarks')}`),
    paragraph(inlinePicture('rIdRemote', 914400, 914400, 'Remote picture')),
    paragraph(inlinePicture('rIdMissing', 914400, 914400, 'Missing picture')),
    paragraph(inlinePicture('rIdTiny', Number(HUGE_EMU), Number(HUGE_EMU), 'Huge picture')),
    paragraph(vml),
    paragraph(run('Huge text', `<w:sz w:val="${HUGE}"/>`), `<w:spacing w:line="${HUGE}" w:lineRule="auto"/><w:ind w:left="${HUGE}"/>`),
    `<w:tbl><w:tblPr><w:tblW w:w="${HUGE}" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="${HUGE}"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="${HUGE}" w:type="dxa"/></w:tcPr>${paragraph(run('Huge table'))}</w:tc></w:tr></w:tbl>`,
    '<w:altChunk r:id="rIdChunk"/>',
    `<w:sectPr><w:pgSz w:w="${HUGE}" w:h="${HUGE}"/><w:pgMar w:top="${HUGE}" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/><w:cols w:num="100000"/></w:sectPr>`,
  ].join('');
  return wordPackage({
    body,
    relationships: [
      external('rIdScript', RELATIONSHIP.hyperlink, 'javascript:alert(document.domain)'),
      external('rIdVbscript', RELATIONSHIP.hyperlink, 'vbscript:msgbox(1)'),
      external('rIdDataHtml', RELATIONSHIP.hyperlink, 'data:text/html,<script>alert(1)</script>'),
      external('rIdFile', RELATIONSHIP.hyperlink, 'file:///etc/passwd'),
      external('rIdCredentials', RELATIONSHIP.hyperlink, 'https://user:secret@example.com/'),
      external('rIdRemote', RELATIONSHIP.image, 'https://example.invalid/tracker.png'),
      external('rIdVmlRemote', RELATIONSHIP.image, 'https://example.invalid/vml.png'),
      { id: 'rIdMissing', type: RELATIONSHIP.image, target: 'media/missing.png' },
    ],
    parts: [
      { name: 'word/styles.xml', contentType: CONTENT_TYPE.styles, xml: styles, relationship: { id: 'rIdStyles', type: RELATIONSHIP.styles } },
      { name: 'word/theme/theme1.xml', contentType: CONTENT_TYPE.theme, xml: theme, relationship: { id: 'rIdTheme', type: RELATIONSHIP.theme } },
      {
        name: 'word/fontTable.xml',
        contentType: CONTENT_TYPE.fontTable,
        xml: fontTable,
        relationship: { id: 'rIdFontTable', type: RELATIONSHIP.fontTable },
        relationships: [{ id: 'rIdFont1', type: RELATIONSHIP.font, target: 'fonts/font1.odttf' }],
      },
      {
        name: 'word/chunk.xml',
        contentType: 'application/xhtml+xml',
        xml: '<html xmlns="http://www.w3.org/1999/xhtml"><body><script>alert(1)</script><iframe src="https://example.invalid/"/></body></html>',
        relationship: { id: 'rIdChunk', type: RELATIONSHIP.altChunk },
      },
    ],
    media: [
      { name: 'word/media/tiny.png', contentType: 'image/png', data: gradient(2, 2), relationship: { id: 'rIdTiny', type: RELATIONSHIP.image } },
      { name: 'word/fonts/font1.odttf', contentType: CONTENT_TYPE.obfuscatedFont, data: Buffer.alloc(64, 0x2a) },
    ],
  });
}

export const samples: Sample[] = [
  { fileName: 'sample.docx', build: sampleDocx },
  { fileName: 'hostile.docx', build: hostileDocx },
];
