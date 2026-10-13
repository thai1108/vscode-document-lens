// Builds real PresentationML packages for tests: content types, package relationships, document properties, a theme,
// one slide master with title/body/other text styles and placeholders, two layouts (Title Slide, Title and Content)
// and the given slides. PowerPoint and macOS Quick Look open the result.
import { writeZip, type ZipInput } from './zipWriter.ts';

export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NAMESPACES = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CONTENT = 'application/vnd.openxmlformats-officedocument';

export const SLIDE_WIDTH = 12192000;
export const SLIDE_HEIGHT = 6858000;
export const DEFAULT_TABLE_STYLE = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

export interface Relationship {
  readonly id: string;
  /** Last segment of the relationship type, e.g. "image", "hyperlink", "chart". */
  readonly type: string;
  readonly target: string;
  readonly external?: boolean;
}

export interface SlideSpec {
  /** 1: Title Slide, 2: Title and Content. */
  readonly layout: 1 | 2;
  /** Shapes inside p:spTree, after its group properties. */
  readonly shapes: string;
  readonly rels?: readonly Relationship[];
  readonly hidden?: boolean;
  /** A p:bg element. */
  readonly background?: string;
  /** The whole slide part instead of one built from `shapes` (for malformed or hostile slides). */
  readonly xml?: string;
  /** Root attributes of p:sld, e.g. ' showMasterSp="0"'. */
  readonly attributes?: string;
}

export interface ExtraPart {
  readonly name: string;
  readonly contentType?: string;
  readonly data: string | Buffer;
}

export interface PptxSpec {
  readonly slides: readonly SlideSpec[];
  /** Additional parts (media, charts, ...); XML strings get the XML declaration. */
  readonly parts?: readonly ExtraPart[];
  /** Body of ppt/tableStyles.xml inside a:tblStyleLst. */
  readonly tableStyles?: string;
  /** Replaces the presentation's p:defaultTextStyle level 1 run defaults (a:defRPr). */
  readonly defaultRunProperties?: string;
}

export function escapeXml(text: string): string {
  return text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
}

export function rels(relationships: readonly Relationship[]): string {
  const items = relationships
    .map((rel) => `<Relationship Id="${rel.id}" Type="${REL}/${rel.type}" Target="${escapeXml(rel.target)}"${rel.external ? ' TargetMode="External"' : ''}/>`)
    .join('');
  return `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}

// Shape helpers ------------------------------------------------------------------------------------------------------

export interface ShapeOptions {
  readonly id: number;
  readonly name: string;
  /** p:ph attributes, e.g. 'type="title"' or 'idx="1"'. */
  readonly ph?: string;
  /** off x, off y, ext cx, ext cy in EMUs. */
  readonly xfrm?: readonly [number, number, number, number];
  readonly rot?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** A preset name, or a full a:custGeom element. */
  readonly geometry?: string;
  readonly fill?: string;
  readonly line?: string;
  readonly effects?: string;
  readonly style?: string;
  readonly bodyPr?: string;
  readonly lstStyle?: string;
  readonly paragraphs?: string;
  readonly textBox?: boolean;
  /** cNvPr children, e.g. a shape-level a:hlinkClick. */
  readonly nvExtra?: string;
}

export function xfrm(values: readonly [number, number, number, number], extra: { rot?: number; flipH?: boolean; flipV?: boolean } = {}, element = 'a:xfrm'): string {
  const [x, y, cx, cy] = values;
  const attributes = `${extra.rot ? ` rot="${extra.rot}"` : ''}${extra.flipH ? ' flipH="1"' : ''}${extra.flipV ? ' flipV="1"' : ''}`;
  return `<${element}${attributes}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></${element}>`;
}

function geometryXml(geometry: string | undefined): string {
  if (!geometry) {
    return '';
  }
  return geometry.startsWith('<') ? geometry : `<a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>`;
}

export function shape(options: ShapeOptions): string {
  const ph = options.ph ? `<p:ph ${options.ph}/>` : '';
  const locks = options.ph ? '<a:spLocks noGrp="1"/>' : '';
  const position = options.xfrm ? xfrm(options.xfrm, options) : '';
  const properties = `${position}${geometryXml(options.geometry)}${options.fill ?? ''}${options.line ?? ''}${options.effects ?? ''}`;
  const text =
    options.paragraphs !== undefined
      ? `<p:txBody>${options.bodyPr ?? '<a:bodyPr/>'}${options.lstStyle ?? '<a:lstStyle/>'}${options.paragraphs}</p:txBody>`
      : '';
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${options.id}" name="${escapeXml(options.name)}">${options.nvExtra ?? ''}</p:cNvPr>` +
    `<p:cNvSpPr${options.textBox ? ' txBox="1"' : ''}>${locks}</p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr>${properties}</p:spPr>${options.style ?? ''}${text}</p:sp>`
  );
}

export function run(text: string, rPr = '<a:rPr lang="en-US" dirty="0"/>'): string {
  return `<a:r>${rPr}<a:t>${escapeXml(text)}</a:t></a:r>`;
}

export function paragraph(content: string, pPr = ''): string {
  return `<a:p>${pPr}${content}<a:endParaRPr lang="en-US" dirty="0"/></a:p>`;
}

export const solid = (color: string) => `<a:solidFill>${color}</a:solidFill>`;
export const scheme = (name: string, transforms = '') => `<a:schemeClr val="${name}">${transforms}</a:schemeClr>`;
export const srgb = (hex: string, transforms = '') => `<a:srgbClr val="${hex}">${transforms}</a:srgbClr>`;
export const noLine = '<a:ln><a:noFill/></a:ln>';

/** The p:style PowerPoint gives new shapes: accent1 fill, darker accent1 outline, light text. */
export const SHAPE_STYLE =
  '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';

// Package parts ------------------------------------------------------------------------------------------------------

const THEME_LINE = (width: number) =>
  `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;

const THEME = `<a:theme xmlns:a="${A}" name="Lens">
<a:themeElements>
<a:clrScheme name="Lens">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Lens">
<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Lens">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="73000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>${THEME_LINE(6350)}${THEME_LINE(12700)}${THEME_LINE(19050)}</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>`;

const TREE_START =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const TEXT_DEFAULTS = (size: number, font: 'mj' | 'mn', color = 'tx1') =>
  `<a:defRPr sz="${size}" kern="1200"><a:solidFill><a:schemeClr val="${color}"/></a:solidFill><a:latin typeface="+${font}-lt"/><a:ea typeface="+${font}-ea"/><a:cs typeface="+${font}-cs"/></a:defRPr>`;

const BODY_LEVELS = [
  [228600, 2800, 1000],
  [685800, 2400, 500],
  [1143000, 2000, 500],
  [1600200, 1800, 500],
  [2057400, 1800, 500],
];

const BODY_STYLE = BODY_LEVELS.map(
  ([marL, size, before], index) =>
    `<a:lvl${index + 1}pPr marL="${marL}" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${before}"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/>${TEXT_DEFAULTS(size ?? 1800, 'mn')}</a:lvl${index + 1}pPr>`,
).join('');

const FOOTER = (type: string, idx: number, x: number, align: string, content: string, size = 'quarter') =>
  shape({
    id: 10 + idx,
    name: `${type} placeholder`,
    ph: `type="${type}" sz="${size}" idx="${idx}"`,
    xfrm: [x, 6356350, type === 'ftr' ? 4114800 : 2743200, 365125],
    bodyPr: '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>',
    lstStyle: `<a:lstStyle><a:lvl1pPr algn="${align}">${TEXT_DEFAULTS(1200, 'mn').replace('<a:schemeClr val="tx1"/>', '<a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr>')}</a:lvl1pPr></a:lstStyle>`,
    paragraphs: paragraph(content),
  });

const MASTER = `<p:sldMaster ${NAMESPACES}>
<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${TREE_START}
${shape({ id: 2, name: 'Title Placeholder 1', ph: 'type="title"', xfrm: [838200, 365125, 10515600, 1325563], geometry: 'rect', bodyPr: '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>', paragraphs: paragraph(run('Click to edit Master title style')) })}
${shape({ id: 3, name: 'Text Placeholder 2', ph: 'type="body" idx="1"', xfrm: [838200, 1825625, 10515600, 4351338], geometry: 'rect', bodyPr: '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr>', paragraphs: paragraph(run('Click to edit Master text styles'), '<a:pPr lvl="0"/>') + paragraph(run('Second level'), '<a:pPr lvl="1"/>') })}
${FOOTER('dt', 2, 838200, 'l', '<a:fld id="{B3E7A0A4-1D6E-4C2E-9A57-6C4C1A2D3E01}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>10/13/2026</a:t></a:fld>', 'half')}
${FOOTER('ftr', 3, 4038600, 'ctr', '')}
${FOOTER('sldNum', 4, 8610600, 'r', '<a:fld id="{B3E7A0A4-1D6E-4C2E-9A57-6C4C1A2D3E02}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>')}
${shape({ id: 7, name: 'Master accent bar', xfrm: [0, 6705600, SLIDE_WIDTH, 152400], geometry: 'rect', fill: solid(scheme('accent1')), line: noLine })}
</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>
<p:txStyles>
<p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>${TEXT_DEFAULTS(4400, 'mj')}</a:lvl1pPr></p:titleStyle>
<p:bodyStyle>${BODY_STYLE}</p:bodyStyle>
<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">${TEXT_DEFAULTS(1800, 'mn')}</a:lvl1pPr></p:otherStyle>
</p:txStyles>
</p:sldMaster>`;

const LAYOUT_FOOTERS = ['dt', 'ftr', 'sldNum'].map((type, i) => shape({ id: 4 + i, name: `${type} ${i}`, ph: `type="${type}" sz="${type === 'dt' ? 'half' : 'quarter'}" idx="${10 + i}"`, paragraphs: paragraph('') })).join('');

const TITLE_LAYOUT = `<p:sldLayout ${NAMESPACES} type="title" preserve="1">
<p:cSld name="Title Slide"><p:spTree>${TREE_START}
${shape({ id: 2, name: 'Title 1', ph: 'type="ctrTitle"', xfrm: [1524000, 1122363, 9144000, 2387600], bodyPr: '<a:bodyPr anchor="b"/>', lstStyle: '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle>', paragraphs: paragraph(run('Click to edit Master title style')) })}
${shape({ id: 3, name: 'Subtitle 2', ph: 'type="subTitle" idx="1"', xfrm: [1524000, 3602038, 9144000, 1655762], bodyPr: '<a:bodyPr/>', lstStyle: '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>', paragraphs: paragraph(run('Click to edit Master subtitle style')) })}
${LAYOUT_FOOTERS}
</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;

const CONTENT_LAYOUT = `<p:sldLayout ${NAMESPACES} type="obj" preserve="1">
<p:cSld name="Title and Content"><p:spTree>${TREE_START}
${shape({ id: 2, name: 'Title 1', ph: 'type="title"', paragraphs: paragraph(run('Click to edit Master title style')) })}
${shape({ id: 3, name: 'Content Placeholder 2', ph: 'idx="1"', paragraphs: paragraph(run('Click to edit Master text styles')) })}
${LAYOUT_FOOTERS}
</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;

const PRESENTATION = (slideCount: number, defaultRunProperties = TEXT_DEFAULTS(1800, 'mn')) => `<p:presentation ${NAMESPACES} saveSubsetFonts="1">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>${Array.from({ length: slideCount }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${10 + i}"/>`).join('')}</p:sldIdLst>
<p:sldSz cx="${SLIDE_WIDTH}" cy="${SLIDE_HEIGHT}"/><p:notesSz cx="6858000" cy="9144000"/>
<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">${defaultRunProperties}</a:lvl1pPr></p:defaultTextStyle>
</p:presentation>`;

function slideXml(slide: SlideSpec): string {
  return (
    slide.xml ??
    `<p:sld ${NAMESPACES}${slide.hidden ? ' show="0"' : ''}${slide.attributes ?? ''}><p:cSld>${slide.background ?? ''}<p:spTree>${TREE_START}${slide.shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  );
}

export function buildPptx(spec: PptxSpec): Buffer {
  const files: ZipInput[] = [];
  const overrides: [string, string][] = [];
  const add = (name: string, data: string | Buffer, contentType?: string) => {
    files.push({ name, data: typeof data === 'string' && name.endsWith('.xml') ? XML_HEADER + data : data });
    if (contentType) {
      overrides.push([name, contentType]);
    }
  };
  const presentationML = `${CONTENT}.presentationml`;
  add('ppt/presentation.xml', PRESENTATION(spec.slides.length, spec.defaultRunProperties), `${presentationML}.presentation.main+xml`);
  add(
    'ppt/_rels/presentation.xml.rels',
    rels([
      { id: 'rId1', type: 'slideMaster', target: 'slideMasters/slideMaster1.xml' },
      { id: 'rId2', type: 'presProps', target: 'presProps.xml' },
      { id: 'rId3', type: 'viewProps', target: 'viewProps.xml' },
      { id: 'rId4', type: 'theme', target: 'theme/theme1.xml' },
      { id: 'rId5', type: 'tableStyles', target: 'tableStyles.xml' },
      ...spec.slides.map((_slide, i) => ({ id: `rId${10 + i}`, type: 'slide', target: `slides/slide${i + 1}.xml` })),
    ]),
  );
  add('ppt/slideMasters/slideMaster1.xml', MASTER, `${presentationML}.slideMaster+xml`);
  add(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    rels([
      { id: 'rId1', type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' },
      { id: 'rId2', type: 'slideLayout', target: '../slideLayouts/slideLayout2.xml' },
      { id: 'rId3', type: 'theme', target: '../theme/theme1.xml' },
    ]),
  );
  for (const [index, layout] of [TITLE_LAYOUT, CONTENT_LAYOUT].entries()) {
    add(`ppt/slideLayouts/slideLayout${index + 1}.xml`, layout, `${presentationML}.slideLayout+xml`);
    add(`ppt/slideLayouts/_rels/slideLayout${index + 1}.xml.rels`, rels([{ id: 'rId1', type: 'slideMaster', target: '../slideMasters/slideMaster1.xml' }]));
  }
  add('ppt/theme/theme1.xml', THEME, `${CONTENT}.theme+xml`);
  add('ppt/presProps.xml', `<p:presentationPr ${NAMESPACES}/>`, `${presentationML}.presProps+xml`);
  add('ppt/viewProps.xml', `<p:viewPr ${NAMESPACES}><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`, `${presentationML}.viewProps+xml`);
  add('ppt/tableStyles.xml', `<a:tblStyleLst xmlns:a="${A}" def="${DEFAULT_TABLE_STYLE}">${spec.tableStyles ?? ''}</a:tblStyleLst>`, `${presentationML}.tableStyles+xml`);
  for (const [index, slide] of spec.slides.entries()) {
    add(`ppt/slides/slide${index + 1}.xml`, slideXml(slide), `${presentationML}.slide+xml`);
    add(`ppt/slides/_rels/slide${index + 1}.xml.rels`, rels([{ id: 'rId1', type: 'slideLayout', target: `../slideLayouts/slideLayout${slide.layout}.xml` }, ...(slide.rels ?? [])]));
  }
  for (const part of spec.parts ?? []) {
    add(part.name, part.data, part.contentType);
  }
  add(
    'docProps/core.xml',
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Document Lens sample</dc:title></cp:coreProperties>',
    'application/vnd.openxmlformats-package.core-properties+xml',
  );
  add(
    'docProps/app.xml',
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Document Lens tests</Application><Slides>${spec.slides.length}</Slides></Properties>`,
    `${CONTENT}.extended-properties+xml`,
  );
  const types =
    `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
    overrides.map(([name, type]) => `<Override PartName="/${name}" ContentType="${type}"/>`).join('') +
    '</Types>';
  const packageRels = rels([{ id: 'rId1', type: 'officeDocument', target: 'ppt/presentation.xml' }])
    .replace(
      '</Relationships>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
    );
  return writeZip([{ name: '[Content_Types].xml', data: types }, { name: '_rels/.rels', data: packageRels }, ...files]);
}

/** A bar chart part with literal data, so the package needs no embedded workbook. */
export function barChart(values: readonly number[]): string {
  const categories = values.map((_value, i) => `<c:pt idx="${i}"><c:v>${String.fromCharCode(65 + i)}</c:v></c:pt>`).join('');
  const points = values.map((value, i) => `<c:pt idx="${i}"><c:v>${value}</c:v></c:pt>`).join('');
  return (
    `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>` +
    '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>' +
    `<c:cat><c:strLit><c:ptCount val="${values.length}"/>${categories}</c:strLit></c:cat>` +
    `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numLit></c:val></c:ser>` +
    '<c:axId val="111"/><c:axId val="222"/></c:barChart>' +
    '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>' +
    '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>' +
    '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
  );
}
