import { quadrantPng } from '../support/png.ts';
import {
  barChart,
  buildPptx,
  DEFAULT_TABLE_STYLE,
  noLine,
  paragraph,
  run,
  scheme,
  shape,
  SHAPE_STYLE,
  solid,
  srgb,
  xfrm,
} from '../support/pptxBuilder.ts';
import type { Sample } from '../support/samples.ts';

// A real four-slide deck built in code: placeholders inheriting the master's text styles, bullets and numbering,
// Japanese text, theme colors with transforms, rotation and flips, a scaled group, a cropped picture, a table with
// merged cells, a hyperlink, a chart (drawn as a placeholder), a custom geometry and a hidden slide.

const JA = '<a:rPr lang="ja-JP" altLang="en-US" dirty="0"/>';
const slideNumber = (id: number) =>
  shape({
    id,
    name: 'Slide Number',
    ph: 'type="sldNum" sz="quarter" idx="12"',
    paragraphs: paragraph('<a:fld id="{3F1D7C2B-8E4A-4C61-9D0E-5B2A7F6C1E90}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>'),
  });
const centered = (text: string, rPr?: string) => paragraph(run(text, rPr), '<a:pPr algn="ctr"/>');

const titleSlide = [
  shape({ id: 2, name: 'Title 1', ph: 'type="ctrTitle"', paragraphs: paragraph(run('Document Lens サンプル', JA)) }),
  shape({
    id: 3,
    name: 'Subtitle 2',
    ph: 'type="subTitle" idx="1"',
    paragraphs: paragraph(run('Synthetic deck · ') + run('example.com/lens', '<a:rPr lang="en-US" dirty="0"><a:hlinkClick r:id="rId2"/></a:rPr>')),
  }),
  slideNumber(4),
].join('');

const numbered = '<a:pPr marL="457200" indent="-457200"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>';
const shapesSlide = [
  shape({ id: 2, name: 'Title 1', ph: 'type="title"', paragraphs: paragraph(run('Shapes and colors')) }),
  shape({
    id: 3,
    name: 'Content Placeholder 2',
    ph: 'idx="1"',
    xfrm: [838200, 1825625, 4800000, 4351338],
    paragraphs: [
      paragraph(run('First point')),
      paragraph(run('Nested point'), '<a:pPr lvl="1"/>'),
      paragraph(run('日本語の箇条書き', JA)),
      paragraph(run('Numbered one'), numbered),
      paragraph(run('Numbered two'), numbered),
    ].join(''),
  }),
  shape({ id: 4, name: 'Accent rect', xfrm: [6000000, 1825625, 1400000, 900000], geometry: 'rect', fill: solid(scheme('accent1', '<a:lumMod val="50000"/>')), line: noLine }),
  shape({
    id: 5,
    name: 'Rounded',
    xfrm: [7600000, 1825625, 1600000, 900000],
    geometry: 'roundRect',
    fill: solid(scheme('accent2')),
    style: SHAPE_STYLE,
    bodyPr: '<a:bodyPr rtlCol="0" anchor="ctr"/>',
    paragraphs: centered('Rounded'),
  }),
  shape({ id: 6, name: 'Translucent ellipse', xfrm: [9400000, 1825625, 1400000, 900000], geometry: 'ellipse', fill: solid(srgb('00B050', '<a:alpha val="50000"/>')), line: noLine }),
  shape({
    id: 7,
    name: 'Rotated triangle',
    xfrm: [6000000, 3000000, 1400000, 1000000],
    rot: 1800000,
    flipH: true,
    geometry: 'rtTriangle',
    fill: solid(scheme('accent6')),
    line: noLine,
    bodyPr: '<a:bodyPr anchor="ctr"/>',
    paragraphs: centered('Rotated', '<a:rPr lang="en-US" sz="1400" b="1" dirty="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr>'),
  }),
  shape({ id: 8, name: 'Star', xfrm: [7700000, 3000000, 1000000, 1000000], geometry: 'star5', fill: solid(scheme('accent4')), line: noLine }),
  shape({ id: 9, name: 'Arrow', xfrm: [8900000, 3100000, 1400000, 800000], geometry: 'rightArrow', fill: solid(scheme('accent5')), line: noLine }),
  shape({ id: 10, name: 'Callout', xfrm: [10500000, 3000000, 1300000, 800000], geometry: 'wedgeRectCallout', style: SHAPE_STYLE, bodyPr: '<a:bodyPr anchor="ctr"/>', paragraphs: centered('Callout') }),
  '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="11" name="Dashed connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>' +
    xfrm([6000000, 4400000, 2800000, 0]) +
    '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:prstDash val="dash"/><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>',
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="12" name="Scaled group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>' +
    '<a:xfrm><a:off x="9000000" y="4300000"/><a:ext cx="2400000" cy="1200000"/><a:chOff x="0" y="0"/><a:chExt cx="1200000" cy="600000"/></a:xfrm></p:grpSpPr>' +
    shape({ id: 13, name: 'Group A', xfrm: [0, 0, 600000, 600000], geometry: 'rect', fill: solid(scheme('accent5')), line: noLine }) +
    shape({ id: 14, name: 'Group B', xfrm: [600000, 0, 600000, 600000], geometry: 'ellipse', fill: solid(scheme('accent3')), line: noLine }) +
    '</p:grpSp>',
  slideNumber(15),
].join('');

const cell = (text: string, attributes = '', properties = '<a:tcPr/>') =>
  `<a:tc${attributes}><a:txBody><a:bodyPr/><a:lstStyle/>${text ? paragraph(run(text)) : '<a:p><a:endParaRPr lang="en-US"/></a:p>'}</a:txBody>${properties}</a:tc>`;
const row = (cells: string) => `<a:tr h="500000">${cells}</a:tr>`;
const table =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Merged table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
  xfrm([3200000, 1825625, 4800000, 1500000], {}, 'p:xfrm') +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
  `<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE}</a:tableStyleId></a:tblPr>` +
  '<a:tblGrid><a:gridCol w="1600000"/><a:gridCol w="1600000"/><a:gridCol w="1600000"/></a:tblGrid>' +
  row(cell('Merged header', ' gridSpan="2"') + cell('', ' hMerge="1"') + cell('Q3')) +
  row(cell('Tall', ' rowSpan="2"') + cell('A') + cell('B')) +
  row(cell('', ' vMerge="1"') + cell('C') + cell('D', '', `<a:tcPr anchor="ctr">${solid(srgb('C00000'))}</a:tcPr>`)) +
  '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';

const pictureSlide = [
  shape({ id: 2, name: 'Title 1', ph: 'type="title"', paragraphs: paragraph(run('Picture and table')) }),
  '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Cropped picture" descr="Four colored quadrants"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rId2"/><a:srcRect l="50000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm([838200, 1825625, 2000000, 3000000])}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></p:spPr></p:pic>`,
  table,
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Sales chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
    xfrm([8400000, 1825625, 3000000, 2200000], {}, 'p:xfrm') +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId3"/></a:graphicData></a:graphic></p:graphicFrame>',
].join('');

const house =
  '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="200" h="150">' +
  '<a:moveTo><a:pt x="0" y="60"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="200" y="60"/></a:lnTo>' +
  '<a:lnTo><a:pt x="200" y="150"/></a:lnTo><a:lnTo><a:pt x="0" y="150"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>';
const hiddenSlide = [
  shape({ id: 2, name: 'Title 1', ph: 'type="title"', paragraphs: paragraph(run('Hidden slide')) }),
  shape({
    id: 3,
    name: 'Gradient rect',
    xfrm: [838200, 1825625, 3000000, 1500000],
    geometry: 'rect',
    fill: `<a:gradFill><a:gsLst><a:gs pos="0">${scheme('accent1')}</a:gs><a:gs pos="100000">${scheme('accent2')}</a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>`,
    line: noLine,
  }),
  shape({ id: 4, name: 'Custom shape', xfrm: [4200000, 1825625, 2000000, 1500000], geometry: house, fill: solid(scheme('accent6')), line: noLine }),
  shape({
    id: 5,
    name: 'Autofit text',
    textBox: true,
    xfrm: [838200, 3600000, 5000000, 800000],
    geometry: 'rect',
    bodyPr: '<a:bodyPr wrap="square"><a:normAutofit fontScale="50000"/></a:bodyPr>',
    paragraphs: paragraph(run('Shrunk to half', '<a:rPr lang="en-US" sz="4000" dirty="0"/>')),
  }),
].join('');

export function sampleDeck(): Buffer {
  return buildPptx({
    slides: [
      { layout: 1, shapes: titleSlide, rels: [{ id: 'rId2', type: 'hyperlink', target: 'https://example.com/lens', external: true }] },
      { layout: 2, shapes: shapesSlide },
      {
        layout: 2,
        shapes: pictureSlide,
        rels: [
          { id: 'rId2', type: 'image', target: '../media/image1.png' },
          { id: 'rId3', type: 'chart', target: '../charts/chart1.xml' },
        ],
      },
      { layout: 2, shapes: hiddenSlide, hidden: true, background: `<p:bg><p:bgPr>${solid(scheme('bg2'))}<a:effectLst/></p:bgPr></p:bg>` },
    ],
    parts: [
      { name: 'ppt/media/image1.png', data: quadrantPng() },
      { name: 'ppt/charts/chart1.xml', contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml', data: barChart([4, 7, 5]) },
    ],
  });
}

export const samples: Sample[] = [{ fileName: 'sample.pptx', build: sampleDeck }];
