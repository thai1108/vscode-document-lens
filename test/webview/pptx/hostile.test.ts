// A presentation is untrusted input: it must not run code, reach the network, navigate the view, or make the viewer
// do unbounded work. These decks try.
import assert from 'node:assert/strict';
import { afterEach, before, describe, it } from 'node:test';
import { Deck } from '../../../src/webview/pptx/deck.ts';
import { PPTX_LIMITS } from '../../../src/webview/pptx/limits.ts';
import { MediaStore } from '../../../src/webview/pptx/media.ts';
import { renderSlide } from '../../../src/webview/pptx/slide.ts';
import { encodePng } from '../../support/png.ts';
import { A, buildPptx, DEFAULT_TABLE_STYLE, P, paragraph, R, run, shape, xfrm } from '../../support/pptxBuilder.ts';
import { byName, installDom, renderAll } from './dom.ts';

const DEFAULT_LIMITS = { ...PPTX_LIMITS };
const box = (id: number, extra: Partial<Parameters<typeof shape>[0]> = {}) =>
  shape({ id, name: `Shape ${id}`, xfrm: [id * 1000, 0, 100000, 100000], geometry: 'rect', ...extra });

async function renderOne(shapes: string, rels: Parameters<typeof buildPptx>[0]['slides'][number]['rels'] = []) {
  const { slides } = await renderAll(buildPptx({ slides: [{ layout: 2, shapes, rels }] }));
  return slides[0] ?? assert.fail('no slide');
}

describe('hostile presentations', () => {
  before(() => installDom());
  afterEach(() => Object.assign(PPTX_LIMITS, DEFAULT_LIMITS));

  it('refuses parts with a DOCTYPE instead of expanding their entities', async () => {
    const laughs = '<!DOCTYPE p:sld [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">]>';
    const bytes = buildPptx({ slides: [{ layout: 2, shapes: '', xml: `${laughs}<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>` }] });
    const deck = await Deck.open(new Uint8Array(bytes));
    await assert.rejects(renderSlide(deck, 0, new MediaStore(deck.pkg)), /ppt\/slides\/slide1\.xml declares a DOCTYPE/);
  });

  it('puts text into the DOM as text', async () => {
    const markup = '<img src=x onerror="alert(1)"><script>alert(2)</script>';
    const slide = await renderOne(box(2, { paragraphs: paragraph(run(markup)) }));
    assert.equal(byName(slide.element, 'Shape 2').textContent, markup);
    assert.equal(slide.element.querySelector('img, script'), null);
  });

  it('links only to http(s) URLs and never loads linked pictures', async () => {
    const link = (rel: string) => `<a:rPr lang="en-US"><a:hlinkClick r:id="${rel}"/></a:rPr>`;
    const shapes =
      box(2, { paragraphs: paragraph(run('script', link('rId2')) + run('file', link('rId3')) + run('web', link('rId4'))) }) +
      `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Linked"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:link="rId5"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm([0, 0, 952500, 952500])}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const slide = await renderOne(shapes, [
      { id: 'rId2', type: 'hyperlink', target: 'javascript:alert(1)', external: true },
      { id: 'rId3', type: 'hyperlink', target: 'file:///etc/passwd', external: true },
      { id: 'rId4', type: 'hyperlink', target: 'https://example.com/ok', external: true },
      { id: 'rId5', type: 'image', target: 'https://tracker.example/pixel.png', external: true },
    ]);
    assert.deepEqual([...slide.element.querySelectorAll('a')].map((a) => a.getAttribute('href')), ['https://example.com/ok']);
    assert.equal(byName(slide.element, 'Linked').textContent, 'Linked picture (not loaded)');
    assert.equal(slide.element.querySelector('image'), null);
  });

  it('caps the shapes, text runs and table cells drawn on one slide', async () => {
    PPTX_LIMITS.maxShapesPerSlide = 20;
    const many = await renderOne(Array.from({ length: 40 }, (_, i) => box(i + 2)).join(''));
    assert.equal(many.truncated, true);
    assert.ok(many.element.querySelectorAll('.pptx-shape').length <= 20);

    Object.assign(PPTX_LIMITS, DEFAULT_LIMITS, { maxRunsPerSlide: 5 });
    const wordy = await renderOne(box(2, { paragraphs: Array.from({ length: 20 }, (_, i) => paragraph(run(`run ${i}`))).join('') }));
    assert.equal(wordy.truncated, true);
    assert.ok(byName(wordy.element, 'Shape 2').querySelectorAll('span').length <= 5);

    Object.assign(PPTX_LIMITS, DEFAULT_LIMITS, { maxTableCells: 4 });
    const cell = '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>';
    const table =
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Big table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
      xfrm([0, 0, 3000000, 1000000], {}, 'p:xfrm') +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr><a:tableStyleId>${DEFAULT_TABLE_STYLE}</a:tableStyleId></a:tblPr><a:tblGrid>${'<a:gridCol w="1000000"/>'.repeat(3)}</a:tblGrid>` +
      `${`<a:tr h="300000">${cell.repeat(3)}</a:tr>`.repeat(3)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
    const big = await renderOne(table);
    assert.equal(big.truncated, true);
    assert.equal(byName(big.element, 'Big table').querySelectorAll('td').length, 4);
  });

  it('stops descending into deeply nested groups', async () => {
    PPTX_LIMITS.maxGroupDepth = 8;
    const group = (inner: string, depth: number) =>
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${100 + depth}" name="Group ${depth}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>${inner}</p:grpSp>`;
    let nested = box(2);
    for (let depth = 0; depth < 30; depth++) {
      nested = group(nested, depth);
    }
    const slide = await renderOne(nested);
    assert.equal(slide.truncated, true);
    assert.equal(slide.element.querySelectorAll('.pptx-group').length, 8);
  });

  it('shows at most the slide limit and says how many it left out', async () => {
    PPTX_LIMITS.maxSlides = 2;
    const deck = await Deck.open(new Uint8Array(buildPptx({ slides: [1, 2, 3, 4].map(() => ({ layout: 2 as const, shapes: '' })) })));
    assert.equal(deck.slideParts.length, 2);
    assert.equal(deck.omittedSlides, 2);
  });

  it('keeps absurd coordinates finite', async () => {
    const slide = await renderOne(box(2, { xfrm: [1e18, -1e18, 1e18, 5] }));
    const shapeBox = byName(slide.element, 'Shape 2');
    for (const property of ['left', 'top', 'width']) {
      assert.match(shapeBox.style.getPropertyValue(property), /^-?\d+(\.\d+)?px$/, property);
    }
  });

  it('keeps the slide when its background picture cannot be read', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 2,
          shapes: box(2, { paragraphs: paragraph(run('still here')) }),
          background: '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>',
          rels: [{ id: 'rId7', type: 'image', target: '../media/bg.png' }],
        },
      ],
      parts: [{ name: 'ppt/media/bg.png', data: encodePng(2, 2, () => [0, 0, 0]) }],
    });
    setCompressionMethod(bytes, 'ppt/media/bg.png', 12);
    const slide = (await renderAll(bytes)).slides[0] ?? assert.fail();
    assert.equal(byName(slide.element, 'Shape 2').textContent, 'still here');
    assert.equal(slide.truncated, true);
    assert.equal(slide.element.querySelector('.pptx-background'), null);
  });
});

/** Patches the compression method of one entry in the central directory, as a damaged or unusual archive has it. */
function setCompressionMethod(zip: Buffer, name: string, method: number): void {
  const encoded = Buffer.from(name);
  for (let offset = 0; offset + 46 + encoded.length <= zip.length; offset++) {
    if (zip.readUInt32LE(offset) === 0x02014b50 && zip.subarray(offset + 46, offset + 46 + encoded.length).equals(encoded)) {
      zip.writeUInt16LE(method, offset + 10);
    }
  }
}

const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

function graphicFrame(id: number, uri: string, data: string): string {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Frame ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${xfrm([0, 0, 952500, 952500], {}, 'p:xfrm')}` +
    `<a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

const smartArt = (id: number) => graphicFrame(id, DGM, `<dgm:relIds xmlns:dgm="${DGM}" r:dm="rIdDm" r:lo="rIdX" r:qs="rIdX" r:cs="rIdX"/>`);
const chartFrame = (id: number) => graphicFrame(id, CHART, `<c:chart xmlns:c="${CHART}" r:id="rIdChart"/>`);

/** A SmartArt drawing holding `content`, with the relationships a slide needs to reach it (and the drawing to reach itself). */
function smartArtDeck(shapes: string, content: string): Buffer {
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const drawing = `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><dsp:spTree>${content}</dsp:spTree></dsp:drawing>`;
  const drawingRels =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rIdDm" Type="${REL}/diagramData" Target="data1.xml"/><Relationship Id="rIdDr" Type="${REL}/diagramDrawing" Target="drawing1.xml"/></Relationships>`;
  return buildPptx({
    slides: [
      {
        layout: 2,
        shapes,
        rels: [
          { id: 'rIdDm', type: 'diagramData', target: '../diagrams/data1.xml' },
          { id: 'rIdDr', type: 'diagramDrawing', target: '../diagrams/drawing1.xml' },
        ],
      },
    ],
    parts: [
      { name: 'ppt/diagrams/data1.xml', data: `<dgm:dataModel xmlns:dgm="${DGM}"/>` },
      { name: 'ppt/diagrams/drawing1.xml', data: drawing },
      { name: 'ppt/diagrams/_rels/drawing1.xml.rels', data: drawingRels },
    ],
  });
}

/** Counts how often the package inflates and decodes each XML part. */
function countReads(deck: Deck): Map<string, number> {
  const reads = new Map<string, number>();
  const read = deck.pkg.text.bind(deck.pkg);
  deck.pkg.text = (name: string) => {
    reads.set(name, (reads.get(name) ?? 0) + 1);
    return read(name);
  };
  return reads;
}

describe('work amplification', () => {
  before(() => installDom());
  afterEach(() => Object.assign(PPTX_LIMITS, DEFAULT_LIMITS));

  it('does not follow SmartArt inside a SmartArt drawing, even one that leads back to itself', async () => {
    const slide = (await renderAll(smartArtDeck(smartArt(2), smartArt(3)))).slides[0] ?? assert.fail();
    const outer = byName(slide.element, 'Frame 2');
    assert.equal(outer.classList.contains('pptx-group'), true, 'the slide’s SmartArt is drawn');
    assert.equal(byName(outer, 'Frame 3').textContent, 'SmartArt graphic', 'the one inside its drawing is only labelled');
    assert.equal(slide.element.querySelectorAll('.pptx-group').length, 1);
    assert.equal(slide.truncated, true);
  });

  it('reads chart and SmartArt parts once per presentation, however many frames and renders use them', async () => {
    const chart = `<c:chartSpace xmlns:c="${CHART}"><c:chart><c:plotArea>${'<c:x/>'.repeat(20000)}<c:pieChart/></c:plotArea></c:chart></c:chartSpace>`;
    const frames = Array.from({ length: 30 }, (_, i) => chartFrame(10 + i)).join('') + Array.from({ length: 30 }, (_, i) => smartArt(100 + i)).join('');
    const bytes = buildPptx({
      slides: [{ layout: 2, shapes: frames, rels: [{ id: 'rIdChart', type: 'chart', target: '../charts/chart1.xml' }, { id: 'rIdDm', type: 'diagramData', target: '../diagrams/data1.xml' }, { id: 'rIdDr', type: 'diagramDrawing', target: '../diagrams/drawing1.xml' }] }],
      parts: [
        { name: 'ppt/charts/chart1.xml', data: chart },
        { name: 'ppt/diagrams/data1.xml', data: `<dgm:dataModel xmlns:dgm="${DGM}"/>` },
        { name: 'ppt/diagrams/drawing1.xml', data: `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="${A}"><dsp:spTree/></dsp:drawing>` },
      ],
    });
    const deck = await Deck.open(new Uint8Array(bytes));
    const reads = countReads(deck);
    const media = new MediaStore(deck.pkg);
    // The main view and the thumbnail draw the same slide.
    for (const slide of [await renderSlide(deck, 0, media), await renderSlide(deck, 0, media)]) {
      assert.deepEqual(new Set([...slide.element.querySelectorAll('.pptx-placeholder')].map((label) => label.textContent)), new Set(['Pie chart']));
      assert.equal(slide.element.querySelectorAll('.pptx-group').length, 30);
    }
    assert.equal(reads.get('ppt/charts/chart1.xml'), 1);
    assert.equal(reads.get('ppt/diagrams/data1.xml'), 1);
    assert.equal(reads.get('ppt/diagrams/drawing1.xml'), 1);
  });

  it('stops reading chart and SmartArt parts past the slide’s byte budget', async () => {
    PPTX_LIMITS.maxPartBytesPerSlide = 100;
    const chart = `<c:chartSpace xmlns:c="${CHART}"><c:chart><c:plotArea>${'<c:x/>'.repeat(100)}<c:barChart/></c:plotArea></c:chart></c:chartSpace>`;
    const bytes = buildPptx({
      slides: [{ layout: 2, shapes: chartFrame(2), rels: [{ id: 'rIdChart', type: 'chart', target: '../charts/chart1.xml' }] }],
      parts: [{ name: 'ppt/charts/chart1.xml', data: chart }],
    });
    const slide = (await renderAll(bytes)).slides[0] ?? assert.fail();
    assert.equal(byName(slide.element, 'Frame 2').textContent, 'Chart');
    assert.equal(slide.truncated, true);
  });

  it('caps gradient stops and defines one gradient per shape, however many paths it has', async () => {
    const custGeom = `<a:custGeom><a:avLst/><a:gdLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${'<a:path w="1" h="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="1" y="1"/></a:lnTo></a:path>'.repeat(300)}</a:pathLst></a:custGeom>`;
    const gradient = `<a:gradFill><a:gsLst>${'<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>'.repeat(2000)}</a:gsLst><a:lin ang="0"/></a:gradFill>`;
    const slide = await renderOne(box(2, { geometry: custGeom, fill: gradient }));
    const drawn = byName(slide.element, 'Shape 2');
    assert.equal(drawn.querySelectorAll('g > path').length, 300);
    assert.equal(drawn.querySelectorAll('linearGradient').length, 1);
    assert.equal(drawn.querySelectorAll('stop').length, PPTX_LIMITS.maxGradientStops);
  });

  it('counts paths, stops and inherited geometry against the slide’s element budget', async () => {
    PPTX_LIMITS.maxElementsPerSlide = 3000;
    // Many shapes sharing one big custom geometry through their layout placeholder.
    const commands = '<a:lnTo><a:pt x="1" y="1"/></a:lnTo>'.repeat(1000);
    const layout =
      `<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr>${xfrm([0, 0, 952500, 952500])}<a:custGeom><a:pathLst><a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${commands}</a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>` +
      '</p:spTree></p:cSld></p:sldLayout>';
    const placeholders = Array.from({ length: 50 }, (_, i) => `<p:sp><p:nvSpPr><p:cNvPr id="${i + 10}" name="P${i}"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`).join('');
    // The builder's own layout part is replaced; the reader keeps the last entry of a name.
    const inherited = (await renderAll(buildPptx({ slides: [{ layout: 2, shapes: placeholders }], parts: [{ name: 'ppt/slideLayouts/slideLayout2.xml', data: layout }] }))).slides[0] ?? assert.fail();
    const segments = [...inherited.element.querySelectorAll('path')].reduce((sum, path) => sum + (path.getAttribute('d')?.split('L').length ?? 1) - 1, 0);
    assert.ok(segments <= 3000, `${segments} segments`);
    assert.equal(inherited.truncated, true);

    // Many shapes sharing one theme gradient through fillRef.
    const style = '<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"/></p:style>';
    const styled = await renderOne(Array.from({ length: 2000 }, (_, i) => box(i + 2, { style })).join(''));
    assert.ok(styled.element.querySelectorAll('stop').length <= 3000);
    assert.equal(styled.truncated, true);
  });

  it('counts empty paragraphs against the run cap', async () => {
    PPTX_LIMITS.maxRunsPerSlide = 5;
    const slide = await renderOne(box(2, { paragraphs: paragraph(run('x')) + '<a:p/>'.repeat(20000) }));
    assert.ok(slide.element.querySelectorAll('.pptx-p').length <= 5);
    assert.equal(slide.truncated, true);
  });
});
