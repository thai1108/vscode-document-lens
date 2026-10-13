// A presentation is untrusted input: it must not run code, reach the network, navigate the view, or make the viewer
// do unbounded work. These decks try.
import assert from 'node:assert/strict';
import { afterEach, before, describe, it } from 'node:test';
import { Deck } from '../../../src/webview/pptx/deck.ts';
import { PPTX_LIMITS } from '../../../src/webview/pptx/limits.ts';
import { MediaStore } from '../../../src/webview/pptx/media.ts';
import { renderSlide } from '../../../src/webview/pptx/slide.ts';
import { buildPptx, DEFAULT_TABLE_STYLE, paragraph, run, shape, xfrm } from '../../support/pptxBuilder.ts';
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
});
