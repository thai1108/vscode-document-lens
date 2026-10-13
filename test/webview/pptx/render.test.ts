// What the renderer builds for the synthetic deck (test/fixtures/pptx.ts). happy-dom does no layout, so this checks
// structure, positions and resolved properties; test/e2e/pptx.spec.ts checks the pixels in Chrome.
import assert from 'node:assert/strict';
import { before, describe, it } from 'node:test';
import type { RenderedSlide } from '../../../src/webview/pptx/slide.ts';
import { sampleDeck } from '../../fixtures/pptx.ts';
import { encodePng } from '../../support/png.ts';
import { buildPptx, paragraph, run, shape, solid, srgb, xfrm } from '../../support/pptxBuilder.ts';
import { byName, installDom, renderAll } from './dom.ts';

const px = (emu: number) => `${Math.round((emu / 9525) * 100) / 100}px`;
const css = (element: Element | null | undefined, property: string) => (element as HTMLElement | null | undefined)?.style.getPropertyValue(property);

let slides: RenderedSlide[];

describe('rendering the sample deck', () => {
  before(async () => {
    installDom();
    slides = (await renderAll(sampleDeck())).slides;
  });

  it('draws every slide at the slide size, with the master and layout shapes', () => {
    assert.equal(slides.length, 4);
    for (const slide of slides) {
      assert.equal(css(slide.element, 'width'), '1280px');
      assert.equal(css(slide.element, 'height'), '720px');
      assert.equal(byName(slide.element, 'Master accent bar').querySelector('path')?.getAttribute('fill'), '#4472c4');
      assert.equal(slide.truncated, false);
    }
    assert.deepEqual(slides.map((slide) => slide.hidden), [false, false, false, true]);
  });

  it('places title and body placeholders where their layout and master put them, with inherited text styles', () => {
    const [title, shapes] = slides;
    const centerTitle = byName(title?.element ?? assert.fail(), 'Title 1');
    assert.deepEqual([css(centerTitle, 'left'), css(centerTitle, 'top'), css(centerTitle, 'width')], [px(1524000), px(1122363), px(9144000)]);
    assert.equal(centerTitle.textContent, 'Document Lens サンプル');
    assert.ok(centerTitle.querySelector('.pptx-anchor-bottom'), 'anchored at the bottom by the layout');
    const titleRun = centerTitle.querySelector('.pptx-p span');
    assert.equal(css(titleRun, 'font-size'), '80px', '60 pt from the layout list style');
    assert.match(css(titleRun, 'font-family') ?? '', /^"Calibri Light", .*"游ゴシック Light"/, 'major latin font, then the Japanese theme font');
    assert.equal(css(centerTitle.querySelector('.pptx-p'), 'text-align'), 'center');

    const body = byName(shapes?.element ?? assert.fail(), 'Content Placeholder 2');
    assert.deepEqual([css(body, 'left'), css(body, 'width')], [px(838200), px(4800000)], 'own position overrides the master');
    const paragraphs = [...body.querySelectorAll('.pptx-p')];
    assert.deepEqual(paragraphs.map((p) => p.querySelector('.pptx-bullet')?.textContent), ['•', '•', '•', '1.', '2.']);
    assert.deepEqual(paragraphs.map((p) => css(p, 'font-size')), ['37.33px', '32px', '37.33px', '37.33px', '37.33px']);
    assert.deepEqual(paragraphs.map((p) => css(p, 'padding-left')), ['24px', '72px', '24px', '48px', '48px']);
    assert.equal(css(paragraphs[0], 'margin-top'), '13.33px', 'space before from the master body style');
    assert.equal(css(paragraphs[0], 'line-height'), '1.08');
    assert.equal(paragraphs[2]?.textContent, '•日本語の箇条書き');
  });

  it('resolves theme colors with their transforms, style references and transparency', () => {
    const slide = slides[1]?.element ?? assert.fail();
    assert.equal(byName(slide, 'Accent rect').querySelector('path')?.getAttribute('fill'), '#203864');
    const ellipse = byName(slide, 'Translucent ellipse').querySelector('path');
    assert.deepEqual([ellipse?.getAttribute('fill'), ellipse?.getAttribute('fill-opacity')], ['#00b050', '0.5']);
    const rounded = byName(slide, 'Rounded');
    const outline = rounded.querySelector('path');
    assert.equal(outline?.getAttribute('fill'), '#ed7d31');
    assert.equal(outline?.getAttribute('stroke-width'), '1.33', 'lnRef 2 is the theme’s 1 pt line');
    assert.equal(css(rounded.querySelector('.pptx-p span'), 'color'), '#ffffff', 'fontRef lt1');
  });

  it('rotates and flips shapes without mirroring their text, and scales groups', () => {
    const slide = slides[1]?.element ?? assert.fail();
    const triangle = byName(slide, 'Rotated triangle');
    assert.equal(css(triangle, 'transform'), 'rotate(30deg)');
    assert.match(triangle.querySelector('g')?.getAttribute('transform') ?? '', /^matrix\(-1 0 0 1 146\.98 0\)$/);
    assert.equal(css(triangle.querySelector('.pptx-text'), 'transform'), '', 'text is not mirrored');

    const group = byName(slide, 'Scaled group');
    assert.deepEqual([css(group, 'left'), css(group, 'width')], [px(9000000), px(2400000)]);
    const [first, second] = [byName(group, 'Group A'), byName(group, 'Group B')];
    assert.deepEqual([css(first, 'left'), css(first, 'width')], ['0px', px(1200000)]);
    assert.deepEqual([css(second, 'left'), css(second, 'width')], [px(1200000), px(1200000)]);

    const connector = byName(slide, 'Dashed connector').querySelector('g > path');
    assert.equal(connector?.getAttribute('stroke-dasharray'), '8 6');
    assert.match(connector?.getAttribute('marker-end') ?? '', /^url\(#pptx-arrow-\d+\)$/);
    assert.equal(byName(slide, 'Star').querySelector('path')?.getAttribute('d')?.split('L').length, 10);
  });

  it('crops pictures, merges table cells and stands in for charts', () => {
    const slide = slides[2]?.element ?? assert.fail();
    const picture = byName(slide, 'Cropped picture');
    assert.equal(picture.getAttribute('aria-label'), 'Four colored quadrants');
    const image = picture.querySelector('pattern image');
    assert.match(image?.getAttribute('href') ?? '', /^blob:/);
    assert.deepEqual([image?.getAttribute('x'), image?.getAttribute('width')], ['-209.97', '419.95'], 'left half cut off');

    const cells = [...byName(slide, 'Merged table').querySelectorAll('td')];
    assert.deepEqual(cells.map((cell) => cell.textContent), ['Merged header', 'Q3', 'Tall', 'A', 'B', 'C', 'D']);
    assert.equal(cells[0]?.colSpan, 2);
    assert.equal(cells[2]?.rowSpan, 2);
    assert.equal(css(cells[0], 'background-color'), '#4472c4', 'header row of the default table style');
    assert.equal(css(cells[0]?.querySelector('span'), 'font-weight'), '700');
    assert.equal(css(cells[3], 'background-color'), '#cfd5ea', 'first band');
    assert.equal(css(cells[6], 'background-color'), '#c00000', 'the cell’s own fill wins');
    assert.equal(css(cells[0], 'border-bottom'), '4px solid #ffffff');

    assert.equal(byName(slide, 'Sales chart').querySelector('.pptx-placeholder')?.textContent, 'Bar chart');
  });

  it('fills in slide numbers, links and the hidden slide’s own background', () => {
    const [title, , , hidden] = slides;
    const link = title?.element.querySelector('a.pptx-link');
    assert.equal(link?.getAttribute('href'), 'https://example.com/lens');
    assert.equal(css(link?.querySelector('span'), 'color'), '#0563c1');
    assert.equal(byName(title?.element ?? assert.fail(), 'Slide Number').textContent, '1');
    assert.equal(byName(slides[1]?.element ?? assert.fail(), 'Slide Number').textContent, '2');

    assert.equal(css(hidden?.element, 'background-color'), '#e7e6e6');
    assert.equal(byName(hidden?.element ?? assert.fail(), 'Custom shape').querySelector('path')?.getAttribute('d'), 'M0 62.99 L104.99 0 L209.97 62.99 L209.97 157.48 L0 157.48 Z');
    assert.equal(css(byName(hidden?.element ?? assert.fail(), 'Autofit text').querySelector('.pptx-p span'), 'font-size'), '26.67px', '40 pt at the stored 50% font scale');
    assert.match(byName(hidden?.element ?? assert.fail(), 'Gradient rect').querySelector('path')?.getAttribute('fill') ?? '', /^url\(#pptx-gradient-\d+\)$/);
  });
});

describe('graphic frames and media', () => {
  before(() => installDom());

  const frame = (id: number, name: string, uri: string, data: string) =>
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${xfrm([0, 0, 1905000, 952500], {}, 'p:xfrm')}` +
    `<a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`;
  const picture = (id: number, name: string, rel: string, nvPr = '') =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr>${nvPr}</p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${rel}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrm([0, 0, 952500, 952500])}<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const emf = Buffer.alloc(64);
  emf.set([1, 0, 0, 0]);
  emf.set([0x20, 0x45, 0x4d, 0x46], 40);

  it('draws SmartArt from the drawing PowerPoint stores with it, and labels what it cannot draw', async () => {
    const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
    const drawing =
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>` +
      `<dsp:sp modelId="{1}"><dsp:nvSpPr><dsp:cNvPr id="0" name="Step"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="952500" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></dsp:spPr>` +
      `<dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Plan</a:t></a:r></a:p></dsp:txBody><dsp:txXfrm><a:off x="952500" y="238125"/><a:ext cx="952500" cy="238125"/></dsp:txXfrm></dsp:sp></dsp:spTree></dsp:drawing>`;
    const dataModel = `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rId4"/></a:ext></dgm:extLst></dgm:dataModel>`;
    const relIds = (dm: string) => `<dgm:relIds xmlns:dgm="${DGM}" r:dm="${dm}" r:lo="rId9" r:qs="rId9" r:cs="rId9"/>`;
    const shapes =
      frame(2, 'Process', DGM, relIds('rId3')) +
      frame(3, 'Lost diagram', DGM, relIds('rId8')) +
      frame(4, 'Worksheet', 'http://schemas.openxmlformats.org/presentationml/2006/ole', `<p:oleObj progId="Excel.Sheet.12" r:id="rId9">${picture(9, 'Preview', 'rId5')}</p:oleObj>`) +
      frame(5, 'Unknown', 'urn:example:widget', '');
    const bytes = buildPptx({
      slides: [
        {
          layout: 2,
          shapes,
          rels: [
            { id: 'rId3', type: 'diagramData', target: '../diagrams/data1.xml' },
            { id: 'rId4', type: 'diagramDrawing', target: '../diagrams/drawing1.xml' },
            { id: 'rId5', type: 'image', target: '../media/preview.png' },
          ],
        },
      ],
      parts: [
        { name: 'ppt/diagrams/data1.xml', data: dataModel },
        { name: 'ppt/diagrams/drawing1.xml', data: drawing },
        { name: 'ppt/media/preview.png', data: encodePng(4, 4, () => [0, 0, 255]) },
      ],
    });
    const slide = (await renderAll(bytes)).slides[0]?.element ?? assert.fail();
    const step = byName(byName(slide, 'Process'), 'Step');
    assert.deepEqual([css(step, 'left'), css(step, 'width')], ['100px', '100px']);
    assert.equal(step.querySelector('path')?.getAttribute('fill'), '#70ad47');
    assert.deepEqual([css(step.querySelector('.pptx-text'), 'top'), step.textContent], ['29.8px', 'Plan'], 'text goes into the stored text box (25 px down, plus the top inset)');
    assert.equal(byName(slide, 'Lost diagram').textContent, 'SmartArt graphic');
    assert.match(byName(slide, 'Worksheet').querySelector('pattern image')?.getAttribute('href') ?? '', /^blob:/, 'an embedded object shows its preview picture');
    assert.equal(byName(slide, 'Unknown').textContent, 'Embedded content');
  });

  it('labels videos, unsupported pictures and ink', async () => {
    const shapes =
      picture(2, 'Clip', 'rId2', '<a:videoFile r:link="rId3"/>') +
      picture(3, 'Vector', 'rId4') +
      `<p:contentPart r:id="rId5"><p:nvContentPartPr><p:cNvPr id="4" name="Scribble"/><p:cNvContentPartPr/><p:nvPr/></p:nvContentPartPr>${xfrm([0, 0, 952500, 952500], {}, 'p:xfrm')}</p:contentPart>`;
    const bytes = buildPptx({
      slides: [
        {
          layout: 2,
          shapes,
          rels: [
            { id: 'rId2', type: 'image', target: '../media/poster.png' },
            { id: 'rId3', type: 'video', target: 'https://example.com/clip.mp4', external: true },
            { id: 'rId4', type: 'image', target: '../media/vector.emf' },
          ],
        },
      ],
      parts: [
        { name: 'ppt/media/poster.png', data: encodePng(2, 2, () => [0, 0, 0]) },
        { name: 'ppt/media/vector.emf', data: emf },
      ],
    });
    const slide = (await renderAll(bytes)).slides[0]?.element ?? assert.fail();
    const clip = byName(slide, 'Clip');
    assert.equal(clip.querySelector('.pptx-media-badge')?.textContent, 'Video');
    assert.ok(clip.querySelector('pattern image'), 'the poster frame is drawn');
    assert.equal(byName(slide, 'Vector').querySelector('.pptx-placeholder')?.textContent, 'EMF picture');
    assert.equal(byName(slide, 'Scribble').textContent, 'Ink');
    assert.equal(slide.querySelector('video, audio, iframe, object, embed'), null);
  });
});

describe('text inheritance outside placeholders', () => {
  before(() => installDom());

  it('gives text boxes the presentation’s default text style over the master’s otherStyle', async () => {
    const box = shape({ id: 2, name: 'Note', textBox: true, xfrm: [0, 0, 2000000, 500000], geometry: 'rect', paragraphs: paragraph(run('Small print')) });
    const bytes = buildPptx({
      slides: [{ layout: 2, shapes: box }],
      defaultRunProperties: `<a:defRPr sz="900">${solid(srgb('FF0000'))}</a:defRPr>`,
    });
    const span = byName((await renderAll(bytes)).slides[0]?.element ?? assert.fail(), 'Note').querySelector('span');
    assert.equal(css(span, 'font-size'), '12px');
    assert.equal(css(span, 'color'), '#ff0000');
  });

  it('leaves out shapes hidden in the selection pane', async () => {
    const shapes = shape({ id: 2, name: 'Shown', xfrm: [0, 0, 952500, 952500], geometry: 'rect' }) + shape({ id: 3, name: 'Draft note', hidden: true, xfrm: [0, 0, 952500, 952500], geometry: 'rect', paragraphs: paragraph(run('secret')) });
    const slide = (await renderAll(buildPptx({ slides: [{ layout: 2, shapes }] }))).slides[0]?.element ?? assert.fail();
    assert.ok(slide.querySelector('[data-name="Shown"]'));
    assert.equal(slide.querySelector('[data-name="Draft note"]'), null);
    assert.equal(slide.textContent?.includes('secret'), false);
  });

  it('hides the master’s shapes on slides that ask for it', async () => {
    const bytes = buildPptx({ slides: [{ layout: 2, shapes: '', attributes: ' showMasterSp="0"' }] });
    const slide = (await renderAll(bytes)).slides[0]?.element ?? assert.fail();
    assert.equal(slide.querySelector('[data-name="Master accent bar"]'), null);
  });
});
