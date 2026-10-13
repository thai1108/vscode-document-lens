import assert from 'node:assert/strict';
import { before, describe, it } from 'node:test';
import { Deck, matchPlaceholder, placeholderOf } from '../../../src/webview/pptx/deck.ts';
import { TextStyleResolver } from '../../../src/webview/pptx/textStyle.ts';
import { child, find } from '../../../src/webview/pptx/xml.ts';
import { sampleDeck } from '../../fixtures/pptx.ts';
import { writeZip } from '../../support/zipWriter.ts';
import { installDom } from './dom.ts';

let deck: Deck;

function shapeNamed(root: Element, name: string): Element {
  const shape = [...(find(root, 'cSld', 'spTree')?.children ?? [])].find((node) => node.querySelector('cNvPr')?.getAttribute('name') === name || find(node, 'nvSpPr', 'cNvPr')?.getAttribute('name') === name);
  return shape ?? assert.fail(`no shape ${name}`);
}

describe('deck structure', () => {
  before(async () => {
    installDom();
    deck = await Deck.open(new Uint8Array(sampleDeck()));
  });

  it('reads the slide size, order and numbering', () => {
    assert.equal(deck.width, 12192000);
    assert.equal(deck.height, 6858000);
    assert.deepEqual(deck.slideParts, ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml', 'ppt/slides/slide4.xml']);
    assert.equal(deck.firstSlideNumber, 1);
    assert.equal(deck.omittedSlides, 0);
    assert.ok(deck.defaultTextStyle);
  });

  it('finds hidden slides from their root tag and their parsed part alike', async () => {
    assert.deepEqual(await Promise.all([0, 1, 2, 3].map((index) => deck.isHidden(index))), [false, false, false, true]);
    assert.equal((await deck.slide(3)).hidden, true);
    assert.equal((await deck.slide(0)).hidden, false);
  });

  it('links slides to their layout, master, theme and color map', async () => {
    const slide = await deck.slide(1);
    assert.equal(slide.layout?.part, 'ppt/slideLayouts/slideLayout2.xml');
    assert.equal(slide.master?.part, 'ppt/slideMasters/slideMaster1.xml');
    assert.equal(slide.layout?.master, slide.master, 'masters are loaded once');
    assert.equal(slide.theme.colors['accent1'], '4472C4');
    assert.equal(slide.theme.minor.latin, 'Calibri');
    assert.equal(slide.theme.minor.scripts['Jpan'], '游ゴシック');
    assert.equal(slide.colorMap['tx1'], 'dk1');
    assert.equal(slide.master?.placeholders.length, 5);
    await assert.rejects(deck.slide(9), /No slide 10/);
  });

  it('matches slide placeholders to layout placeholders by index and to master ones by kind', async () => {
    const slide = await deck.slide(0);
    const layout = slide.layout ?? assert.fail();
    const master = slide.master ?? assert.fail();
    const title = placeholderOf(shapeNamed(slide.root, 'Title 1')) ?? assert.fail();
    assert.equal(matchPlaceholder(layout.placeholders, title, true), shapeNamed(layout.root, 'Title 1'));
    assert.equal(matchPlaceholder(master.placeholders, title, false), shapeNamed(master.root, 'Title Placeholder 1'));
    const subtitle = placeholderOf(shapeNamed(slide.root, 'Subtitle 2')) ?? assert.fail();
    assert.equal(matchPlaceholder(layout.placeholders, subtitle, true), shapeNamed(layout.root, 'Subtitle 2'));
    assert.equal(matchPlaceholder(master.placeholders, subtitle, false), shapeNamed(master.root, 'Text Placeholder 2'));
    const number = placeholderOf(shapeNamed(slide.root, 'Slide Number')) ?? assert.fail();
    assert.equal(matchPlaceholder(master.placeholders, number, false), shapeNamed(master.root, 'sldNum placeholder'));
  });

  it('resolves text properties through the master styles and placeholder list styles', async () => {
    const slide = await deck.slide(1);
    const master = slide.master ?? assert.fail();
    const body = shapeNamed(slide.root, 'Content Placeholder 2');
    const layoutBody = shapeNamed(slide.layout?.root ?? assert.fail(), 'Content Placeholder 2');
    const lstStyle = (shape: Element) => child(child(shape, 'txBody'), 'lstStyle');
    const styles = new TextStyleResolver({
      base: [deck.defaultTextStyle, master.textStyles.body],
      shape: [lstStyle(shapeNamed(master.root, 'Text Placeholder 2')), lstStyle(layoutBody), lstStyle(body)],
    });
    const first = styles.paragraph(undefined);
    assert.equal(first.marL, 228600);
    assert.equal(first.indent, -228600);
    assert.deepEqual([first.bullet.kind, first.bullet.char], ['char', '•']);
    assert.deepEqual(first.lineSpacing, { pct: 0.9 });
    assert.deepEqual(first.spaceBefore, { pts: 10 });
    assert.equal(styles.run(0, undefined, undefined).size, 28);
    assert.equal(styles.run(0, undefined, undefined).latin, '+mn-lt');
    assert.equal(styles.run(1, undefined, undefined).size, 24);
    const titleStyles = new TextStyleResolver({ base: [deck.defaultTextStyle, master.textStyles.title], shape: [] });
    assert.equal(titleStyles.run(0, undefined, undefined).size, 44);
    assert.equal(titleStyles.run(0, undefined, undefined).latin, '+mj-lt');
    assert.equal(titleStyles.paragraph(undefined).bullet.kind, 'none');
  });
});

describe('broken packages', () => {
  before(() => installDom());

  it('needs a presentation part', async () => {
    const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>';
    await assert.rejects(Deck.open(new Uint8Array(writeZip([{ name: '_rels/.rels', data: rels }]))), /The presentation part is missing/);
    await assert.rejects(
      Deck.open(new Uint8Array(writeZip([{ name: '_rels/.rels', data: rels }, { name: 'ppt/presentation.xml', data: '<w:document xmlns:w="urn:w"/>' }]))),
      /The main part is not a presentation/,
    );
  });

  it('skips slides whose parts are missing and defaults a missing slide size', async () => {
    const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>';
    const presentationRels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>';
    const presentation = '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId9"/></p:sldIdLst></p:presentation>';
    const deck = await Deck.open(new Uint8Array(writeZip([
      { name: '_rels/.rels', data: rels },
      { name: 'ppt/presentation.xml', data: presentation },
      { name: 'ppt/_rels/presentation.xml.rels', data: presentationRels },
    ])));
    assert.deepEqual(deck.slideParts, []);
    assert.deepEqual([deck.width, deck.height], [9144000, 6858000]);
  });
});
