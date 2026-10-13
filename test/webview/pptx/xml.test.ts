import assert from 'node:assert/strict';
import { before, describe, it } from 'node:test';
import { MAX_XML_PART_SIZE } from '../../../src/limits.ts';
import { relsPartName, resolvePartName } from '../../../src/webview/pptx/package.ts';
import { bool, child, children, decodeXml, find, kids, num, parseXml, relAttr } from '../../../src/webview/pptx/xml.ts';

describe('package parts as XML', () => {
  before(async () => (await import('./dom.ts')).installDom());

  it('parses a part into its root element', () => {
    const root = parseXml('<?xml version="1.0"?><p:sld xmlns:p="urn:p" show="0"><p:cSld/></p:sld>', 'slide1.xml');
    assert.equal(root.localName, 'sld');
    assert.equal(bool(root, 'show'), false);
    assert.equal(child(root, 'cSld')?.localName, 'cSld');
  });

  it('refuses DOCTYPEs (entity expansion), malformed XML and oversized parts', () => {
    const laughs = '<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;&lol;">]><r>&lol2;</r>';
    assert.throws(() => parseXml(laughs, 'ppt/slides/slide1.xml'), /ppt\/slides\/slide1\.xml declares a DOCTYPE/);
    assert.throws(() => parseXml('<r><!doctype x></r>', 'a.xml'), /declares a DOCTYPE/);
    assert.throws(() => parseXml('<a><b></a>', 'broken.xml'), /broken\.xml is not well-formed XML/);
    assert.throws(() => parseXml(' '.repeat(MAX_XML_PART_SIZE + 1), 'huge.xml'), /huge\.xml is too large to parse \(limit 64 MB\)/);
  });

  it('resolves Markup Compatibility blocks to their fallback', () => {
    const root = parseXml(
      '<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><a/><mc:AlternateContent><mc:Choice Requires="x"><choice/></mc:Choice><mc:Fallback><fallback/></mc:Fallback></mc:AlternateContent>' +
        '<mc:AlternateContent><mc:Choice Requires="y"><only/></mc:Choice></mc:AlternateContent></r>',
      'mc.xml',
    );
    assert.deepEqual(kids(root).map((node) => node.localName), ['a', 'fallback', 'only']);
    assert.equal(children(root, 'choice').length, 0);
    assert.equal(find(root, 'fallback')?.localName, 'fallback');
    assert.equal(find(root, 'a', 'missing'), undefined);
  });

  it('reads numbers, percentages, booleans and relationship ids', () => {
    const root = parseXml('<r xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" a="12" p="50%" n="x" t="true" f="0" r:embed="rId7"/>', 'attrs.xml');
    assert.equal(num(root, 'a'), 12);
    assert.equal(num(root, 'p'), 50000);
    assert.equal(num(root, 'n'), undefined);
    assert.equal(num(root, 'missing'), undefined);
    assert.equal(bool(root, 't'), true);
    assert.equal(bool(root, 'f'), false);
    assert.equal(bool(root, 'n'), undefined);
    assert.equal(relAttr(root, 'embed'), 'rId7');
    assert.equal(relAttr(root, 'link'), undefined);
  });

  it('decodes UTF-8 and UTF-16 parts', () => {
    assert.equal(decodeXml(new TextEncoder().encode('<a>é</a>')), '<a>é</a>');
    const utf16 = new Uint8Array([0xff, 0xfe, ...[...'<a/>'].flatMap((c) => [c.charCodeAt(0), 0])]);
    assert.equal(decodeXml(utf16), '<a/>');
  });
});

describe('relationship targets', () => {
  it('resolves relative, absolute and encoded targets against the source part', () => {
    assert.equal(resolvePartName('ppt/slides/slide1.xml', '../media/image1.png'), 'ppt/media/image1.png');
    assert.equal(resolvePartName('ppt/slides/slide1.xml', '/ppt/media/a.png'), 'ppt/media/a.png');
    assert.equal(resolvePartName('ppt/presentation.xml', 'slides/./slide2.xml'), 'ppt/slides/slide2.xml');
    assert.equal(resolvePartName('ppt/slides/slide1.xml', '../media/my%20picture.png#frag'), 'ppt/media/my picture.png');
    assert.equal(resolvePartName('ppt/slides/slide1.xml', '../../../../etc/passwd'), 'etc/passwd', 'cannot climb above the package root');
    assert.equal(resolvePartName('', 'ppt/presentation.xml'), 'ppt/presentation.xml');
    assert.equal(relsPartName('ppt/slides/slide1.xml'), 'ppt/slides/_rels/slide1.xml.rels');
    assert.equal(relsPartName(''), '_rels/.rels');
  });
});
