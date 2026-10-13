// checkPackage with happy-dom's DOMParser standing in for the XML fragment parsing it uses in the webview (happy-dom
// parses fragments as HTML); that parsing, and Chrome's stricter inflating, are covered by test/e2e/docx.spec.ts.
import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import { Window } from 'happy-dom';
import { checkPackage, countStartTags, type PackageCheckOptions } from '../../../src/webview/docx/packageCheck.ts';
import { samples } from '../../fixtures/docx.ts';
import { CONTENT_TYPE, paragraph, RELATIONSHIP, run, WORD_NAMESPACES, wordPackage } from '../../support/wordPackage.ts';
import { prependedDecoy } from '../../support/zipTamper.ts';
import { writeZip } from '../../support/zipWriter.ts';

let window: Window;
before(() => {
  window = new Window({ url: 'https://webview.test/' });
});
after(() => window.happyDOM.close());

const parsed: string[] = [];
function options(extra: PackageCheckOptions = {}): PackageCheckOptions {
  return {
    parseXml(name, text) {
      parsed.push(name);
      const document = new window.DOMParser().parseFromString(text, 'application/xml');
      if (document.getElementsByTagName('parsererror').length > 0) {
        throw new Error(`the document is damaged (${name} is not well-formed XML)`);
      }
      return document as unknown as ParentNode;
    },
    ...extra,
  };
}

const header = (xml: string, name = 'word/header1.xml') => ({ name, contentType: CONTENT_TYPE.header, xml, relationship: { id: 'rIdHeader', type: RELATIONSHIP.header } });
const paragraphsXml = (root: string, count: number) => `<w:${root} ${WORD_NAMESPACES}>${paragraph(run('x')).repeat(count)}</w:${root}>`;

describe('docx package check', () => {
  it('accepts the samples, reading every part docx-preview parses, and reports the entry sizes', async () => {
    for (const sample of samples) {
      parsed.length = 0;
      const { sizes } = await checkPackage(new Uint8Array(sample.build()), options());
      assert.ok(parsed.includes('word/document.xml') && parsed.includes('word/_rels/document.xml.rels'), sample.fileName);
      assert.ok(sizes.get('word/document.xml'));
    }
    assert.ok(parsed.includes('word/styles.xml'));
    assert.ok(!parsed.includes('word/chunk.xml'), 'an altChunk is not parsed by docx-preview');
  });

  it('refuses a part over the limit, whatever its name, but not a picture', async () => {
    const big = paragraphsXml('hdr', 200);
    await assert.rejects(
      checkPackage(wordPackage({ body: paragraph(run('x')), parts: [header(big, 'word/media/header.png')] }), options({ maxXmlPartSize: 4096 })),
      /its part word\/media\/header\.png is larger than 0 MB/,
    );
    const picture = { name: 'word/media/big.png', contentType: 'image/png', data: new Uint8Array(100_000), relationship: { id: 'rIdPicture', type: RELATIONSHIP.image } };
    await checkPackage(wordPackage({ body: paragraph(run('x')), media: [picture] }), options({ maxXmlPartSize: 4096 }));
  });

  it('refuses XML parts beyond the total size or element budget before parsing them', async () => {
    const document = wordPackage({ body: paragraph(run('x')), parts: [header(paragraphsXml('hdr', 100))] });
    parsed.length = 0;
    await assert.rejects(checkPackage(document, options({ maxXmlTotalSize: 2000 })), /its XML parts add up to more than/);
    assert.ok(!parsed.includes('word/header1.xml'));
    parsed.length = 0;
    await assert.rejects(checkPackage(document, options({ maxXmlElements: 300 })), /it has more than 300 XML elements/);
    assert.ok(!parsed.includes('word/header1.xml'));
  });

  it('refuses a malformed part', async () => {
    await assert.rejects(
      checkPackage(wordPackage({ body: paragraph(run('x')), parts: [header(`<w:hdr ${WORD_NAMESPACES}><w:p></w:hdr>`)] }), options()),
      /word\/header1\.xml is not well-formed XML/,
    );
  });

  it('refuses archives JSZip would read differently, and forged sizes', async () => {
    const decoy = writeZip([
      { name: '[Content_Types].xml', data: '<Types/>' },
      { name: 'word/document.xml', data: '<w:document/>' },
      { name: 'x/../word/document.xml', data: '<not-closed>' },
    ]);
    await assert.rejects(checkPackage(decoy, options()), /uses ZIP features Document Lens does not read \(unusual entry name "x\/\.\.\/word\/document\.xml"\)/);
    const small = wordPackage({ body: paragraph(run('decoy')) });
    const forged = wordPackage({ body: paragraph(run(' '.repeat(200_000))), extraEntries: [{ name: 'word/media/forged.png', data: Buffer.alloc(10_000), declaredSize: 100 }] });
    await assert.rejects(checkPackage(prependedDecoy(small, forged), options()), /unexpected data around the ZIP central directory/);
    await assert.rejects(checkPackage(forged, options()), /the file is damaged: word\/media\/forged\.png is larger than its ZIP header says/);
  });

  it('counts start tags', () => {
    assert.equal(countStartTags('<?xml version="1.0"?><!-- c --><a><b/><c>t</c></a>'), 3);
  });
});
