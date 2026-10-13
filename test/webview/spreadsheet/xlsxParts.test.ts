// The pieces the .xlsx reader is built on: the streaming XML scanner, the ZIP reader and relationship parts.
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { parseRelationships, relationshipsPartOf, resolvePartName } from '../../../src/webview/spreadsheet/xlsx/relationships.ts';
import { decodeEntities, decodeOoxmlEscapes, scanXml, XmlScanner, type XmlHandler } from '../../../src/webview/spreadsheet/xlsx/xml.ts';
import { ZipReader } from '../../../src/webview/spreadsheet/xlsx/zip.ts';
import { writeZip } from '../../support/zipWriter.ts';

/** Every event the scanner reports, as strings. */
function events(chunks: string[], wantText = true): string[] {
  const seen: string[] = [];
  const handler: XmlHandler = {
    open(tag) {
      seen.push(`<${tag.name}${tag.selfClosing ? '/' : ''} r=${tag.attr('r') ?? '-'} id=${tag.attr('id') ?? '-'}>`);
    },
    close(name) {
      seen.push(`</${name}>`);
    },
    text(text) {
      seen.push(`"${text}"`);
    },
  };
  const scanner = new XmlScanner(handler);
  scanner.wantText = wantText;
  for (const chunk of chunks) {
    scanner.write(chunk);
  }
  scanner.end();
  return seen;
}

async function collect(chunks: AsyncIterable<string>): Promise<string> {
  let text = '';
  for await (const chunk of chunks) {
    text += chunk;
  }
  return text;
}

describe('XML scanner', () => {
  const xml = '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "boom">]><x:c r="A1" xmlns:x="urn:x"><v>1 &lt; 2 &amp;a;</v><!-- <c r="B1"/> --><x:is r:id="rId7" t=\'a>b\'/><t><![CDATA[<raw>]]></t></x:c>';

  it('reports elements without prefixes, attributes, decoded text and CDATA; skips comments and DTDs', () => {
    assert.deepEqual(events([xml]), [
      '<c r=A1 id=->',
      '<v r=- id=->',
      '"1 < 2 &a;"',
      '</v>',
      '<is/ r=- id=rId7>',
      '</is>',
      '<t r=- id=->',
      '"<raw>"',
      '</t>',
      '</c>',
    ]);
  });

  it('gives the same events however the input is split into chunks', () => {
    const whole = events([xml]);
    for (let size = 1; size < 12; size++) {
      const chunks: string[] = [];
      for (let i = 0; i < xml.length; i += size) {
        chunks.push(xml.slice(i, i + size));
      }
      assert.deepEqual(events(chunks).join('').replaceAll('""', ''), whole.join(''), `chunks of ${size}`);
    }
  });

  it('skips text unless asked for it', () => {
    assert.ok(!events([xml], false).some((event) => event.startsWith('"')));
  });

  it('refuses a single token larger than the cap instead of buffering it forever', () => {
    const scanner = new XmlScanner({});
    scanner.wantText = true;
    assert.throws(() => {
      for (let i = 0; i < 20; i++) {
        scanner.write('x'.repeat(1024 * 1024));
      }
    }, /too large/);
  });

  it('decodes entities and OOXML escapes, leaving unknown ones alone', () => {
    assert.equal(decodeEntities('&#x41;&#66;&quot;&apos;&gt;&unknown;&#0;&#xD800;'), 'AB"\'>&unknown;��');
    assert.equal(decodeOoxmlEscapes('a_x000D_b_x005F_c'), 'a\rb_c');
  });

  it('scans a whole string with scanXml', () => {
    const names: string[] = [];
    scanXml('<a><b/></a>', { open: (tag) => names.push(tag.name) });
    assert.deepEqual(names, ['a', 'b']);
  });
});

describe('relationships', () => {
  it('resolves targets against the source part and keeps external ones', () => {
    const rels = parseRelationships(
      '<Relationships><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet%201.xml"/>' +
        '<Relationship Id="rId2" Type="x/hyperlink" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/>' +
        '<Relationship Id="rId3" Type="x/styles" Target="/xl/./styles.xml"/><Relationship Id="rId4" Type="x/theme" Target="../theme/theme1.xml"/></Relationships>',
      'xl/workbook.xml',
    );
    assert.deepEqual(rels.get('rId1'), { target: 'xl/worksheets/sheet 1.xml', type: 'worksheet', external: false });
    assert.deepEqual(rels.get('rId2'), { target: 'https://example.com/?a=1&b=2', type: 'hyperlink', external: true });
    assert.equal(rels.get('rId3')?.target, 'xl/styles.xml');
    assert.equal(rels.get('rId4')?.target, 'theme/theme1.xml');
    assert.equal(relationshipsPartOf('xl/worksheets/sheet1.xml'), 'xl/worksheets/_rels/sheet1.xml.rels');
    assert.equal(resolvePartName('xl/', '../../../etc/passwd'), 'etc/passwd');
  });
});

describe('ZipReader in the webview', () => {
  const archive = writeZip([
    { name: 'xl/workbook.xml', data: '<workbook>日本語</workbook>' },
    { name: 'stored.txt', data: 'plain', store: true },
    { name: 'big.xml', data: '<r>' + 'x'.repeat(300_000) + '</r>' },
  ]);

  it('reads stored and deflated entries, matching names without case', async () => {
    const zip = ZipReader.open(archive);
    assert.equal(await zip.readText('XL/Workbook.xml'), '<workbook>日本語</workbook>');
    assert.equal(await zip.readText('stored.txt'), 'plain');
    assert.equal((await collect(zip.textChunks('big.xml'))).length, 300_007);
    assert.equal(zip.has('missing.xml'), false);
    await assert.rejects(zip.readText('missing.xml'), /not found/);
  });

  it('refuses parts above the size given for reading them whole', async () => {
    const zip = ZipReader.open(archive);
    await assert.rejects(zip.readText('big.xml', 1000), /too large to read/);
  });

  it('stops a deflate stream that inflates beyond its declared size', async () => {
    const forged = writeZip([{ name: 'bomb.xml', data: 'a'.repeat(200_000), declaredSize: 1000 }]);
    await assert.rejects(ZipReader.open(forged).readText('bomb.xml'), /inflates beyond its declared size/);
  });

  it('reports truncated, damaged and size-mismatched entries', async () => {
    const short = writeZip([{ name: 'a.xml', data: 'a'.repeat(1000), declaredSize: 5000 }]);
    await assert.rejects(ZipReader.open(short).readText('a.xml'), /size mismatch/);
    const badHeader = Buffer.from(archive);
    badHeader[0] = 0;
    await assert.rejects(ZipReader.open(badHeader).readText('xl/workbook.xml'), /Corrupt ZIP local header/);
    // Bytes that are not deflate data, labelled as deflated in both headers.
    const notDeflate = writeZip([{ name: 'a.xml', data: Buffer.alloc(16, 0xff), store: true }]);
    notDeflate.writeUInt16LE(8, 8);
    notDeflate.writeUInt16LE(8, notDeflate.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02])) + 10);
    await assert.rejects(ZipReader.open(notDeflate).readText('a.xml'), /Corrupt ZIP entry \(bad compressed data\)/);
    assert.throws(() => ZipReader.open(Buffer.from('not a zip')), /Not a ZIP archive/);
  });

  it('applies the archive limits again', () => {
    const tooMany = writeZip(Array.from({ length: 5 }, (_value, i) => ({ name: `${i}.xml`, data: 'x' })));
    assert.throws(() => ZipReader.open(tooMany, { maxEntries: 4, maxEntrySize: 1e9, maxTotalSize: 1e9, maxCompressionRatio: 1100 }), /too many entries/);
    const huge = writeZip([{ name: 'a.xml', data: 'x', declaredSize: 300 * 1024 * 1024 }]);
    assert.throws(() => ZipReader.open(huge), /ZIP entry too large/);
    const ratio = writeZip([{ name: 'a.xml', data: 'x'.repeat(100), declaredSize: 50 * 1024 * 1024 }]);
    assert.throws(() => ZipReader.open(ratio), /compressed suspiciously well/);
  });
});
