import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { checkCompoundFile, checkOoxmlPackage, checkPdf, officeDocumentTarget } from '../src/preflight.ts';
import { compoundFileHeaderStub } from './support/compoundFile.ts';
import { OFFICE_DOCUMENT_RELATIONSHIP, ooxmlPackage } from './support/ooxml.ts';
import { textPdf } from './support/pdfWriter.ts';
import { writeZip } from './support/zipWriter.ts';

const docxPackage = () => ooxmlPackage({ name: 'word/document.xml', xml: '<w:document/>' });

describe('checkPdf', () => {
  it('accepts a header at the start or after up to 1 KB of junk', () => {
    checkPdf(textPdf([['x']]));
    checkPdf(Buffer.concat([Buffer.alloc(1000, 0x20), textPdf([['x']])]));
  });

  it('refuses files without a PDF header near the start', () => {
    assert.throws(() => checkPdf(Buffer.from('hello')), /Not a \.pdf file \(no %PDF- header\)/);
    assert.throws(() => checkPdf(Buffer.concat([Buffer.alloc(2000, 0x20), textPdf([['x']])])), /no %PDF- header/);
    assert.throws(() => checkPdf(new Uint8Array()), /no %PDF- header/);
  });

  it('reads only the given view of a larger buffer', () => {
    const pdf = textPdf([['x']]);
    const shared = Buffer.concat([Buffer.from('junk'.repeat(400)), pdf]);
    checkPdf(shared.subarray(1600));
    assert.throws(() => checkPdf(shared.subarray(0, 1200)), /no %PDF- header/);
  });
});

describe('checkCompoundFile', () => {
  it('accepts the compound file signature', () => {
    checkCompoundFile(compoundFileHeaderStub(), '.doc');
  });

  it('explains a ZIP package under a legacy extension', () => {
    assert.throws(() => checkCompoundFile(docxPackage(), '.doc'), /This \.doc file is a ZIP package, probably a \.docx file/);
  });

  it('refuses other files', () => {
    assert.throws(() => checkCompoundFile(Buffer.from('{\\rtf1 hello}'), '.doc'), /Not a \.doc file \(no compound file header\)/);
  });
});

describe('checkOoxmlPackage', () => {
  it('accepts a package whose relationship points at an existing main part', () => {
    checkOoxmlPackage(docxPackage(), '.docx');
  });

  it('names the right extension when the package is another OOXML format', () => {
    const xlsx = ooxmlPackage({ name: 'xl/workbook.xml', xml: '<workbook/>' });
    assert.throws(() => checkOoxmlPackage(xlsx, '.docx'), /This \.docx file is actually a \.xlsx file/);
  });

  it('recognizes password-protected packages and legacy binaries', () => {
    const encrypted = Buffer.concat([compoundFileHeaderStub(), Buffer.from('EncryptedPackage', 'utf16le')]);
    assert.throws(() => checkOoxmlPackage(encrypted, '.xlsx'), /This \.xlsx file is password-protected/);
    assert.throws(() => checkOoxmlPackage(compoundFileHeaderStub(), '.pptx'), /older binary file, probably a \.ppt file/);
  });

  it('refuses archives without package parts or main part', () => {
    assert.throws(() => checkOoxmlPackage(writeZip([{ name: 'a.txt', data: 'a' }]), '.docx'), /package parts are missing/);
    const dangling = writeZip([
      { name: '[Content_Types].xml', data: '<Types/>' },
      { name: '_rels/.rels', data: `<Relationships><Relationship Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/document.xml"/></Relationships>` },
    ]);
    assert.throws(() => checkOoxmlPackage(dangling, '.docx'), /main document part is missing/);
  });

  it('reports broken and oversized archives as invalid files', () => {
    assert.throws(() => checkOoxmlPackage(Buffer.from('plain text'), '.docx'), /Not a valid \.docx file: Not a ZIP archive/);
    const forged = writeZip([{ name: '[Content_Types].xml', data: 'x'.repeat(10), declaredSize: 300 * 1024 * 1024 }]);
    assert.throws(() => checkOoxmlPackage(forged, '.docx'), /Not a valid \.docx file: ZIP entry too large/);
  });
});

describe('officeDocumentTarget', () => {
  it('finds the main part among other relationships, with either quote style and a leading slash', () => {
    const rels =
      `<Relationships><Relationship Id="rId2" Type="http://x/core-properties" Target="docProps/core.xml"/>` +
      `<Relationship Target='/ppt/presentation.xml' Id='rId1' Type='http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument'/></Relationships>`;
    assert.equal(officeDocumentTarget(rels), 'ppt/presentation.xml');
  });

  it('returns undefined without an officeDocument relationship', () => {
    assert.equal(officeDocumentTarget('<Relationships/>'), undefined);
  });
});
