import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { describe, it } from 'node:test';
import { ForgedSizeError, inflateEntry, readZipDirectory, readZipText, resolvePath, ZipFormatError } from '../../../src/webview/docx/zipDirectory.ts';
import { samples } from '../../fixtures/docx.ts';
import { patchEntry, prependedDecoy, withUnicodePath } from '../../support/zipTamper.ts';
import { writeZip } from '../../support/zipWriter.ts';

// The JSZip docx-preview reads files with.
const JSZip = createRequire(import.meta.url)('jszip') as { loadAsync(data: Uint8Array): Promise<{ files: Record<string, unknown> }> };

const jszipNames = async (zip: Uint8Array) => Object.keys((await JSZip.loadAsync(zip)).files).sort();
const ourNames = (zip: Uint8Array) => readZipDirectory(zip).map((entry) => entry.name).sort();

describe('docx ZIP directory', () => {
  it('lists the entries JSZip lists, with declared sizes', async () => {
    const zip = writeZip([
      { name: 'word/document.xml', data: '<w:document/>' },
      { name: 'word/media/画像.png', data: Buffer.alloc(300), store: true },
      { name: 'word/media/', data: '', store: true },
    ], 'archive comment');
    const entries = readZipDirectory(zip);
    assert.deepEqual(
      entries.map(({ name, size, method, isDirectory }) => ({ name, size, method, isDirectory })),
      [
        { name: 'word/document.xml', size: 13, method: 8, isDirectory: false },
        { name: 'word/media/画像.png', size: 300, method: 0, isDirectory: false },
        { name: 'word/media/', size: 0, method: 0, isDirectory: true },
      ],
    );
    assert.deepEqual(ourNames(zip), await jszipNames(zip));
  });

  it('lists the same entries as JSZip for every sample', async () => {
    for (const sample of samples) {
      const zip = new Uint8Array(sample.build());
      assert.deepEqual(ourNames(zip), await jszipNames(zip), sample.fileName);
    }
  });

  it('refuses names JSZip would rewrite, so a decoy cannot hide the part docx-preview reads', async () => {
    for (const name of ['x/../word/document.xml', './word/document.xml', 'word//document.xml', 'word\\document.xml']) {
      const zip = writeZip([{ name, data: '<a/>' }]);
      assert.throws(() => readZipDirectory(zip), ZipFormatError, name);
    }
    const decoy = writeZip([{ name: 'word/document.xml', data: '<small/>' }, { name: 'x/../word/document.xml', data: '<hidden/>' }]);
    assert.deepEqual(await jszipNames(decoy), ['word/document.xml'], 'JSZip keeps only the second one, under the first one’s name');
    assert.throws(() => readZipDirectory(decoy), /unusual entry name/);
  });

  it('refuses a Unicode path field JSZip would take the name from', async () => {
    const zip = withUnicodePath(writeZip([{ name: 'decoy.bin', data: '<a/>' }]), 'decoy.bin', 'word/document.xml');
    assert.deepEqual(await jszipNames(zip), ['word/document.xml']);
    assert.throws(() => readZipDirectory(zip), /has a second, Unicode name/);
  });

  it('refuses invalid UTF-8 names, duplicates and names that differ between the two headers', () => {
    const invalid = patchEntry(writeZip([{ name: 'word/é.xml', data: 'x' }]), 'word/é.xml', (copy, record) => {
      copy.writeUInt8(0xe9, record.offset + 46 + 5);
      copy.writeUInt8(0x41, record.offset + 46 + 6);
    });
    assert.throws(() => readZipDirectory(invalid), /unusual entry name/);
    assert.throws(() => readZipDirectory(writeZip([{ name: 'a.xml', data: '1' }, { name: 'a.xml', data: '2' }])), /appears twice/);
    const renamed = patchEntry(writeZip([{ name: 'a.xml', data: '1' }]), 'a.xml', (copy, record) => copy.write('b', record.localHeader + 30));
    assert.throws(() => readZipDirectory(renamed), /named differently in its local header/);
  });

  it('refuses a directory attribute on a name without a trailing slash (JSZip would add one)', () => {
    const zip = patchEntry(writeZip([{ name: 'word/document.xml', data: '<a/>' }]), 'word/document.xml', (copy, record) => copy.writeUInt32LE(0x10, record.offset + 38));
    assert.throws(() => readZipDirectory(zip), /unusual entry name/);
  });

  it('refuses data before the central directory or after the end record, so both readers see one directory', async () => {
    const decoy = writeZip([{ name: 'word/document.xml', data: 'decoy' }]);
    const real = writeZip([{ name: 'word/document.xml', data: 'r'.repeat(10_000), declaredSize: 100 }]);
    const both = prependedDecoy(decoy, real);
    // JSZip shifts its offsets past the gap and reads the real entry (declared 100 bytes), not the decoy (5 bytes).
    assert.equal(((await JSZip.loadAsync(both)).files['word/document.xml'] as { _data: { uncompressedSize: number } })._data.uncompressedSize, 100);
    assert.throws(() => readZipDirectory(both), /unexpected data around the ZIP central directory/);
    assert.throws(() => readZipDirectory(Buffer.concat([decoy, Buffer.from('trailing')])), /unexpected data around the ZIP central directory/);
  });

  it('refuses ZIP64, encryption, other compression methods and stray central directory records', () => {
    const zip = writeZip([{ name: 'a.xml', data: '<a/>' }, { name: 'b.xml', data: '<b/>' }]);
    const eocd = zip.length - 22;
    const edited = (edit: (copy: Buffer) => void) => {
      const copy = Buffer.from(zip);
      edit(copy);
      return copy;
    };
    assert.throws(() => readZipDirectory(edited((copy) => copy.writeUInt16LE(0xffff, eocd + 10))), /ZIP64/);
    const oneRecord = edited((copy) => {
      copy.writeUInt16LE(1, eocd + 8);
      copy.writeUInt16LE(1, eocd + 10);
    });
    assert.throws(() => readZipDirectory(oneRecord), /unexpected data in the ZIP central directory/);
    assert.throws(() => readZipDirectory(patchEntry(zip, 'a.xml', (copy, record) => copy.writeUInt16LE(1, record.offset + 8))), /encrypted/);
    assert.throws(() => readZipDirectory(patchEntry(zip, 'a.xml', (copy, record) => copy.writeUInt16LE(12, record.offset + 10))), /compression method 12/);
    assert.throws(() => readZipDirectory(new TextEncoder().encode('not a zip file')), /no ZIP end of central directory/);
  });

  it('resolves paths like JSZip', () => {
    assert.equal(resolvePath('x/../word/document.xml'), 'word/document.xml');
    assert.equal(resolvePath('./word//a.xml'), 'word/a.xml');
    assert.equal(resolvePath('word/media/'), 'word/media/');
    assert.equal(resolvePath('/word/a.xml'), '/word/a.xml');
  });

  it('reads stored and deflated entries as text', async () => {
    const zip = writeZip([
      { name: 'stored.xml', data: '<a>日本語</a>', store: true },
      { name: 'deflated.xml', data: `<b>${'x'.repeat(100_000)}</b>` },
    ]);
    const [stored, deflated] = readZipDirectory(zip);
    assert.ok(stored && deflated);
    assert.equal(await readZipText(zip, stored), '<a>日本語</a>');
    assert.equal((await readZipText(zip, deflated)).length, 100_007);
  });

  it('holds an entry to its declared size', async () => {
    const zip = writeZip([
      { name: 'stored.bin', data: Buffer.alloc(1000), store: true, declaredSize: 10 },
      { name: 'deflated.bin', data: Buffer.alloc(1_000_000), declaredSize: 1000 },
      { name: 'smaller.bin', data: Buffer.alloc(1000), declaredSize: 5000 },
      { name: 'honest.bin', data: Buffer.alloc(1_000_000) },
    ]);
    const [stored, deflated, smaller, honest] = readZipDirectory(zip);
    assert.ok(stored && deflated && smaller && honest);
    await assert.rejects(inflateEntry(zip, stored), ForgedSizeError);
    let seen = 0;
    await assert.rejects(
      inflateEntry(zip, deflated, (chunk) => {
        seen += chunk.byteLength;
      }),
      ForgedSizeError,
    );
    assert.ok(seen <= 1000, `stopped after ${seen} bytes`);
    await assert.rejects(inflateEntry(zip, smaller), /smaller than its ZIP header says/);
    await inflateEntry(zip, honest);
  });
});
