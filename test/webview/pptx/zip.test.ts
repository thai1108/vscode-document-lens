// The webview's own ZIP reader must hold the same line as the host's preflight (test/zip-hostile.test.ts), since it
// decides how much memory a file can make the viewer allocate.
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { ZipReader } from '../../../src/webview/pptx/zip.ts';
import { writeZip } from '../../support/zipWriter.ts';

const EOCD_SIZE = 22;
const centralOffset = (zip: Buffer) => zip.readUInt32LE(zip.length - EOCD_SIZE + 16);
const open = (zip: Buffer, limits?: Parameters<typeof ZipReader.open>[1]) => ZipReader.open(new Uint8Array(zip), limits);

describe('webview ZIP reader', () => {
  it('reads stored and deflated entries with UTF-8 names', async () => {
    const zip = open(writeZip([
      { name: 'stored.bin', data: Buffer.from([0, 255, 1]), store: true },
      { name: 'ppt/日本語.xml', data: '<x>テキスト</x>'.repeat(200) },
    ]));
    assert.deepEqual(zip.names(), ['stored.bin', 'ppt/日本語.xml']);
    assert.deepEqual([...(await zip.read('stored.bin'))], [0, 255, 1]);
    assert.equal(new TextDecoder().decode(await zip.read('ppt/日本語.xml')), '<x>テキスト</x>'.repeat(200));
    assert.equal(zip.size('stored.bin'), 3);
    assert.equal(zip.has('missing'), false);
    await assert.rejects(zip.read('missing'), /ZIP entry not found: missing/);
  });

  it('applies the OOXML limits before inflating anything', () => {
    const limits = { maxEntries: 3, maxEntrySize: 1000, maxTotalSize: 1500, maxCompressionRatio: 1100 };
    assert.throws(() => open(writeZip(['a', 'b', 'c', 'd'].map((name) => ({ name, data: name }))), limits), /too many entries \(4; limit 3\)/);
    assert.throws(() => open(writeZip([{ name: 'big', data: 'x'.repeat(1001) }]), limits), /ZIP entry too large: big/);
    assert.throws(() => open(writeZip([{ name: 'a', data: 'x'.repeat(900) }, { name: 'b', data: 'x'.repeat(900) }]), limits), /ZIP archive too large when unpacked/);
    assert.throws(() => open(writeZip([{ name: 'forged', data: 'x'.repeat(100), declaredSize: 2 * 1024 * 1024 }])), /compressed suspiciously well: forged/);
  });

  it('stops a deflate stream at its declared size (ZIP bomb)', async () => {
    const zip = open(writeZip([{ name: 'bomb', data: Buffer.alloc(5_000_000), declaredSize: 10 }]));
    await assert.rejects(zip.read('bomb'), /inflates beyond its declared size\): bomb/);
  });

  it('reports damaged archives and entries', async () => {
    assert.throws(() => open(Buffer.from('not a zip')), /Not a ZIP archive/);
    const zip64 = writeZip([{ name: 'a', data: 'a' }]);
    zip64.writeUInt32LE(0xffffffff, zip64.length - EOCD_SIZE + 16);
    assert.throws(() => open(zip64), /ZIP64 archives are not supported/);
    assert.throws(() => open(writeZip([{ name: 'z', data: 'abc', declaredSize: 0xffffffff }])), /ZIP64/);
    const wrongDirectory = writeZip([{ name: 'a', data: 'a' }]);
    wrongDirectory.writeUInt32LE(0, wrongDirectory.length - EOCD_SIZE + 16);
    assert.throws(() => open(wrongDirectory), /Corrupt ZIP central directory at entry 0/);

    await assert.rejects(open(writeZip([{ name: 'short', data: 'abc', declaredSize: 10 }])).read('short'), /size mismatch\): short/);
    await assert.rejects(open(writeZip([{ name: 'short', data: 'abc', store: true, declaredSize: 10 }])).read('short'), /size mismatch\): short/);
    await assert.rejects(open(writeZip([{ name: 'cut', data: 'abc', store: true, declaredCompressedSize: 1_000_000 }])).read('cut'), /truncated data\): cut/);
    const badHeader = writeZip([{ name: 'a', data: 'a' }]);
    badHeader.writeUInt32LE(0, 0);
    await assert.rejects(open(badHeader).read('a'), /Corrupt ZIP local header: a/);
    const garbage = writeZip([{ name: 'g', data: Buffer.from([1, 2, 3, 4, 5, 6]), store: true }]);
    garbage.writeUInt16LE(8, centralOffset(garbage) + 10);
    await assert.rejects(open(garbage).read('g'), /Corrupt ZIP entry \(bad deflate data\): g|size mismatch/);
  });

  it('refuses compression methods and encryption it cannot read', async () => {
    const method = writeZip([{ name: 'a.txt', data: 'hello', store: true }]);
    method.writeUInt16LE(12, centralOffset(method) + 10);
    await assert.rejects(open(method).read('a.txt'), /Unsupported ZIP compression method 12: a\.txt/);
    const encrypted = writeZip([{ name: 'secret', data: 'hello', store: true }]);
    encrypted.writeUInt16LE(0x0801, centralOffset(encrypted) + 8);
    await assert.rejects(open(encrypted).read('secret'), /Encrypted ZIP entry: secret/);
  });
});
