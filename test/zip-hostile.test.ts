import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { ZipArchive } from '../src/zip.ts';
import { writeZip } from './support/zipWriter.ts';

const EOCD_SIZE = 22;

function sample(): Buffer {
  return writeZip([{ name: 'a.txt', data: 'hello', store: true }]);
}

function centralOffset(zip: Buffer): number {
  return zip.readUInt32LE(zip.length - EOCD_SIZE + 16);
}

describe('ZipArchive with broken archives', () => {
  it('refuses ZIP64 archives', () => {
    const zip = sample();
    zip.writeUInt32LE(0xffffffff, zip.length - EOCD_SIZE + 16);
    assert.throws(() => ZipArchive.open(zip), /ZIP64 archives are not supported/);
  });

  it('reports a central directory that points at the wrong place', () => {
    const zip = sample();
    zip.writeUInt32LE(0, zip.length - EOCD_SIZE + 16);
    assert.throws(() => ZipArchive.open(zip), /Corrupt ZIP central directory at entry 0/);
  });

  it('reports a central directory that runs past the end of the file', () => {
    const zip = sample();
    zip.writeUInt16LE(5, zip.length - EOCD_SIZE + 10);
    assert.throws(() => ZipArchive.open(zip), /Corrupt ZIP central directory at entry 1/);
  });

  it('reports a damaged local header when the entry is read', () => {
    const zip = sample();
    zip.writeUInt32LE(0, 0);
    const archive = ZipArchive.open(zip);
    assert.throws(() => archive.read('a.txt'), /Corrupt ZIP local header: a\.txt/);
  });

  it('reports compression methods other than stored and deflate', () => {
    const zip = sample();
    zip.writeUInt16LE(12, centralOffset(zip) + 10);
    const archive = ZipArchive.open(zip);
    assert.throws(() => archive.read('a.txt'), /Unsupported ZIP compression method 12: a\.txt/);
  });

  it('reports truncated archives as not ZIP', () => {
    const zip = writeZip([{ name: 'a.txt', data: 'x'.repeat(1000) }]);
    assert.throws(() => ZipArchive.open(zip.subarray(0, zip.length - 30)), /Not a ZIP archive/);
  });

  it('opens an empty archive', () => {
    const archive = ZipArchive.open(writeZip([]));
    assert.deepEqual(archive.names(), []);
  });

  it('reads binary stored entries unchanged', () => {
    const bytes = Buffer.from([0, 255, 1, 254, 0x50, 0x4b, 0x05, 0x06]);
    const archive = ZipArchive.open(writeZip([{ name: 'bin', data: bytes, store: true }]));
    assert.deepEqual([...archive.read('bin')], [...bytes]);
  });
});

describe('ZipArchive against ZIP bombs', () => {
  const limits = { maxEntries: 3, maxEntrySize: 1000, maxTotalSize: 1500, maxCompressionRatio: 1100 };

  it('refuses entries and archives whose declared size exceeds the limits', () => {
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'big', data: 'x'.repeat(1001) }]), limits), /ZIP entry too large: big/);
    assert.throws(
      () => ZipArchive.open(writeZip([{ name: 'a', data: 'x'.repeat(900) }, { name: 'b', data: 'x'.repeat(900) }]), limits),
      /ZIP archive too large when unpacked/,
    );
  });

  it('refuses archives with more entries than the limit', () => {
    const entries = ['a', 'b', 'c', 'd'].map((name) => ({ name, data: name }));
    assert.throws(() => ZipArchive.open(writeZip(entries), limits), /too many entries \(4; limit 3\)/);
    assert.equal(ZipArchive.open(writeZip(entries.slice(0, 3)), limits).names().length, 3);
  });

  it('refuses large entries whose declared size is beyond what deflate can produce', () => {
    assert.throws(
      () => ZipArchive.open(writeZip([{ name: 'forged', data: 'x'.repeat(100), declaredSize: 2 * 1024 * 1024 }])),
      /ZIP entry compressed suspiciously well: forged \(more than 1100:1\)/,
    );
    // Real deflate output of a large run of zeros stays within the ratio.
    const zeros = ZipArchive.open(writeZip([{ name: 'zeros', data: Buffer.alloc(4 * 1024 * 1024) }]));
    assert.equal(zeros.read('zeros').length, 4 * 1024 * 1024);
  });

  it('stops a deflate stream that inflates beyond its declared size', () => {
    const zip = ZipArchive.open(writeZip([{ name: 'bomb', data: Buffer.alloc(5_000_000), declaredSize: 10 }]));
    assert.throws(() => zip.read('bomb'), /inflates beyond its declared size\): bomb/);
  });

  it('rejects size mismatches, truncated data and ZIP64 markers', () => {
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'short', data: 'abc', declaredSize: 10 }])).read('short'), /\(size mismatch\): short/);
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'cut', data: 'abc', store: true, declaredCompressedSize: 1_000_000 }])).read('cut'), /\(truncated data\): cut/);
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'z64', data: 'abc', declaredSize: 0xffffffff }])), /ZIP64 archives are not supported/);
  });
});
