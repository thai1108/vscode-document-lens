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
    assert.throws(() => ZipArchive.open(zip), /the central directory is not where the end record says/);
  });

  it('reports a central directory that runs past the end of the file', () => {
    const zip = sample();
    zip.writeUInt16LE(5, zip.length - EOCD_SIZE + 8);
    zip.writeUInt16LE(5, zip.length - EOCD_SIZE + 10);
    assert.throws(() => ZipArchive.open(zip), /Corrupt ZIP central directory at entry 1/);
  });

  it('reports a damaged local header when the archive is opened', () => {
    const zip = sample();
    zip.writeUInt32LE(0, 0);
    assert.throws(() => ZipArchive.open(zip), /Corrupt ZIP local header: a\.txt/);
  });

  it('reports compression methods other than stored and deflate', () => {
    const zip = sample();
    zip.writeUInt16LE(12, centralOffset(zip) + 10);
    zip.writeUInt16LE(12, 8);
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

describe('ZipArchive against archives that readers see differently', () => {
  const unicodePath = (name: string) => {
    const value = Buffer.concat([Buffer.from([1, 0, 0, 0, 0]), Buffer.from(name, 'utf8')]);
    const header = Buffer.alloc(4);
    header.writeUInt16LE(0x7075, 0);
    header.writeUInt16LE(value.length, 2);
    return Buffer.concat([header, value]);
  };

  it('refuses data prepended to the archive', () => {
    const zip = writeZip([{ name: 'a.txt', data: 'A' }], '', { prefix: Buffer.from('MZ junk before the archive') });
    assert.throws(() => ZipArchive.open(zip), /the central directory is not where the end record says/);
  });

  it('refuses data after the archive and end records with inconsistent counts', () => {
    assert.throws(() => ZipArchive.open(Buffer.concat([sample(), Buffer.from('trailing')])), /data after the archive comment/);
    const zip = sample();
    zip.writeUInt16LE(2, zip.length - EOCD_SIZE + 8);
    assert.throws(() => ZipArchive.open(zip), /inconsistent entry counts/);
    const zip64 = sample();
    zip64.writeUInt16LE(0xffff, zip64.length - EOCD_SIZE + 10);
    assert.throws(() => ZipArchive.open(zip64), /ZIP64 archives are not supported/);
  });

  it('refuses a directory size that does not match its entries', () => {
    const zip = writeZip([{ name: 'a.txt', data: 'A' }, { name: 'b.txt', data: 'B' }]);
    zip.writeUInt16LE(1, zip.length - EOCD_SIZE + 8);
    zip.writeUInt16LE(1, zip.length - EOCD_SIZE + 10);
    assert.throws(() => ZipArchive.open(zip), /its size does not match its entries/);
  });

  it('refuses names a reader would normalize', () => {
    for (const name of ['x/../word/document.xml', './word/document.xml', 'word//document.xml', 'word\\document.xml', '/word/document.xml', 'word/./a', '..', 'a\0b']) {
      assert.throws(() => ZipArchive.open(writeZip([{ name, data: 'x' }])), /not canonical/, JSON.stringify(name));
    }
    assert.deepEqual(ZipArchive.open(writeZip([{ name: 'word/', data: '' }, { name: 'word/画面.xml', data: 'x' }])).names(), ['word/', 'word/画面.xml']);
  });

  it('refuses non-ASCII names without the UTF-8 flag and invalid UTF-8', () => {
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'café.xml', data: 'x', noUtf8Flag: true }])), /not valid UTF-8/);
    const zip = writeZip([{ name: 'ab.xml', data: 'x' }]);
    const at = centralOffset(zip) + 46;
    zip[at] = 0xff;
    zip[30] = 0xff;
    assert.throws(() => ZipArchive.open(zip), /not valid UTF-8/);
  });

  it('refuses a Unicode Path extra field that names another entry, and accepts a matching one', () => {
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'decoy.bin', data: 'x', extra: unicodePath('word/document.xml') }])), /second, different name: decoy\.bin/);
    assert.equal(ZipArchive.open(writeZip([{ name: 'a.xml', data: 'x', extra: unicodePath('a.xml') }])).readText('a.xml'), 'x');
    const broken = Buffer.from([0x75, 0x70, 0xff, 0x00]);
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'a.xml', data: 'x', extra: broken }])), /Corrupt ZIP extra field: a\.xml/);
  });

  it('refuses duplicate names and local headers that disagree with the central directory', () => {
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'a.xml', data: 'small' }, { name: 'a.xml', data: 'other' }])), /two entries named a\.xml/);
    assert.throws(() => ZipArchive.open(writeZip([{ name: 'a.xml', data: 'x', localName: 'b.xml' }])), /local header disagrees with the central directory: a\.xml/);
    const zip = sample();
    zip.writeUInt32LE(99, 22);
    assert.throws(() => ZipArchive.open(zip), /local header disagrees/);
  });

  it('refuses entries that share bytes', () => {
    const zip = writeZip([{ name: 'a.txt', data: 'A'.repeat(50), store: true }, { name: 'b.txt', data: 'B'.repeat(50), store: true }]);
    // Point the second entry at the first one's local header.
    const second = centralOffset(zip) + 46 + 'a.txt'.length;
    zip.writeUInt32LE(0, second + 42);
    assert.throws(() => ZipArchive.open(zip), /local header disagrees with the central directory: b\.txt/);
    const overlapping = writeZip([{ name: 'a.txt', data: 'A'.repeat(50), store: true }, { name: 'b.txt', data: 'B'.repeat(50), store: true }]);
    overlapping.writeUInt32LE(70, centralOffset(overlapping) + 20);
    overlapping.writeUInt32LE(70, 18);
    assert.throws(() => ZipArchive.open(overlapping), /ZIP entries overlap: a\.txt and b\.txt/);
  });
});
