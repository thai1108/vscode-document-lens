import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { ForgedSizeError, inflateEntry, readZipDirectory, readZipText } from '../../../src/webview/docx/zipDirectory.ts';
import { writeZip } from '../../support/zipWriter.ts';

const CENTRAL_HEADER = Buffer.from([0x50, 0x4b, 0x01, 0x02]);

/** Clears the UTF-8 flag of every central directory entry and replaces `from` with `to` (same length) in its names. */
function asLegacyNames(zip: Buffer, from: Buffer, to: Buffer): Buffer {
  const copy = Buffer.from(zip);
  for (let offset = copy.indexOf(CENTRAL_HEADER); offset >= 0; offset = copy.indexOf(CENTRAL_HEADER, offset + 4)) {
    copy.writeUInt16LE(copy.readUInt16LE(offset + 8) & ~0x0800, offset + 8);
    const nameStart = offset + 46;
    const name = copy.subarray(nameStart, nameStart + copy.readUInt16LE(offset + 28));
    const at = name.indexOf(from);
    if (at >= 0) {
      to.copy(name, at);
    }
  }
  return copy;
}

describe('docx ZIP directory', () => {
  it('lists entry names (UTF-8) and declared sizes', () => {
    const zip = writeZip([
      { name: 'word/document.xml', data: '<w:document/>' },
      { name: 'word/media/画像.png', data: Buffer.alloc(300), store: true },
    ]);
    const entries = readZipDirectory(zip);
    assert.deepEqual(
      entries.map(({ name, size, method, ambiguousName }) => ({ name, size, method, ambiguousName })),
      [
        { name: 'word/document.xml', size: 13, method: 8, ambiguousName: false },
        { name: 'word/media/画像.png', size: 300, method: 0, ambiguousName: false },
      ],
    );
  });

  it('marks names JSZip may decode differently', () => {
    const zip = writeZip([{ name: 'word/é.xml', data: 'x' }, { name: 'word/a.xml', data: 'y' }]);
    const entries = readZipDirectory(asLegacyNames(zip, Buffer.from('é'), Buffer.from([0xe9, 0x41])));
    assert.deepEqual(entries.map((entry) => entry.ambiguousName), [true, false]);
  });

  it('refuses data that is not a ZIP archive', () => {
    assert.throws(() => readZipDirectory(new TextEncoder().encode('not a zip file')), /Not a ZIP archive/);
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

  it('stops an entry that inflates past its declared size', async () => {
    const zip = writeZip([
      { name: 'stored.bin', data: Buffer.alloc(1000), store: true, declaredSize: 10 },
      { name: 'deflated.bin', data: Buffer.alloc(1_000_000), declaredSize: 1000 },
      { name: 'honest.bin', data: Buffer.alloc(1_000_000) },
    ]);
    const [stored, deflated, honest] = readZipDirectory(zip);
    assert.ok(stored && deflated && honest);
    await assert.rejects(inflateEntry(zip, stored), ForgedSizeError);
    let seen = 0;
    await assert.rejects(
      inflateEntry(zip, deflated, (chunk) => {
        seen += chunk.byteLength;
      }),
      ForgedSizeError,
    );
    assert.ok(seen <= 1000, `stopped after ${seen} bytes`);
    await inflateEntry(zip, honest);
  });
});
