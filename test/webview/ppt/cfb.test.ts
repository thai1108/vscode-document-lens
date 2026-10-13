import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CompoundFile, CompoundFileError } from '../../../src/webview/ppt/cfb.ts';
import { writeCompoundFile, type CompoundFileOptions, type CompoundStream } from '../../support/cfbWriter.ts';

const LIMITS = { maxDirectoryEntries: 1000 };

/** Deterministic bytes, so a misplaced sector shows up as different content. */
function pattern(length: number, seed = 1): Uint8Array {
  const bytes = new Uint8Array(length);
  let state = seed;
  for (let i = 0; i < length; i++) {
    state = (state * 1103515245 + 12345) & 0x7fffffff;
    bytes[i] = state >> 16;
  }
  return bytes;
}

/** Large deterministic content: a 509-byte cycle, so a sector read from the wrong place does not match. */
function largePattern(length: number): Uint8Array {
  return new Uint8Array(Buffer.alloc(length, pattern(509, 11)));
}

/** The viewer receives a plain Uint8Array; the writer returns a Buffer. */
function plain(bytes: Uint8Array): Uint8Array {
  return new Uint8Array(bytes);
}

function open(streams: CompoundStream[], options?: CompoundFileOptions): CompoundFile {
  return CompoundFile.open(plain(writeCompoundFile(streams, options)), LIMITS);
}

function startOf(name: string, layout: { streams: Map<string, { start: number }> }): number {
  const start = layout.streams.get(name)?.start;
  assert.ok(start !== undefined, name);
  return start;
}

describe('CompoundFile', () => {
  it('reads small streams from the mini stream and large ones from regular sectors', () => {
    const small = pattern(100, 2);
    const large = pattern(70_000, 3);
    const file = open([
      { name: 'Small', data: small },
      { name: 'Large', data: large },
      { name: 'Empty', data: new Uint8Array() },
    ]);
    assert.deepEqual(file.stream('Small'), small);
    assert.deepEqual(file.stream('Large'), large);
    assert.deepEqual(file.stream('Empty'), new Uint8Array());
    assert.equal(file.stream('Missing'), undefined);
    assert.deepEqual(file.streamNames().sort(), ['Empty', 'Large', 'Small']);
  });

  it('finds streams by name regardless of case, also with many siblings', () => {
    const names = Array.from({ length: 40 }, (_, i) => `Stream ${i}`);
    const file = open(names.map((name, i) => ({ name, data: pattern(10 + i, i) })));
    assert.equal(file.has('stream 17'), true);
    assert.deepEqual(file.stream('STREAM 39'), pattern(49, 39));
  });

  it('returns a view into the file for contiguous streams instead of copying', () => {
    const bytes = plain(writeCompoundFile([{ name: 'Large', data: pattern(10_000) }]));
    const stream = CompoundFile.open(bytes, LIMITS).stream('Large');
    assert.equal(stream?.buffer, bytes.buffer);
  });

  it('reads version 4 files (4096-byte sectors)', () => {
    const large = pattern(20_000, 5);
    const file = open([{ name: 'Large', data: large }, { name: 'Small', data: pattern(300, 6) }], { version: 4 });
    assert.deepEqual(file.stream('Large'), large);
    assert.deepEqual(file.stream('Small'), pattern(300, 6));
  });

  it('follows the DIFAT chain past the 109 FAT sectors listed in the header', () => {
    // 109 FAT sectors map about 7 MB of 512-byte sectors; each DIFAT sector lists 127 more (about 8 MB).
    const large = largePattern(16_000_000);
    const bytes = writeCompoundFile([{ name: 'Large', data: large }]);
    assert.equal(bytes.readUInt32LE(72), 2, 'the file needs two DIFAT sectors');
    assert.deepEqual(CompoundFile.open(plain(bytes), LIMITS).stream('Large'), large);
  });

  it('reads a stream whose sectors are scattered', () => {
    const data = pattern(5000, 9);
    const bytes = writeCompoundFile([{ name: 'Data', data }], {
      patchFat: (fat, layout) => {
        // Swap the order of the stream's middle sectors: s → s+2 → s+1 → s+3.
        const start = startOf('Data', layout);
        fat[start] = start + 2;
        fat[start + 2] = start + 1;
        fat[start + 1] = start + 3;
      },
    });
    const stream = CompoundFile.open(plain(bytes), LIMITS).stream('Data');
    assert.deepEqual(stream?.subarray(0, 512), data.subarray(0, 512));
    assert.deepEqual(stream?.subarray(512, 1024), data.subarray(1024, 1536));
    assert.deepEqual(stream?.subarray(1024, 1536), data.subarray(512, 1024));
  });

  it('recognizes the signature', () => {
    assert.equal(CompoundFile.isCompoundFile(writeCompoundFile([])), true);
    assert.equal(CompoundFile.isCompoundFile(new TextEncoder().encode('PK\u0003\u0004')), false);
    assert.throws(() => CompoundFile.open(new Uint8Array(600), LIMITS), /no compound file header/);
  });
});

describe('CompoundFile with hostile files', () => {
  it('reports a FAT chain that loops', () => {
    const file = open([{ name: 'Data', data: pattern(5000) }], {
      patchFat: (fat, layout) => {
        const start = startOf('Data', layout);
        fat[start + 3] = start + 1;
      },
    });
    assert.throws(() => file.stream('Data'), (error: Error) => error instanceof CompoundFileError && /sector chain loops/.test(error.message));
  });

  it('reports a FAT chain that points outside the file or at a free sector', () => {
    for (const next of [0x7fffffff, 0xffffffff]) {
      const file = open([{ name: 'Data', data: pattern(5000) }], {
        patchFat: (fat, layout) => {
          fat[startOf('Data', layout) + 1] = next;
        },
      });
      assert.throws(() => file.stream('Data'), /sector chain is broken/);
    }
  });

  it('reports a mini FAT chain that loops', () => {
    const file = open([{ name: 'Small', data: pattern(300) }], {
      patchMiniFat: (miniFat, layout) => {
        const start = startOf('Small', layout);
        miniFat[start + 2] = start;
      },
    });
    assert.throws(() => file.stream('Small'), /sector chain loops/);
  });

  it('gives a truncated stream when the chain ends before the declared size', () => {
    const data = pattern(5000);
    const file = open([{ name: 'Data', data }], {
      patchFat: (fat, layout) => {
        fat[startOf('Data', layout) + 1] = 0xfffffffe;
      },
    });
    assert.deepEqual(file.stream('Data'), data.subarray(0, 1024));
  });

  it('reads a truncated file without reading past its end', () => {
    const data = pattern(5000);
    const bytes = writeCompoundFile([{ name: 'Data', data }]);
    const file = CompoundFile.open(plain(bytes.subarray(0, bytes.length - 700)), LIMITS);
    assert.throws(() => file.stream('Data'), /outside the file/);
  });

  it('refuses a DIFAT chain that loops or ends early', () => {
    const bytes = writeCompoundFile([{ name: 'Large', data: largePattern(16_000_000) }]);
    const firstDifat = bytes.readUInt32LE(68);
    // Point the first DIFAT sector's "next" (its last entry) at itself.
    bytes.writeUInt32LE(firstDifat, (firstDifat + 1) * 512 + 127 * 4);
    assert.throws(() => CompoundFile.open(plain(bytes), LIMITS), /DIFAT chain loops/);
    bytes.writeUInt32LE(1, 72);
    assert.throws(() => CompoundFile.open(plain(bytes), LIMITS), /DIFAT chain ends early/);
  });

  it('reads only the FAT sectors that map the file, however many the header claims', () => {
    const data = pattern(5000);
    const bytes = writeCompoundFile([{ name: 'Data', data }]);
    // As many FAT sectors as the file has sectors: only the first one exists.
    bytes.writeUInt32LE(Math.floor(bytes.length / 512) - 1, 44);
    assert.deepEqual(CompoundFile.open(plain(bytes), LIMITS).stream('Data'), data);
  });

  it('refuses more FAT sectors than the file holds', () => {
    const bytes = writeCompoundFile([{ name: 'Data', data: pattern(100) }]);
    bytes.writeUInt32LE(0x10000000, 44);
    assert.throws(() => CompoundFile.open(bytes, LIMITS), /more FAT sectors than the file holds/);
  });

  it('refuses unsupported headers', () => {
    const bytes = writeCompoundFile([]);
    bytes.writeUInt16LE(10, 30);
    assert.throws(() => CompoundFile.open(bytes, LIMITS), /unsupported sector size/);
    bytes.writeUInt16LE(9, 30);
    bytes.writeUInt16LE(0xfeff, 28);
    assert.throws(() => CompoundFile.open(bytes, LIMITS), /byte order/);
  });

  it('refuses mini sector sizes and mini stream cutoffs other than the fixed 64 and 4096 bytes', () => {
    const data = pattern(3000);
    const valid = writeCompoundFile([{ name: 'Small', data }]);
    assert.deepEqual(CompoundFile.open(plain(valid), LIMITS).stream('Small'), data);
    for (const [offset, value, size] of [
      [32, 1, 2],
      [32, 7, 2],
      [56, 0xffffffff, 4],
      [56, 0, 4],
      [56, 8192, 4],
    ] as const) {
      const bytes = Buffer.from(valid);
      bytes.writeUIntLE(value, offset, size);
      assert.throws(() => CompoundFile.open(plain(bytes), LIMITS), /invalid mini stream layout/, `header field at ${offset} = ${value}`);
    }
  });

  it('refuses at once a file that would read its stream as millions of 2-byte mini sectors', () => {
    // 2-byte mini sectors and a 4 GB cutoff made a 120 MB file use 2.4 GB of memory before these were checked.
    const bytes = writeCompoundFile([{ name: 'PowerPoint Document', data: largePattern(4 * 1024 * 1024) }]);
    bytes.writeUInt16LE(1, 32);
    bytes.writeUInt32LE(0xffffffff, 56);
    const started = performance.now();
    assert.throws(() => CompoundFile.open(plain(bytes), LIMITS), /invalid mini stream layout/);
    assert.ok(performance.now() - started < 100);
  });

  it('refuses a directory without a root entry', () => {
    const bytes = writeCompoundFile([{ name: 'Data', data: pattern(10) }]);
    const directory = (bytes.readUInt32LE(48) + 1) * 512;
    bytes.writeUInt8(1, directory + 66);
    assert.throws(() => CompoundFile.open(bytes, LIMITS), /no root directory entry/);
  });

  it('survives directory entries that point back at each other', () => {
    const bytes = writeCompoundFile([
      { name: 'A', data: pattern(10) },
      { name: 'B', data: pattern(20) },
      { name: 'C', data: pattern(30) },
    ]);
    const directory = (bytes.readUInt32LE(48) + 1) * 512;
    // Entries 1–3 are A, B, C; B is the root's child with A left and C right. Make A and C point back at B.
    bytes.writeUInt32LE(2, directory + 128 + 68);
    bytes.writeUInt32LE(2, directory + 3 * 128 + 72);
    const file = CompoundFile.open(bytes, LIMITS);
    assert.deepEqual(file.streamNames().sort(), ['A', 'B', 'C']);
  });

  it('caps the size a directory entry may claim at the file size, and refuses streams above a given size', () => {
    const bytes = writeCompoundFile([{ name: 'Data', data: pattern(5000) }], { version: 4 });
    const directory = (bytes.readUInt32LE(48) + 1) * 4096;
    // Entry 1's size becomes 2^40 + 5000 bytes.
    bytes.writeUInt32LE(0x100, directory + 128 + 124);
    const file = CompoundFile.open(plain(bytes), LIMITS);
    const stream = file.stream('Data');
    assert.ok(stream && stream.byteLength <= bytes.length);
    assert.deepEqual(stream.subarray(0, 5000), pattern(5000));
    assert.throws(() => file.stream('Data', 10_000), /the Data stream is too large/);
  });

  it('reads at most the configured number of directory entries', () => {
    const file = CompoundFile.open(
      writeCompoundFile(Array.from({ length: 10 }, (_, i) => ({ name: `S${i}`, data: pattern(5) }))),
      { maxDirectoryEntries: 4 },
    );
    assert.ok(file.streamNames().length <= 3);
  });
});
