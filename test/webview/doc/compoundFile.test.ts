import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CompoundFile } from '../../../src/webview/doc/compoundFile.ts';
import { writeCompoundFile, type CompoundFileOptions } from '../../support/compoundFileWriter.ts';

const END_OF_CHAIN = 0xfffffffe;

function bytes(length: number, seed: number): Buffer {
  return Buffer.from(Array.from({ length }, (_, i) => (i * 31 + seed) & 0xff));
}

const SMALL = bytes(1000, 1);
const LARGE = bytes(20_000, 2);

function sample(options: CompoundFileOptions = {}): Buffer {
  return writeCompoundFile(
    [
      { name: 'Small', data: SMALL },
      { name: 'Large', data: LARGE },
      { name: 'Empty', data: new Uint8Array() },
    ],
    options,
  );
}

function chainOf(chains: ReadonlyMap<string, readonly number[]>, name: string): readonly number[] {
  const chain = chains.get(name);
  assert.ok(chain && chain.length >= 2, `${name} spans at least two sectors`);
  return chain;
}

describe('CompoundFile', () => {
  it('reads streams from the mini stream and from regular sectors', () => {
    const file = CompoundFile.open(sample());
    assert.deepEqual(Buffer.from(file.readStream('Small')), SMALL);
    assert.deepEqual(Buffer.from(file.readStream('Large')), LARGE);
    assert.equal(file.readStream('Empty').byteLength, 0);
  });

  it('reads version 4 files with 4096-byte sectors', () => {
    const file = CompoundFile.open(sample({ version: 4 }));
    assert.deepEqual(Buffer.from(file.readStream('Small')), SMALL);
    assert.deepEqual(Buffer.from(file.readStream('Large')), LARGE);
  });

  it('finds root streams by name, ignoring case, and only at the root', () => {
    const file = CompoundFile.open(sample());
    assert.equal(file.has('SMALL'), true);
    assert.equal(file.has('large'), true);
    assert.equal(file.has('Missing'), false);
    assert.throws(() => file.readStream('Missing'), /Corrupt compound file: the Missing stream is missing/);
  });

  it('walks a directory tree with many siblings', () => {
    const streams = Array.from({ length: 40 }, (_, i) => ({ name: `Stream ${i}`, data: bytes(i * 10, i) }));
    const file = CompoundFile.open(writeCompoundFile(streams));
    for (const stream of streams) {
      assert.deepEqual(Buffer.from(file.readStream(stream.name)), Buffer.from(stream.data));
    }
  });
});

describe('CompoundFile with hostile files', () => {
  it('refuses files that are not compound files', () => {
    assert.throws(() => CompoundFile.open(Buffer.from('plain text')), /Not a compound file/);
    const header = sample();
    header[0] = 0;
    assert.throws(() => CompoundFile.open(header), /Not a compound file/);
  });

  it('refuses unknown byte orders and sector sizes', () => {
    const order = sample();
    order.writeUInt16LE(0xfeff, 0x1c);
    assert.throws(() => CompoundFile.open(order), /unknown byte order/);
    const shift = sample();
    shift.writeUInt16LE(30, 0x1e);
    assert.throws(() => CompoundFile.open(shift), /unsupported sector size/);
  });

  it('detects a FAT chain that loops back on itself', () => {
    const looping = sample({
      editFat(fat, chains) {
        const chain = chainOf(chains, 'Large');
        fat[chain[2] ?? 0] = chain[0] ?? 0;
      },
    });
    assert.throws(() => CompoundFile.open(looping).readStream('Large'), /the Large stream loops back on itself/);
  });

  it('detects a FAT chain that points outside the file or ends early', () => {
    const outside = sample({
      editFat(fat, chains) {
        fat[chainOf(chains, 'Large')[0] ?? 0] = 0x7fffffff;
      },
    });
    assert.throws(() => CompoundFile.open(outside).readStream('Large'), /the Large stream points outside the file/);
    const short = sample({
      editFat(fat, chains) {
        fat[chainOf(chains, 'Large')[1] ?? 0] = END_OF_CHAIN;
      },
    });
    assert.throws(() => CompoundFile.open(short).readStream('Large'), /the Large stream is shorter than its declared size/);
  });

  it('detects a looping directory chain', () => {
    const streams = Array.from({ length: 6 }, (_, i) => ({ name: `S${i}`, data: bytes(10, i) }));
    const looping = writeCompoundFile(streams, {
      editFat(fat, chains) {
        const directory = chainOf(chains, '<directory>');
        fat[directory[1] ?? 0] = directory[0] ?? 0;
      },
    });
    assert.throws(() => CompoundFile.open(looping), /the directory loops back on itself/);
  });

  it('detects a looping mini FAT chain', () => {
    const looping = sample({
      editMiniFat(miniFat, chains) {
        const chain = chainOf(chains, 'Small');
        miniFat[chain[3] ?? 0] = chain[1] ?? 0;
      },
    });
    assert.throws(() => CompoundFile.open(looping).readStream('Small'), /the Small stream loops back on itself/);
  });

  it('detects directory trees that loop or point outside the directory', () => {
    const looping = sample({
      editDirectory(entries) {
        const root = entries[0];
        const child = root ? entries[root.child] : undefined;
        if (child) {
          child.left = root?.child ?? 0;
        }
      },
    });
    assert.throws(() => CompoundFile.open(looping), /the directory tree loops/);
    const dangling = sample({
      editDirectory(entries) {
        const child = entries[entries[0]?.child ?? 0];
        if (child) {
          child.right = 5000;
        }
      },
    });
    assert.throws(() => CompoundFile.open(dangling), /a directory entry points outside the directory/);
  });

  it('refuses a missing root entry and streams larger than the file', () => {
    const noRoot = sample({
      editDirectory(entries) {
        const root = entries[0];
        if (root) {
          root.type = 1;
        }
      },
    });
    assert.throws(() => CompoundFile.open(noRoot), /the root entry is missing/);
    const huge = sample({
      editDirectory(entries) {
        const large = entries.find((entry) => entry.name === 'Large');
        if (large) {
          large.size = 0x7fffffff;
        }
      },
    });
    assert.throws(() => CompoundFile.open(huge).readStream('Large'), /the Large stream is larger than the file/);
  });

  it('refuses a header that claims more FAT sectors than the file holds', () => {
    const file = sample();
    file.writeUInt32LE(0xffffffff, 0x2c);
    assert.throws(() => CompoundFile.open(file), /lists 4294967295 FAT sectors, more than the file holds/);
  });

  it('reads only the FAT sectors that cover the file, whatever count the header claims', () => {
    const file = sample();
    file.writeUInt32LE(file.length / 512 - 1, 0x2c);
    assert.deepEqual(Buffer.from(CompoundFile.open(file).readStream('Large')), LARGE);
  });

  it('refuses DIFAT chains that end early or loop', () => {
    // Only the FAT sectors that cover the file are read, so the DIFAT matters for files of more than 109 * 128 sectors.
    const bare = (sectors: number, fatSectors: number, difatStart: number, difatSectors: number) => {
      const file = Buffer.alloc((sectors + 1) * 512);
      sample().copy(file, 0, 0, 0x4c);
      file.writeUInt32LE(fatSectors, 0x2c);
      file.writeUInt32LE(difatStart, 0x44);
      file.writeUInt32LE(difatSectors, 0x48);
      for (let i = 0; i < 109; i++) {
        file.writeUInt32LE(i, 0x4c + i * 4);
      }
      return file;
    };
    assert.throws(() => CompoundFile.open(bare(110 * 128, 110, END_OF_CHAIN, 0)), /the DIFAT chain ends before all FAT sectors are listed/);

    // One DIFAT sector (sector 200) lists 127 more FAT sectors, then links back to itself.
    const looping = bare(240 * 128, 240, 200, 5);
    const difat = (200 + 1) * 512;
    for (let i = 0; i < 127; i++) {
      looping.writeUInt32LE(300 + i, difat + i * 4);
    }
    looping.writeUInt32LE(200, difat + 127 * 4);
    assert.throws(() => CompoundFile.open(looping), /the DIFAT chain loops/);
  });

  it('refuses a truncated file whose streams run past its end', () => {
    const file = sample();
    // The last two sectors hold the end of Large; the file is still longer than the stream.
    const truncated = file.subarray(0, file.length - 2 * 512);
    assert.ok(truncated.length > LARGE.length);
    assert.throws(() => CompoundFile.open(truncated).readStream('Large'), /points outside the file/);
  });
});
