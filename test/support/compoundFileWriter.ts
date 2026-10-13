// Minimal OLE compound file writer ([MS-CFB] version 3 with 512-byte sectors, or version 4 with 4096-byte sectors) for
// building test files: streams under the root storage, small ones in the mini stream. Hooks let tests break the FAT or
// directory to forge hostile files.

const MINI_SECTOR = 64;
const MINI_CUTOFF = 4096;
const END_OF_CHAIN = 0xfffffffe;
const FREE = 0xffffffff;
const FAT_SECTOR = 0xfffffffd;
const NO_STREAM = 0xffffffff;

export interface CompoundStream {
  name: string;
  data: Uint8Array;
}

export interface DirectoryEntryInput {
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

export interface CompoundFileOptions {
  version?: 3 | 4;
  /** Edits the FAT before it is written; `chains` holds each stream's sectors (and '<directory>', '<mini stream>'). */
  editFat?(fat: Uint32Array, chains: ReadonlyMap<string, readonly number[]>): void;
  /** Edits the mini FAT before it is written; `chains` holds each small stream's mini sectors. */
  editMiniFat?(miniFat: Uint32Array, chains: ReadonlyMap<string, readonly number[]>): void;
  /** Edits the directory entries (0 is the root) before they are written. */
  editDirectory?(entries: DirectoryEntryInput[]): void;
}

/** [MS-CFB] orders siblings by name length, then by upper-cased name. */
function compareNames(a: string, b: string): number {
  return a.length - b.length || (a.toUpperCase() < b.toUpperCase() ? -1 : a.toUpperCase() > b.toUpperCase() ? 1 : 0);
}

export function writeCompoundFile(streams: CompoundStream[], options: CompoundFileOptions = {}): Buffer {
  const version = options.version ?? 3;
  const sectorSize = version === 4 ? 4096 : 512;
  const small = (stream: CompoundStream) => stream.data.byteLength < MINI_CUTOFF;
  const sectorsFor = (bytes: number, size: number) => Math.ceil(bytes / size);

  // Mini stream: small streams back to back, each padded to 64 bytes.
  const miniChains = new Map<string, number[]>();
  let miniSectors = 0;
  for (const stream of streams.filter(small)) {
    const count = sectorsFor(stream.data.byteLength, MINI_SECTOR);
    miniChains.set(stream.name, Array.from({ length: count }, (_, i) => miniSectors + i));
    miniSectors += count;
  }
  const miniFatSectors = sectorsFor(miniSectors * 4, sectorSize);
  const miniContainerSectors = sectorsFor(miniSectors * MINI_SECTOR, sectorSize);
  const directorySectors = sectorsFor((streams.length + 1) * 128, sectorSize);
  const entriesPerSector = sectorSize / 128;
  const bigSectors = streams.filter((stream) => !small(stream)).reduce((sum, stream) => sum + sectorsFor(stream.data.byteLength, sectorSize), 0);
  const others = directorySectors + miniFatSectors + miniContainerSectors + bigSectors;
  const fatSectors = Math.ceil(others / (sectorSize / 4 - 1));
  if (fatSectors > 109) {
    throw new Error('writeCompoundFile: too large for a header-only DIFAT');
  }

  // Sector layout: FAT, directory, mini FAT, mini stream, then the large streams.
  let next = fatSectors;
  const take = (count: number) => Array.from({ length: count }, () => next++);
  const chains = new Map<string, number[]>();
  chains.set('<directory>', take(directorySectors));
  const miniFatChain = take(miniFatSectors);
  chains.set('<mini stream>', take(miniContainerSectors));
  for (const stream of streams.filter((candidate) => !small(candidate))) {
    chains.set(stream.name, take(sectorsFor(stream.data.byteLength, sectorSize)));
  }

  const fat = new Uint32Array((fatSectors * sectorSize) / 4).fill(FREE);
  for (let i = 0; i < fatSectors; i++) {
    fat[i] = FAT_SECTOR;
  }
  const link = (chain: readonly number[], table: Uint32Array) =>
    chain.forEach((sector, i) => {
      table[sector] = chain[i + 1] ?? END_OF_CHAIN;
    });
  [...chains.values(), miniFatChain].forEach((chain) => link(chain, fat));
  options.editFat?.(fat, chains);

  const miniFat = new Uint32Array((miniFatSectors * sectorSize) / 4).fill(FREE);
  miniChains.forEach((chain) => link(chain, miniFat));
  options.editMiniFat?.(miniFat, miniChains);

  // Directory: the root, then the streams as a binary search tree under it (all black; readers do not rebalance).
  const sorted = streams.map((stream, i) => ({ stream, id: i + 1 })).sort((a, b) => compareNames(a.stream.name, b.stream.name));
  const entries: DirectoryEntryInput[] = [
    { name: 'Root Entry', type: 5, left: NO_STREAM, right: NO_STREAM, child: NO_STREAM, start: chains.get('<mini stream>')?.[0] ?? END_OF_CHAIN, size: miniSectors * MINI_SECTOR },
    ...streams.map((stream) => ({
      name: stream.name,
      type: 2,
      left: NO_STREAM,
      right: NO_STREAM,
      child: NO_STREAM,
      start: (small(stream) ? miniChains.get(stream.name)?.[0] : chains.get(stream.name)?.[0]) ?? END_OF_CHAIN,
      size: stream.data.byteLength,
    })),
  ];
  const tree = (low: number, high: number): number => {
    if (low > high) {
      return NO_STREAM;
    }
    const mid = (low + high) >> 1;
    const node = sorted[mid];
    const entry = node ? entries[node.id] : undefined;
    if (!node || !entry) {
      return NO_STREAM;
    }
    entry.left = tree(low, mid - 1);
    entry.right = tree(mid + 1, high);
    return node.id;
  };
  const root = entries[0];
  if (root) {
    root.child = tree(0, sorted.length - 1);
  }
  options.editDirectory?.(entries);

  const out = Buffer.alloc((next + 1) * sectorSize);
  const at = (sector: number) => (sector + 1) * sectorSize;

  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(out, 0);
  out.writeUInt16LE(0x003e, 0x18);
  out.writeUInt16LE(version, 0x1a);
  out.writeUInt16LE(0xfffe, 0x1c);
  out.writeUInt16LE(version === 4 ? 12 : 9, 0x1e);
  out.writeUInt16LE(6, 0x20);
  out.writeUInt32LE(version === 4 ? directorySectors : 0, 0x28);
  out.writeUInt32LE(fatSectors, 0x2c);
  out.writeUInt32LE(chains.get('<directory>')?.[0] ?? END_OF_CHAIN, 0x30);
  out.writeUInt32LE(MINI_CUTOFF, 0x38);
  out.writeUInt32LE(miniFatChain[0] ?? END_OF_CHAIN, 0x3c);
  out.writeUInt32LE(miniFatSectors, 0x40);
  out.writeUInt32LE(END_OF_CHAIN, 0x44);
  out.writeUInt32LE(0, 0x48);
  for (let i = 0; i < 109; i++) {
    out.writeUInt32LE(i < fatSectors ? i : FREE, 0x4c + i * 4);
  }

  fat.forEach((value, i) => out.writeUInt32LE(value, sectorSize + i * 4));
  const writeChain = (chain: readonly number[], data: Uint8Array) =>
    chain.forEach((sector, i) => out.set(data.subarray(i * sectorSize, (i + 1) * sectorSize), at(sector)));

  const directory = Buffer.alloc(directorySectors * sectorSize);
  entries.forEach((entry, i) => {
    const offset = i * 128;
    const name = Buffer.from(`${entry.name}\0`, 'utf16le');
    name.copy(directory, offset, 0, Math.min(name.length, 64));
    directory.writeUInt16LE(Math.min(name.length, 64), offset + 0x40);
    directory.writeUInt8(entry.type, offset + 0x42);
    directory.writeUInt8(1, offset + 0x43);
    directory.writeUInt32LE(entry.left, offset + 0x44);
    directory.writeUInt32LE(entry.right, offset + 0x48);
    directory.writeUInt32LE(entry.child, offset + 0x4c);
    directory.writeUInt32LE(entry.start, offset + 0x74);
    directory.writeUInt32LE(entry.size, offset + 0x78);
  });
  for (let i = entries.length; i < directorySectors * entriesPerSector; i++) {
    directory.writeUInt32LE(NO_STREAM, i * 128 + 0x44);
    directory.writeUInt32LE(NO_STREAM, i * 128 + 0x48);
    directory.writeUInt32LE(NO_STREAM, i * 128 + 0x4c);
  }
  writeChain(chains.get('<directory>') ?? [], directory);

  writeChain(miniFatChain, Buffer.from(miniFat.buffer, miniFat.byteOffset, miniFat.byteLength));
  const miniStream = Buffer.alloc(miniContainerSectors * sectorSize);
  for (const stream of streams.filter(small)) {
    const first = miniChains.get(stream.name)?.[0] ?? 0;
    miniStream.set(stream.data, first * MINI_SECTOR);
  }
  writeChain(chains.get('<mini stream>') ?? [], miniStream);

  for (const stream of streams.filter((candidate) => !small(candidate))) {
    writeChain(chains.get(stream.name) ?? [], stream.data);
  }
  return out;
}
