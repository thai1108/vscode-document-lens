/**
 * Minimal OLE compound file ([MS-CFB]) writer for test fixtures: streams in the root storage, version 3 (512-byte
 * sectors) or 4 (4096-byte sectors). Like real writers it keeps streams under 4096 bytes in the mini stream and moves
 * the FAT sector list into DIFAT sectors past 109 FAT sectors. `patchFat` edits the FAT before it is written, for
 * broken-file tests.
 */

const END_OF_CHAIN = 0xfffffffe;
const FREE_SECTOR = 0xffffffff;
const FAT_SECTOR = 0xfffffffd;
const DIFAT_SECTOR = 0xfffffffc;
const NO_STREAM = 0xffffffff;
const MINI_SECTOR_SIZE = 64;
const MINI_STREAM_CUTOFF = 4096;
const DIRECTORY_ENTRY_SIZE = 128;
const HEADER_DIFAT_ENTRIES = 109;

export interface CompoundStream {
  name: string;
  data: Uint8Array;
}

/** Where things ended up, for tests that patch the file. */
export interface CompoundLayout {
  /** First sector (regular or mini) of each stream. */
  streams: Map<string, { start: number; mini: boolean }>;
  fatSectors: number[];
  sectorSize: number;
}

export interface CompoundFileOptions {
  version?: 3 | 4;
  /** CLSID of the root storage (the application that wrote the file), as 16 bytes. */
  rootClsid?: Uint8Array;
  patchFat?: (fat: number[], layout: CompoundLayout) => void;
  patchMiniFat?: (miniFat: number[], layout: CompoundLayout) => void;
}

export function writeCompoundFile(streams: readonly CompoundStream[], options: CompoundFileOptions = {}): Buffer {
  const version = options.version ?? 3;
  const sectorSize = version === 4 ? 4096 : 512;
  const perSector = sectorSize / 4;
  const sectorsFor = (bytes: number) => Math.ceil(bytes / sectorSize);
  const layout: CompoundLayout = { streams: new Map(), fatSectors: [], sectorSize };

  // Small streams go into the mini stream, one 64-byte mini sector chain each.
  const miniFat: number[] = [];
  const miniParts: Buffer[] = [];
  for (const stream of streams.filter((candidate) => candidate.data.byteLength < MINI_STREAM_CUTOFF)) {
    const count = Math.ceil(stream.data.byteLength / MINI_SECTOR_SIZE);
    const start = count ? miniFat.length : END_OF_CHAIN;
    for (let i = 0; i < count; i++) {
      miniFat.push(i === count - 1 ? END_OF_CHAIN : miniFat.length + 1);
    }
    layout.streams.set(stream.name, { start, mini: true });
    miniParts.push(padTo(Buffer.from(stream.data), MINI_SECTOR_SIZE));
  }
  options.patchMiniFat?.(miniFat, layout);
  const miniStream = Buffer.concat(miniParts);
  const miniFatBytes = uint32s([...miniFat, ...new Array<number>((perSector - (miniFat.length % perSector)) % perSector).fill(FREE_SECTOR)]);

  const regular = streams.filter((candidate) => candidate.data.byteLength >= MINI_STREAM_CUTOFF);
  const directoryBytes = sectorsFor((streams.length + 1) * DIRECTORY_ENTRY_SIZE) * sectorSize;
  const dataSectors = sectorsFor(directoryBytes) + sectorsFor(miniFatBytes.length) + sectorsFor(miniStream.length) + regular.reduce((sum, stream) => sum + sectorsFor(stream.data.byteLength), 0);

  // The FAT must also map its own sectors and the DIFAT sectors.
  const difatSectorsFor = (fatSectors: number) => Math.max(0, Math.ceil((fatSectors - HEADER_DIFAT_ENTRIES) / (perSector - 1)));
  let fatCount = 1;
  while (fatCount * perSector < dataSectors + fatCount + difatSectorsFor(fatCount)) {
    fatCount += 1;
  }
  const difatCount = difatSectorsFor(fatCount);

  const fat: number[] = new Array<number>(fatCount * perSector).fill(FREE_SECTOR);
  let next = 0;
  const allocate = (count: number, mark?: number): number => {
    const start = count ? next : END_OF_CHAIN;
    for (let i = 0; i < count; i++) {
      fat[next] = mark ?? (i === count - 1 ? END_OF_CHAIN : next + 1);
      next += 1;
    }
    return start;
  };
  const fatStart = allocate(fatCount, FAT_SECTOR);
  layout.fatSectors = Array.from({ length: fatCount }, (_, i) => fatStart + i);
  const difatStart = allocate(difatCount, DIFAT_SECTOR);
  const directoryStart = allocate(sectorsFor(directoryBytes));
  const miniFatStart = allocate(sectorsFor(miniFatBytes.length));
  const miniStreamStart = allocate(sectorsFor(miniStream.length));
  for (const stream of regular) {
    layout.streams.set(stream.name, { start: allocate(sectorsFor(stream.data.byteLength)), mini: false });
  }
  options.patchFat?.(fat, layout);

  // Directory: the root, then the streams as a balanced binary tree ordered by (name length, upper-case name).
  const sorted = [...streams].sort((a, b) => a.name.length - b.name.length || compareUpper(a.name, b.name));
  const directory = Buffer.alloc(directoryBytes);
  const tree = (low: number, high: number): number => {
    if (low > high) {
      return NO_STREAM;
    }
    const middle = (low + high) >> 1;
    const stream = sorted[middle];
    if (stream) {
      const location = layout.streams.get(stream.name);
      writeEntry(directory, middle + 1, {
        name: stream.name,
        type: 2,
        left: tree(low, middle - 1),
        right: tree(middle + 1, high),
        child: NO_STREAM,
        start: location?.start ?? END_OF_CHAIN,
        size: stream.data.byteLength,
      });
    }
    return middle + 1;
  };
  const rootChild = tree(0, sorted.length - 1);
  writeEntry(directory, 0, {
    name: 'Root Entry',
    type: 5,
    left: NO_STREAM,
    right: NO_STREAM,
    child: rootChild,
    start: miniStream.length ? miniStreamStart : END_OF_CHAIN,
    size: miniStream.length,
    clsid: options.rootClsid,
  });
  for (let i = sorted.length + 1; i < directoryBytes / DIRECTORY_ENTRY_SIZE; i++) {
    writeEntry(directory, i, { name: '', type: 0, left: NO_STREAM, right: NO_STREAM, child: NO_STREAM, start: 0, size: 0 });
  }

  const header = Buffer.alloc(sectorSize);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(header, 0);
  header.writeUInt16LE(0x003e, 24);
  header.writeUInt16LE(version, 26);
  header.writeUInt16LE(0xfffe, 28);
  header.writeUInt16LE(version === 4 ? 12 : 9, 30);
  header.writeUInt16LE(6, 32);
  header.writeUInt32LE(version === 4 ? sectorsFor(directoryBytes) : 0, 40);
  header.writeUInt32LE(fatCount, 44);
  header.writeUInt32LE(directoryStart, 48);
  header.writeUInt32LE(MINI_STREAM_CUTOFF, 56);
  header.writeUInt32LE(miniFat.length ? miniFatStart : END_OF_CHAIN, 60);
  header.writeUInt32LE(sectorsFor(miniFatBytes.length), 64);
  header.writeUInt32LE(difatCount ? difatStart : END_OF_CHAIN, 68);
  header.writeUInt32LE(difatCount, 72);
  for (let i = 0; i < HEADER_DIFAT_ENTRIES; i++) {
    header.writeUInt32LE(layout.fatSectors[i] ?? FREE_SECTOR, 76 + i * 4);
  }

  const difat: number[] = [];
  for (let i = 0; i < difatCount; i++) {
    const entries = layout.fatSectors.slice(HEADER_DIFAT_ENTRIES + i * (perSector - 1), HEADER_DIFAT_ENTRIES + (i + 1) * (perSector - 1));
    difat.push(...entries, ...new Array<number>(perSector - 1 - entries.length).fill(FREE_SECTOR), i === difatCount - 1 ? END_OF_CHAIN : difatStart + i + 1);
  }
  const regularData = regular.map((stream) => padTo(Buffer.from(stream.data), sectorSize));
  return Buffer.concat([header, uint32s(fat), uint32s(difat), directory, miniFatBytes, padTo(miniStream, sectorSize), ...regularData]);
}

interface EntryFields {
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
  clsid?: Uint8Array;
}

function writeEntry(directory: Buffer, index: number, entry: EntryFields): void {
  const offset = index * DIRECTORY_ENTRY_SIZE;
  if (entry.name.length > 31) {
    throw new Error(`Compound file names are at most 31 characters: ${entry.name}`);
  }
  directory.write(entry.name, offset, 'utf16le');
  directory.writeUInt16LE(entry.name ? (entry.name.length + 1) * 2 : 0, offset + 64);
  directory.writeUInt8(entry.type, offset + 66);
  directory.writeUInt8(1, offset + 67);
  directory.writeUInt32LE(entry.left, offset + 68);
  directory.writeUInt32LE(entry.right, offset + 72);
  directory.writeUInt32LE(entry.child, offset + 76);
  if (entry.clsid) {
    directory.set(entry.clsid.subarray(0, 16), offset + 80);
  }
  directory.writeUInt32LE(entry.start, offset + 116);
  directory.writeUInt32LE(entry.size, offset + 120);
}

function compareUpper(a: string, b: string): number {
  const x = a.toUpperCase();
  const y = b.toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

function uint32s(values: readonly number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 4);
  values.forEach((value, i) => buffer.writeUInt32LE(value >>> 0, i * 4));
  return buffer;
}

function padTo(buffer: Buffer, multiple: number): Buffer {
  const padding = (multiple - (buffer.length % multiple)) % multiple;
  return padding ? Buffer.concat([buffer, Buffer.alloc(padding)]) : buffer;
}
