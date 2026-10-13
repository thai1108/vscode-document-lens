// Reader for OLE compound files ([MS-CFB]), the container of .doc files: a small FAT file system inside one file.
// The file is untrusted, so every sector index, chain, size and directory link is checked against the file before use:
// chains that loop or leave the file, and streams larger than the file, are reported as corruption.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const HEADER_SIZE = 512;
const HEADER_DIFAT_ENTRIES = 109;
const MAX_REGULAR_SECTOR = 0xfffffffa;
const END_OF_CHAIN = 0xfffffffe;
const FREE_SECTOR = 0xffffffff;
const NO_STREAM = 0xffffffff;
const DIRECTORY_ENTRY_SIZE = 128;
const MINI_SECTOR_SIZE = 64;
/** Streams smaller than this live in the mini stream ([MS-CFB] fixes it at 4096). */
const MINI_STREAM_CUTOFF = 4096;

const ENTRY_STORAGE = 1;
const ENTRY_STREAM = 2;
const ENTRY_ROOT = 5;

interface DirectoryEntry {
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

function corrupt(detail: string): Error {
  return new Error(`Corrupt compound file: ${detail}`);
}

export class CompoundFile {
  private readonly bytes: Uint8Array;
  private readonly view: DataView;
  private readonly sectorSize: number;
  private readonly sectorCount: number;
  private readonly fat: Uint32Array;
  private readonly directory: Uint8Array;
  private readonly root: DirectoryEntry;
  private readonly rootChildren: Map<string, DirectoryEntry>;
  private miniStream: Uint8Array | undefined;
  private miniFat: Uint32Array | undefined;
  private readonly miniFatStart: number;
  private readonly miniFatSectors: number;
  private readonly majorVersion: number;

  private constructor(bytes: Uint8Array) {
    this.bytes = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (bytes.byteLength < HEADER_SIZE || !SIGNATURE.every((byte, i) => bytes[i] === byte)) {
      throw new Error('Not a compound file (no compound file header)');
    }
    const u16 = (offset: number) => this.view.getUint16(offset, true);
    const u32 = (offset: number) => this.view.getUint32(offset, true);
    if (u16(0x1c) !== 0xfffe) {
      throw corrupt('unknown byte order');
    }
    const sectorShift = u16(0x1e);
    if ((sectorShift !== 9 && sectorShift !== 12) || u16(0x20) !== 6) {
      throw corrupt(`unsupported sector size (2^${sectorShift})`);
    }
    this.sectorSize = 1 << sectorShift;
    // Sector n starts at (n + 1) * sectorSize; a short last sector is read as if padded with zeros.
    this.sectorCount = Math.max(0, Math.ceil((bytes.byteLength - this.sectorSize) / this.sectorSize));
    this.fat = this.readFat(u32(0x2c), u32(0x44), u32(0x48));
    this.miniFatStart = u32(0x3c);
    this.miniFatSectors = u32(0x40);

    this.directory = this.readChain(u32(0x30), undefined, 'the directory');
    this.majorVersion = u16(0x1a);
    const root = this.entry(0);
    if (root?.type !== ENTRY_ROOT) {
      throw corrupt('the root entry is missing');
    }
    this.root = root;
    this.rootChildren = this.childrenOf(root);
  }

  static open(bytes: Uint8Array): CompoundFile {
    return new CompoundFile(bytes);
  }

  /** Whether the root storage holds a stream or storage of this name (case-insensitive, as in [MS-CFB]). */
  has(name: string): boolean {
    return this.rootChildren.has(name.toUpperCase());
  }

  /** The contents of a stream directly inside the root storage. */
  readStream(name: string): Uint8Array {
    const entry = this.rootChildren.get(name.toUpperCase());
    if (entry?.type !== ENTRY_STREAM) {
      throw corrupt(`the ${name} stream is missing`);
    }
    if (entry.size > this.bytes.byteLength) {
      throw corrupt(`the ${name} stream is larger than the file`);
    }
    if (entry.size === 0) {
      return new Uint8Array();
    }
    return entry.size < MINI_STREAM_CUTOFF ? this.readMiniChain(entry.start, entry.size, `the ${name} stream`) : this.readChain(entry.start, entry.size, `the ${name} stream`);
  }

  /**
   * The FAT, gathered from the sector list in the header and the DIFAT sectors it continues in. Only the FAT sectors
   * that describe sectors inside the file are read, so a forged sector count cannot inflate it.
   */
  private readFat(fatSectors: number, difatStart: number, difatSectors: number): Uint32Array {
    if (fatSectors > this.sectorCount) {
      throw corrupt(`the header lists ${fatSectors} FAT sectors, more than the file holds`);
    }
    const perFatSector = this.sectorSize / 4;
    const wanted = Math.min(fatSectors, Math.ceil(this.sectorCount / perFatSector));
    const locations: number[] = [];
    for (let i = 0; i < HEADER_DIFAT_ENTRIES && locations.length < wanted; i++) {
      locations.push(this.view.getUint32(0x4c + i * 4, true));
    }
    const perDifatSector = perFatSector - 1;
    const seen = new Set<number>();
    let next = difatStart;
    for (let i = 0; locations.length < wanted; i++) {
      if (i >= difatSectors || next > MAX_REGULAR_SECTOR || next >= this.sectorCount) {
        throw corrupt('the DIFAT chain ends before all FAT sectors are listed');
      }
      if (seen.has(next)) {
        throw corrupt('the DIFAT chain loops');
      }
      seen.add(next);
      const offset = this.sectorOffset(next);
      for (let j = 0; j < perDifatSector && locations.length < wanted; j++) {
        locations.push(this.u32At(offset + j * 4));
      }
      next = this.u32At(offset + perDifatSector * 4);
    }
    const fat = new Uint32Array(wanted * perFatSector);
    locations.forEach((sector, i) => {
      if (sector >= this.sectorCount) {
        throw corrupt('a FAT sector lies outside the file');
      }
      const offset = this.sectorOffset(sector);
      for (let j = 0; j < perFatSector; j++) {
        fat[i * perFatSector + j] = this.u32At(offset + j * 4);
      }
    });
    return fat;
  }

  /**
   * Follows a FAT chain from `start`. With a known `size`, reads just enough sectors for it; without one (the directory),
   * reads to the end of the chain. Either way the walk visits each sector at most once.
   */
  private readChain(start: number, size: number | undefined, what: string): Uint8Array {
    const needed = size === undefined ? this.sectorCount : Math.ceil(size / this.sectorSize);
    const sectors = this.walk(start, this.fat, this.sectorCount, needed, size === undefined, what);
    const out = new Uint8Array(size ?? sectors.length * this.sectorSize);
    sectors.forEach((sector, i) => {
      const offset = this.sectorOffset(sector);
      const length = Math.min(this.sectorSize, out.byteLength - i * this.sectorSize);
      out.set(this.bytes.subarray(offset, Math.min(offset + length, this.bytes.byteLength)), i * this.sectorSize);
    });
    return out;
  }

  private readMiniChain(start: number, size: number, what: string): Uint8Array {
    const miniStream = this.loadMiniStream();
    const miniFat = (this.miniFat ??= this.loadMiniFat());
    const capacity = Math.floor(miniStream.byteLength / MINI_SECTOR_SIZE);
    const sectors = this.walk(start, miniFat, capacity, Math.ceil(size / MINI_SECTOR_SIZE), false, what);
    const out = new Uint8Array(size);
    sectors.forEach((sector, i) => {
      const offset = sector * MINI_SECTOR_SIZE;
      const length = Math.min(MINI_SECTOR_SIZE, size - i * MINI_SECTOR_SIZE);
      out.set(miniStream.subarray(offset, offset + length), i * MINI_SECTOR_SIZE);
    });
    return out;
  }

  /** Sector numbers of a chain; refuses chains that leave `count` sectors, loop, or end before `needed` sectors. */
  private walk(start: number, table: Uint32Array, count: number, needed: number, toEnd: boolean, what: string): number[] {
    const sectors: number[] = [];
    const seen = new Uint8Array(count);
    let sector = start;
    while (toEnd ? sector !== END_OF_CHAIN : sectors.length < needed) {
      if (sector === END_OF_CHAIN || sector === FREE_SECTOR) {
        throw corrupt(`${what} is shorter than its declared size`);
      }
      if (sector >= count || sector >= table.length) {
        throw corrupt(`${what} points outside the file`);
      }
      if (seen[sector]) {
        throw corrupt(`${what} loops back on itself`);
      }
      seen[sector] = 1;
      sectors.push(sector);
      sector = table[sector] ?? FREE_SECTOR;
    }
    return sectors;
  }

  private loadMiniStream(): Uint8Array {
    if (!this.miniStream) {
      const { size, start } = this.root;
      this.miniStream = size > 0 && start !== END_OF_CHAIN ? this.readChain(start, Math.min(size, this.bytes.byteLength), 'the mini stream') : new Uint8Array();
    }
    return this.miniStream;
  }

  private loadMiniFat(): Uint32Array {
    if (this.miniFatSectors === 0 || this.miniFatStart === END_OF_CHAIN) {
      return new Uint32Array();
    }
    if (this.miniFatSectors > this.sectorCount) {
      throw corrupt('the mini FAT is larger than the file');
    }
    const bytes = this.readChain(this.miniFatStart, this.miniFatSectors * this.sectorSize, 'the mini FAT');
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return Uint32Array.from({ length: bytes.byteLength / 4 }, (_, i) => view.getUint32(i * 4, true));
  }

  /** Directory entry `id`, parsed on demand; undefined past the end of the directory. */
  private entry(id: number): DirectoryEntry | undefined {
    const offset = id * DIRECTORY_ENTRY_SIZE;
    if (offset + DIRECTORY_ENTRY_SIZE > this.directory.byteLength) {
      return undefined;
    }
    const view = new DataView(this.directory.buffer, this.directory.byteOffset + offset, DIRECTORY_ENTRY_SIZE);
    const nameBytes = Math.min(view.getUint16(0x40, true), 64);
    const nameLength = Math.max(0, Math.floor(nameBytes / 2) - 1);
    let name = '';
    for (let i = 0; i < nameLength; i++) {
      name += String.fromCharCode(view.getUint16(i * 2, true));
    }
    // Version 3 writers did not always clear the high half of the size, so [MS-CFB] says to ignore it there.
    const high = this.majorVersion === 3 ? 0 : view.getUint32(0x7c, true);
    return {
      name,
      type: view.getUint8(0x42),
      left: view.getUint32(0x44, true),
      right: view.getUint32(0x48, true),
      child: view.getUint32(0x4c, true),
      start: view.getUint32(0x74, true),
      size: high > 0 ? Number.MAX_SAFE_INTEGER : view.getUint32(0x78, true),
    };
  }

  /** The entries of a storage, found by walking its red-black tree of siblings; a link to a visited entry is corruption. */
  private childrenOf(storage: DirectoryEntry): Map<string, DirectoryEntry> {
    const children = new Map<string, DirectoryEntry>();
    const seen = new Set<number>();
    const pending = [storage.child];
    while (pending.length) {
      const id = pending.pop() ?? NO_STREAM;
      if (id === NO_STREAM) {
        continue;
      }
      const entry = this.entry(id);
      if (!entry || seen.has(id)) {
        throw corrupt(entry ? 'the directory tree loops' : 'a directory entry points outside the directory');
      }
      seen.add(id);
      if (entry.type === ENTRY_STREAM || entry.type === ENTRY_STORAGE) {
        children.set(entry.name.toUpperCase(), entry);
      }
      pending.push(entry.left, entry.right);
    }
    return children;
  }

  private sectorOffset(sector: number): number {
    return (sector + 1) * this.sectorSize;
  }

  /** A little-endian uint32 at `offset`, or FREE_SECTOR past the end of a truncated file. */
  private u32At(offset: number): number {
    return offset + 4 <= this.bytes.byteLength ? this.view.getUint32(offset, true) : FREE_SECTOR;
  }
}
