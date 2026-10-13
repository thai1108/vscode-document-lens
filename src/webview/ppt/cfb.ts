// Read-only view of an OLE compound file ([MS-CFB]) held in memory: enough to read the streams in the root storage.
// Every sector number, chain and size comes from the file, so each is checked against the file's real extent; sector
// chains that loop or leave the file are reported instead of followed.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const HEADER_SIZE = 512;
const HEADER_DIFAT_ENTRIES = 109;
const DIRECTORY_ENTRY_SIZE = 128;
const END_OF_CHAIN = 0xfffffffe;
const NO_STREAM = 0xffffffff;
const ENTRY_STREAM = 2;
const ENTRY_ROOT = 5;
// [MS-CFB] fixes both; trusting other values would let a tiny mini sector turn a stream into millions of sectors.
const MINI_SECTOR_SHIFT = 6;
const MINI_STREAM_CUTOFF = 4096;

export class CompoundFileError extends Error {}

interface DirectoryEntry {
  readonly name: string;
  readonly type: number;
  readonly left: number;
  readonly right: number;
  readonly child: number;
  readonly start: number;
  readonly size: number;
}

export interface CompoundFileLimits {
  maxDirectoryEntries: number;
}

export class CompoundFile {
  private readonly bytes: Uint8Array;
  private readonly sectorSize: number;
  /** Sectors actually present in the file (a partial last sector counts). */
  private readonly sectorCount: number;
  private readonly fat: Uint32Array;
  private readonly firstMiniFatSector: number;
  private readonly miniFatSectorCount: number;
  private readonly entries: readonly DirectoryEntry[];
  private mini: { fat: Uint32Array; stream: Uint8Array } | undefined;

  private constructor(bytes: Uint8Array, limits: CompoundFileLimits) {
    this.bytes = bytes;
    const header = new DataView(bytes.buffer, bytes.byteOffset, HEADER_SIZE);
    const u16 = (offset: number) => header.getUint16(offset, true);
    const u32 = (offset: number) => header.getUint32(offset, true);
    if (u16(28) !== 0xfffe) {
      throw new CompoundFileError('compound file has a wrong byte order mark');
    }
    const sectorShift = u16(30);
    if (sectorShift !== 9 && sectorShift !== 12) {
      throw new CompoundFileError(`compound file has an unsupported sector size (2^${sectorShift})`);
    }
    this.sectorSize = 1 << sectorShift;
    this.sectorCount = Math.ceil(Math.max(0, bytes.byteLength - this.sectorSize) / this.sectorSize);
    if (u16(32) !== MINI_SECTOR_SHIFT || u32(56) !== MINI_STREAM_CUTOFF) {
      throw new CompoundFileError('compound file has an invalid mini stream layout');
    }
    this.firstMiniFatSector = u32(60);
    this.miniFatSectorCount = u32(64);
    this.fat = this.readFat(u32(44), u32(68), u32(72), header);

    const directory = this.readSectors(this.fat, u32(48), limits.maxDirectoryEntries * DIRECTORY_ENTRY_SIZE, this.regular());
    const count = Math.floor(directory.byteLength / DIRECTORY_ENTRY_SIZE);
    const view = new DataView(directory.buffer, directory.byteOffset, directory.byteLength);
    const wideSizes = u16(26) >= 4;
    const entries: DirectoryEntry[] = [];
    for (let i = 0; i < count; i++) {
      entries.push(readDirectoryEntry(directory, view, i * DIRECTORY_ENTRY_SIZE, wideSizes));
    }
    if (entries[0]?.type !== ENTRY_ROOT) {
      throw new CompoundFileError('compound file has no root directory entry');
    }
    this.entries = entries;
  }

  static isCompoundFile(bytes: Uint8Array): boolean {
    return bytes.byteLength >= HEADER_SIZE && SIGNATURE.every((byte, i) => bytes[i] === byte);
  }

  static open(bytes: Uint8Array, limits: CompoundFileLimits): CompoundFile {
    if (!CompoundFile.isCompoundFile(bytes)) {
      throw new CompoundFileError('no compound file header');
    }
    return new CompoundFile(bytes, limits);
  }

  /** Names of the streams directly inside the root storage. */
  streamNames(): string[] {
    return this.rootChildren()
      .filter((entry) => entry.type === ENTRY_STREAM)
      .map((entry) => entry.name);
  }

  has(name: string): boolean {
    return this.find(name) !== undefined;
  }

  /** The stream's bytes (a view into the file when its sectors are contiguous), or undefined when there is none. */
  stream(name: string, maxSize = Infinity): Uint8Array | undefined {
    const entry = this.find(name);
    if (!entry) {
      return undefined;
    }
    // A stream cannot hold more bytes than the file, whatever its directory entry claims.
    const size = Math.min(entry.size, this.bytes.byteLength);
    if (size > maxSize) {
      throw new CompoundFileError(`the ${name} stream is too large`);
    }
    if (size >= MINI_STREAM_CUTOFF) {
      return this.readSectors(this.fat, entry.start, size, this.regular());
    }
    const mini = this.miniStream();
    const unit = 1 << MINI_SECTOR_SHIFT;
    return this.readSectors(mini.fat, entry.start, size, { bytes: mini.stream, first: 0, unit, count: Math.floor(mini.stream.byteLength / unit) });
  }

  /** The file's own sectors; sector 0 starts right after the header sector. */
  private regular(): SectorSource {
    return { bytes: this.bytes, first: this.sectorSize, unit: this.sectorSize, count: this.sectorCount };
  }

  private find(name: string): DirectoryEntry | undefined {
    const wanted = name.toUpperCase();
    return this.rootChildren().find((entry) => entry.type === ENTRY_STREAM && entry.name.toUpperCase() === wanted);
  }

  /** The root storage's children: a binary tree threaded through left/right sibling ids, walked with cycle checks. */
  private rootChildren(): DirectoryEntry[] {
    const found: DirectoryEntry[] = [];
    const seen = new Set<number>();
    const pending = [this.entries[0]?.child ?? NO_STREAM];
    while (pending.length) {
      const id = pending.pop() ?? NO_STREAM;
      const entry = this.entries[id];
      if (!entry || seen.has(id)) {
        continue;
      }
      seen.add(id);
      found.push(entry);
      pending.push(entry.left, entry.right);
    }
    return found;
  }

  private miniStream(): { fat: Uint32Array; stream: Uint8Array } {
    if (!this.mini) {
      const root = this.entries[0];
      const fat = this.readSectors(this.fat, this.firstMiniFatSector, this.miniFatSectorCount * this.sectorSize, this.regular());
      const stream = root ? this.readSectors(this.fat, root.start, Math.min(root.size, this.bytes.byteLength), this.regular()) : new Uint8Array();
      this.mini = { fat: uint32Table(fat), stream };
    }
    return this.mini;
  }

  /** FAT sector numbers come from the 109 in the header, then from the DIFAT sector chain. */
  private readFat(fatSectorCount: number, firstDifatSector: number, difatSectorCount: number, header: DataView): Uint32Array {
    if (fatSectorCount > this.sectorCount) {
      throw new CompoundFileError('compound file declares more FAT sectors than the file holds');
    }
    // FAT entries past the file's last sector could only lead outside the file, so only the FAT sectors that map the
    // file are read.
    const perFatSector = this.sectorSize / 4;
    const needed = Math.min(fatSectorCount, Math.ceil(this.sectorCount / perFatSector));
    const fatSectors: number[] = [];
    for (let i = 0; i < HEADER_DIFAT_ENTRIES && fatSectors.length < needed; i++) {
      fatSectors.push(header.getUint32(76 + i * 4, true));
    }
    const perDifatSector = perFatSector - 1;
    const seen = new Set<number>();
    let difat = firstDifatSector;
    for (let read = 0; fatSectors.length < needed; read++) {
      if (read >= difatSectorCount || difat >= this.sectorCount) {
        throw new CompoundFileError('compound file DIFAT chain ends early');
      }
      if (seen.has(difat)) {
        throw new CompoundFileError('compound file DIFAT chain loops');
      }
      seen.add(difat);
      const entries = uint32Table(this.sector(difat));
      fatSectors.push(...entries.subarray(0, Math.min(perDifatSector, needed - fatSectors.length)));
      difat = entries[perDifatSector] ?? END_OF_CHAIN;
    }
    const fat = new Uint32Array(needed * perFatSector);
    fatSectors.forEach((sector, i) => fat.set(uint32Table(this.sector(sector)), i * perFatSector));
    return fat;
  }

  /** Sector `n` of the file, which starts right after the header sector. */
  private sector(n: number): Uint8Array {
    if (n >= this.sectorCount) {
      throw new CompoundFileError('compound file sector lies outside the file');
    }
    const offset = (n + 1) * this.sectorSize;
    const bytes = this.bytes.subarray(offset, offset + this.sectorSize);
    if (bytes.byteLength === this.sectorSize) {
      return bytes;
    }
    // A truncated last sector reads as if padded with zeros.
    const padded = new Uint8Array(this.sectorSize);
    padded.set(bytes);
    return padded;
  }

  /**
   * Follows a sector chain from `start` until `size` bytes are read or the chain ends (a chain shorter than its
   * stream's declared size gives a truncated stream). A sector that repeats means the chain loops.
   */
  private readSectors(table: Uint32Array, start: number, size: number, source: SectorSource): Uint8Array {
    const wanted = Math.ceil(size / source.unit);
    const chain: number[] = [];
    const seen = new Uint8Array(table.length);
    for (let sector = start; chain.length < wanted && sector !== END_OF_CHAIN; sector = table[sector] ?? END_OF_CHAIN) {
      if (sector >= table.length) {
        throw new CompoundFileError('compound file sector chain is broken');
      }
      if (sector >= source.count) {
        throw new CompoundFileError('compound file sector lies outside the file');
      }
      if (seen[sector]) {
        throw new CompoundFileError('compound file sector chain loops');
      }
      seen[sector] = 1;
      chain.push(sector);
    }
    const length = Math.min(size, chain.length * source.unit);
    const offsetOf = (sector: number) => source.first + sector * source.unit;
    // Sectors that follow each other need no copy.
    const first = chain[0];
    if (first !== undefined && isContiguous(chain) && offsetOf(first) + length <= source.bytes.byteLength) {
      return source.bytes.subarray(offsetOf(first), offsetOf(first) + length);
    }
    // A truncated last sector of the file reads as if padded with zeros.
    const out = new Uint8Array(length);
    chain.forEach((sector, i) => {
      const offset = offsetOf(sector);
      out.set(source.bytes.subarray(offset, offset + Math.min(source.unit, length - i * source.unit)), i * source.unit);
    });
    return out;
  }
}

/** Where sector n of a chain lives: `bytes[first + n * unit ...]`, for n below `count`. */
interface SectorSource {
  readonly bytes: Uint8Array;
  readonly first: number;
  readonly unit: number;
  readonly count: number;
}

function isContiguous(chain: readonly number[]): boolean {
  return chain.every((sector, i) => i === 0 || sector === (chain[i - 1] ?? -2) + 1);
}

function uint32Table(bytes: Uint8Array): Uint32Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const table = new Uint32Array(Math.floor(bytes.byteLength / 4));
  for (let i = 0; i < table.length; i++) {
    table[i] = view.getUint32(i * 4, true);
  }
  return table;
}

function readDirectoryEntry(bytes: Uint8Array, view: DataView, offset: number, wideSizes: boolean): DirectoryEntry {
  // The name length counts the terminating NUL; a value beyond the 64-byte field is corrupt and gets clamped.
  const nameLength = Math.min(Math.max(view.getUint16(offset + 64, true) - 2, 0), 62) & ~1;
  const sizeLow = view.getUint32(offset + 120, true);
  // Version 3 files may leave garbage in the high half of the size.
  const sizeHigh = wideSizes ? view.getUint32(offset + 124, true) : 0;
  return {
    name: new TextDecoder('utf-16le').decode(bytes.subarray(offset, offset + nameLength)),
    type: view.getUint8(offset + 66),
    left: view.getUint32(offset + 68, true),
    right: view.getUint32(offset + 72, true),
    child: view.getUint32(offset + 76, true),
    start: view.getUint32(offset + 116, true),
    size: sizeHigh * 2 ** 32 + sizeLow,
  };
}
