import { inflateRawSync } from 'node:zlib';
import { formatMegabytes, OOXML_ZIP_LIMITS } from './limits.ts';

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
const LOCAL_HEADER_SIGNATURE = 0x04034b50;
const EOCD_MIN_SIZE = 22;
const MAX_COMMENT_SIZE = 0xffff;
const ZIP64_MARKER = 0xffffffff;
const ZIP64_COUNT_MARKER = 0xffff;

const FLAG_DATA_DESCRIPTOR = 0x0008;
const FLAG_UTF8 = 0x0800;
const UNICODE_PATH_EXTRA_FIELD = 0x7075;

const METHOD_STORED = 0;
const METHOD_DEFLATE = 8;

/** Guards against ZIP bombs: every entry is inflated to exactly its declared size, which these caps bound. */
export interface ZipLimits {
  maxEntries: number;
  maxEntrySize: number;
  maxTotalSize: number;
  /** Highest declared uncompressed/compressed ratio for entries larger than RATIO_CHECK_MIN_SIZE. */
  maxCompressionRatio: number;
}

export const DEFAULT_ZIP_LIMITS: ZipLimits = OOXML_ZIP_LIMITS;

/** Small entries (e.g. a run of spaces in an XML part) legitimately compress extremely well. */
const RATIO_CHECK_MIN_SIZE = 1024 * 1024;

interface ZipEntry {
  readonly name: string;
  readonly method: number;
  readonly compressedSize: number;
  readonly size: number;
  readonly localHeaderOffset: number;
  readonly dataOffset: number;
}

/**
 * Read-only view of a ZIP archive held in memory (stored and deflate entries, no ZIP64).
 *
 * The viewers unpack the same bytes again with their own readers (e.g. JSZip), so an archive that two readers could
 * see differently is refused: data before the first entry or behind the central directory, names a reader would
 * decode or normalize differently, duplicate names, local headers that disagree with the central directory and
 * overlapping entries.
 */
export class ZipArchive {
  private readonly buffer: Buffer;
  private readonly entries: ReadonlyMap<string, ZipEntry>;

  private constructor(buffer: Buffer, entries: ReadonlyMap<string, ZipEntry>) {
    this.buffer = buffer;
    this.entries = entries;
  }

  static open(buffer: Buffer, limits: ZipLimits = DEFAULT_ZIP_LIMITS): ZipArchive {
    const eocd = findEndOfCentralDirectory(buffer);
    const count = buffer.readUInt16LE(eocd + 10);
    const directorySize = buffer.readUInt32LE(eocd + 12);
    const directoryOffset = buffer.readUInt32LE(eocd + 16);
    if (count === ZIP64_COUNT_MARKER || directorySize === ZIP64_MARKER || directoryOffset === ZIP64_MARKER) {
      throw new Error('ZIP64 archives are not supported');
    }
    if (buffer.readUInt16LE(eocd + 4) !== 0 || buffer.readUInt16LE(eocd + 6) !== 0 || buffer.readUInt16LE(eocd + 8) !== count) {
      throw new Error('Corrupt ZIP end record (multi-disk or inconsistent entry counts)');
    }
    // JSZip shifts every offset when the directory does not end where the end record starts (data prepended to the
    // archive), so such archives would show other readers other entries.
    if (directoryOffset + directorySize !== eocd) {
      throw new Error('Corrupt ZIP archive (the central directory is not where the end record says)');
    }
    if (count > limits.maxEntries) {
      throw new Error(`ZIP archive has too many entries (${count}; limit ${limits.maxEntries})`);
    }

    const entries = new Map<string, ZipEntry>();
    let totalSize = 0;
    let offset = directoryOffset;
    for (let i = 0; i < count; i++) {
      if (offset + 46 > eocd || buffer.readUInt32LE(offset) !== CENTRAL_HEADER_SIGNATURE) {
        throw new Error(`Corrupt ZIP central directory at entry ${i}`);
      }
      const flags = buffer.readUInt16LE(offset + 8);
      const nameLength = buffer.readUInt16LE(offset + 28);
      const extraLength = buffer.readUInt16LE(offset + 30);
      const commentLength = buffer.readUInt16LE(offset + 32);
      const end = offset + 46 + nameLength + extraLength + commentLength;
      if (end > eocd) {
        throw new Error(`Corrupt ZIP central directory at entry ${i}`);
      }
      const rawName = buffer.subarray(offset + 46, offset + 46 + nameLength);
      const name = entryName(rawName, flags, buffer.subarray(offset + 46 + nameLength, offset + 46 + nameLength + extraLength));
      if (entries.has(name)) {
        throw new Error(`ZIP archive has two entries named ${name}`);
      }
      const method = buffer.readUInt16LE(offset + 10);
      const compressedSize = buffer.readUInt32LE(offset + 20);
      const size = buffer.readUInt32LE(offset + 24);
      const localHeaderOffset = buffer.readUInt32LE(offset + 42);
      if (compressedSize === ZIP64_MARKER || size === ZIP64_MARKER || localHeaderOffset === ZIP64_MARKER) {
        throw new Error('ZIP64 archives are not supported');
      }
      if (size > limits.maxEntrySize) {
        throw new Error(`ZIP entry too large: ${name} (${formatMegabytes(size)}; limit ${formatMegabytes(limits.maxEntrySize)})`);
      }
      if (size > RATIO_CHECK_MIN_SIZE && size > compressedSize * limits.maxCompressionRatio) {
        throw new Error(`ZIP entry compressed suspiciously well: ${name} (more than ${limits.maxCompressionRatio}:1)`);
      }
      totalSize += size;
      if (totalSize > limits.maxTotalSize) {
        throw new Error(`ZIP archive too large when unpacked (limit ${formatMegabytes(limits.maxTotalSize)})`);
      }
      const dataOffset = localDataOffset(buffer, { name, method, compressedSize, size, localHeaderOffset }, rawName, flags);
      if (dataOffset + compressedSize > directoryOffset) {
        throw new Error(`Corrupt ZIP entry (truncated data): ${name}`);
      }
      entries.set(name, { name, method, compressedSize, size, localHeaderOffset, dataOffset });
      offset = end;
    }
    if (offset !== eocd) {
      throw new Error('Corrupt ZIP central directory (its size does not match its entries)');
    }
    rejectOverlaps([...entries.values()]);
    return new ZipArchive(buffer, entries);
  }

  has(name: string): boolean {
    return this.entries.has(name);
  }

  names(): string[] {
    return [...this.entries.keys()];
  }

  read(name: string): Buffer {
    const entry = this.entries.get(name);
    if (!entry) {
      throw new Error(`ZIP entry not found: ${name}`);
    }
    const data = this.buffer.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize);
    let output: Buffer;
    switch (entry.method) {
      case METHOD_STORED:
        output = data;
        break;
      case METHOD_DEFLATE:
        output = inflateExactly(data, entry);
        break;
      default:
        throw new Error(`Unsupported ZIP compression method ${entry.method}: ${name}`);
    }
    if (output.length !== entry.size) {
      throw new Error(`Corrupt ZIP entry (size mismatch): ${name}`);
    }
    return output;
  }

  readText(name: string): string {
    return this.read(name).toString('utf8');
  }
}

/** Inflates at most the declared size; a stream that would produce more is rejected instead of filling memory. */
function inflateExactly(data: Buffer, entry: ZipEntry): Buffer {
  try {
    return inflateRawSync(data, { maxOutputLength: Math.max(1, entry.size) });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ERR_BUFFER_TOO_LARGE') {
      throw new Error(`Corrupt ZIP entry (inflates beyond its declared size): ${entry.name}`, { cause: error });
    }
    throw error;
  }
}

function findEndOfCentralDirectory(buffer: Buffer): number {
  const lowest = Math.max(0, buffer.length - EOCD_MIN_SIZE - MAX_COMMENT_SIZE);
  for (let offset = buffer.length - EOCD_MIN_SIZE; offset >= lowest; offset--) {
    if (buffer.readUInt32LE(offset) === EOCD_SIGNATURE) {
      // The archive comment runs to the end of the file; anything else after the end record is not part of the archive.
      if (offset + EOCD_MIN_SIZE + buffer.readUInt16LE(offset + 20) !== buffer.length) {
        throw new Error('Corrupt ZIP end record (data after the archive comment)');
      }
      return offset;
    }
  }
  throw new Error('Not a ZIP archive (end of central directory not found)');
}

/**
 * The entry name, refusing names that readers decode or normalize differently: JSZip resolves "." and ".." segments
 * and repeated slashes and prefers a Unicode Path extra field; names that are not UTF-8 are decoded per reader.
 */
function entryName(raw: Buffer, flags: number, extra: Buffer): string {
  const ascii = raw.every((byte) => byte < 0x80);
  const name = raw.toString('utf8');
  if (!ascii && (!(flags & FLAG_UTF8) || !Buffer.from(name, 'utf8').equals(raw))) {
    throw new Error(`ZIP entry name is not valid UTF-8: ${name}`);
  }
  const segments = name.split('/');
  const canonical =
    name !== '' &&
    !/[\\\0]/.test(name) &&
    segments.every((segment, i) => (segment !== '' || (i > 0 && i === segments.length - 1)) && segment !== '.' && segment !== '..');
  if (!canonical) {
    throw new Error(`ZIP entry name is not canonical: ${name}`);
  }
  const unicodePath = extraField(extra, UNICODE_PATH_EXTRA_FIELD, name);
  // Version (1 byte) and the CRC-32 of the header name (4 bytes) precede the UTF-8 name.
  if (unicodePath && unicodePath.subarray(5).toString('utf8') !== name) {
    throw new Error(`ZIP entry has a second, different name: ${name}`);
  }
  return name;
}

function extraField(extra: Buffer, id: number, name: string): Buffer | undefined {
  for (let offset = 0; offset < extra.length; ) {
    if (offset + 4 > extra.length || offset + 4 + extra.readUInt16LE(offset + 2) > extra.length) {
      throw new Error(`Corrupt ZIP extra field: ${name}`);
    }
    const size = extra.readUInt16LE(offset + 2);
    if (extra.readUInt16LE(offset) === id) {
      return extra.subarray(offset + 4, offset + 4 + size);
    }
    offset += 4 + size;
  }
  return undefined;
}

/** Where the entry's data starts, after checking that its local header agrees with the central directory. */
function localDataOffset(buffer: Buffer, entry: Omit<ZipEntry, 'dataOffset'>, rawName: Buffer, flags: number): number {
  const header = entry.localHeaderOffset;
  if (header + 30 > buffer.length || buffer.readUInt32LE(header) !== LOCAL_HEADER_SIGNATURE) {
    throw new Error(`Corrupt ZIP local header: ${entry.name}`);
  }
  const nameLength = buffer.readUInt16LE(header + 26);
  const extraLength = buffer.readUInt16LE(header + 28);
  // With a data descriptor the local header may leave the sizes zero; they follow the data instead.
  const sizesAgree = flags & FLAG_DATA_DESCRIPTOR || (buffer.readUInt32LE(header + 18) === entry.compressedSize && buffer.readUInt32LE(header + 22) === entry.size);
  if (!buffer.subarray(header + 30, header + 30 + nameLength).equals(rawName) || buffer.readUInt16LE(header + 8) !== entry.method || !sizesAgree) {
    throw new Error(`ZIP local header disagrees with the central directory: ${entry.name}`);
  }
  return header + 30 + nameLength + extraLength;
}

/** Entries sharing bytes let a small archive declare far more data than it holds (overlapping "ZIP bombs"). */
function rejectOverlaps(entries: ZipEntry[]): void {
  const ordered = entries.toSorted((a, b) => a.localHeaderOffset - b.localHeaderOffset);
  for (let i = 1; i < ordered.length; i++) {
    const previous = ordered[i - 1];
    const next = ordered[i];
    if (previous && next && previous.dataOffset + previous.compressedSize > next.localHeaderOffset) {
      throw new Error(`ZIP entries overlap: ${previous.name} and ${next.name}`);
    }
  }
}
