// Read-only ZIP access in the webview. The host already checked the central directory (src/preflight.ts), but the
// viewer re-applies the same limits before trusting any size, and inflates each entry lazily with the browser's own
// DecompressionStream, stopping as soon as the output would exceed the size the directory declared.
import { formatMegabytes, OOXML_ZIP_LIMITS } from '../../limits.ts';

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
const LOCAL_HEADER_SIGNATURE = 0x04034b50;
const EOCD_MIN_SIZE = 22;
const MAX_COMMENT_SIZE = 0xffff;
const ZIP64_MARKER = 0xffffffff;
const METHOD_STORED = 0;
const METHOD_DEFLATE = 8;
const FLAG_ENCRYPTED = 0x1;
/** Small entries (e.g. a run of spaces in an XML part) legitimately compress extremely well. */
const RATIO_CHECK_MIN_SIZE = 1024 * 1024;

export interface ZipLimits {
  maxEntries: number;
  maxEntrySize: number;
  maxTotalSize: number;
  maxCompressionRatio: number;
}

interface ZipEntry {
  readonly name: string;
  readonly method: number;
  readonly flags: number;
  readonly compressedSize: number;
  readonly size: number;
  readonly localHeaderOffset: number;
}

export class ZipReader {
  private readonly bytes: Uint8Array;
  private readonly view: DataView;
  private readonly entries: ReadonlyMap<string, ZipEntry>;

  private constructor(bytes: Uint8Array, entries: ReadonlyMap<string, ZipEntry>) {
    this.bytes = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.entries = entries;
  }

  static open(bytes: Uint8Array, limits: ZipLimits = OOXML_ZIP_LIMITS): ZipReader {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const eocd = findEndOfCentralDirectory(view);
    const count = view.getUint16(eocd + 10, true);
    const directoryOffset = view.getUint32(eocd + 16, true);
    if (directoryOffset === ZIP64_MARKER) {
      throw new Error('ZIP64 archives are not supported');
    }
    if (count > limits.maxEntries) {
      throw new Error(`ZIP archive has too many entries (${count}; limit ${limits.maxEntries})`);
    }
    const decoder = new TextDecoder();
    const entries = new Map<string, ZipEntry>();
    let totalSize = 0;
    let offset = directoryOffset;
    for (let i = 0; i < count; i++) {
      if (offset + 46 > view.byteLength || view.getUint32(offset, true) !== CENTRAL_HEADER_SIGNATURE) {
        throw new Error(`Corrupt ZIP central directory at entry ${i}`);
      }
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      if (offset + 46 + nameLength > view.byteLength) {
        throw new Error(`Corrupt ZIP central directory at entry ${i}`);
      }
      const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
      const entry: ZipEntry = {
        name,
        flags: view.getUint16(offset + 8, true),
        method: view.getUint16(offset + 10, true),
        compressedSize: view.getUint32(offset + 20, true),
        size: view.getUint32(offset + 24, true),
        localHeaderOffset: view.getUint32(offset + 42, true),
      };
      if (entry.compressedSize === ZIP64_MARKER || entry.size === ZIP64_MARKER || entry.localHeaderOffset === ZIP64_MARKER) {
        throw new Error('ZIP64 archives are not supported');
      }
      if (entry.size > limits.maxEntrySize) {
        throw new Error(`ZIP entry too large: ${name} (${formatMegabytes(entry.size)}; limit ${formatMegabytes(limits.maxEntrySize)})`);
      }
      if (entry.size > RATIO_CHECK_MIN_SIZE && entry.size > entry.compressedSize * limits.maxCompressionRatio) {
        throw new Error(`ZIP entry compressed suspiciously well: ${name} (more than ${limits.maxCompressionRatio}:1)`);
      }
      totalSize += entry.size;
      if (totalSize > limits.maxTotalSize) {
        throw new Error(`ZIP archive too large when unpacked (limit ${formatMegabytes(limits.maxTotalSize)})`);
      }
      entries.set(name, entry);
      offset += 46 + nameLength + extraLength + commentLength;
    }
    return new ZipReader(bytes, entries);
  }

  has(name: string): boolean {
    return this.entries.has(name);
  }

  names(): string[] {
    return [...this.entries.keys()];
  }

  /** Declared uncompressed size, without inflating. */
  size(name: string): number | undefined {
    return this.entries.get(name)?.size;
  }

  async read(name: string): Promise<Uint8Array> {
    const entry = this.entries.get(name);
    if (!entry) {
      throw new Error(`ZIP entry not found: ${name}`);
    }
    if (entry.flags & FLAG_ENCRYPTED) {
      throw new Error(`Encrypted ZIP entry: ${name}`);
    }
    const header = entry.localHeaderOffset;
    if (header + 30 > this.view.byteLength || this.view.getUint32(header, true) !== LOCAL_HEADER_SIGNATURE) {
      throw new Error(`Corrupt ZIP local header: ${name}`);
    }
    const start = header + 30 + this.view.getUint16(header + 26, true) + this.view.getUint16(header + 28, true);
    if (start + entry.compressedSize > this.view.byteLength) {
      throw new Error(`Corrupt ZIP entry (truncated data): ${name}`);
    }
    const data = this.bytes.subarray(start, start + entry.compressedSize);
    switch (entry.method) {
      case METHOD_STORED:
        if (data.byteLength !== entry.size) {
          throw new Error(`Corrupt ZIP entry (size mismatch): ${name}`);
        }
        return data;
      case METHOD_DEFLATE:
        return inflateExactly(data, entry);
      default:
        throw new Error(`Unsupported ZIP compression method ${entry.method}: ${name}`);
    }
  }
}

/** Inflates into a buffer of exactly the declared size; a stream that would produce more is cancelled. */
async function inflateExactly(data: Uint8Array, entry: ZipEntry): Promise<Uint8Array> {
  const output = new Uint8Array(entry.size);
  let length = 0;
  const reader = new Blob([data as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      if (length + value.byteLength > entry.size) {
        void reader.cancel().catch(() => undefined);
        throw new Error(`Corrupt ZIP entry (inflates beyond its declared size): ${entry.name}`);
      }
      output.set(value, length);
      length += value.byteLength;
    }
  } catch (error) {
    if ((error as Error).message.startsWith('Corrupt ZIP entry')) {
      throw error;
    }
    throw new Error(`Corrupt ZIP entry (bad deflate data): ${entry.name}`, { cause: error });
  }
  if (length !== entry.size) {
    throw new Error(`Corrupt ZIP entry (size mismatch): ${entry.name}`);
  }
  return output;
}

function findEndOfCentralDirectory(view: DataView): number {
  const lowest = Math.max(0, view.byteLength - EOCD_MIN_SIZE - MAX_COMMENT_SIZE);
  for (let offset = view.byteLength - EOCD_MIN_SIZE; offset >= lowest; offset--) {
    if (view.getUint32(offset, true) === EOCD_SIGNATURE) {
      return offset;
    }
  }
  throw new Error('Not a ZIP archive (end of central directory not found)');
}
