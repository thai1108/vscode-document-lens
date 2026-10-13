// Read-only ZIP access inside the webview, for .xlsx packages. The host has already checked the central directory
// (src/preflight.ts), but the viewer checks it again against the same limits and inflates every entry to exactly its
// declared size with the browser's own DecompressionStream, so a forged header cannot make it unpack more.
import { formatMegabytes, MAX_XML_PART_SIZE, OOXML_ZIP_LIMITS } from '../../../limits.ts';

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
const LOCAL_HEADER_SIGNATURE = 0x04034b50;
const EOCD_MIN_SIZE = 22;
const MAX_COMMENT_SIZE = 0xffff;
const ZIP64_MARKER = 0xffffffff;
const RATIO_CHECK_MIN_SIZE = 1024 * 1024;
const FLAG_ENCRYPTED = 0x0001;
/** Compressed bytes handed to the inflater at a time. */
const INPUT_CHUNK = 64 * 1024;
/** Longest stretch of reading before the page gets a turn to paint and handle input. */
const MAX_BUSY_MS = 30;

interface ZipEntry {
  readonly name: string;
  readonly flags: number;
  readonly method: number;
  readonly compressedSize: number;
  readonly size: number;
  readonly localHeaderOffset: number;
}

export class ZipReader {
  private readonly bytes: Uint8Array;
  private readonly view: DataView;
  /** By lower-case name: OPC part names are case-insensitive and writers are not consistent about case. */
  private readonly entries: ReadonlyMap<string, ZipEntry>;

  private constructor(bytes: Uint8Array, entries: ReadonlyMap<string, ZipEntry>) {
    this.bytes = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.entries = entries;
  }

  static open(bytes: Uint8Array, limits = OOXML_ZIP_LIMITS): ZipReader {
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
    let total = 0;
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
      const entry: ZipEntry = {
        name: decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength)),
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
        throw new Error(`ZIP entry too large: ${entry.name} (${formatMegabytes(entry.size)}; limit ${formatMegabytes(limits.maxEntrySize)})`);
      }
      if (entry.size > RATIO_CHECK_MIN_SIZE && entry.size > entry.compressedSize * limits.maxCompressionRatio) {
        throw new Error(`ZIP entry compressed suspiciously well: ${entry.name} (more than ${limits.maxCompressionRatio}:1)`);
      }
      total += entry.size;
      if (total > limits.maxTotalSize) {
        throw new Error(`ZIP archive too large when unpacked (limit ${formatMegabytes(limits.maxTotalSize)})`);
      }
      entries.set(entry.name.toLowerCase(), entry);
      offset += 46 + nameLength + extraLength + commentLength;
    }
    return new ZipReader(bytes, entries);
  }

  has(name: string): boolean {
    return this.entries.has(name.toLowerCase());
  }

  /** Uncompressed size, or undefined for a missing entry. */
  size(name: string): number | undefined {
    return this.entries.get(name.toLowerCase())?.size;
  }

  /** The whole entry as text; refused above `maxSize` (by default the cap on XML parts read in one piece). */
  async readText(name: string, maxSize = MAX_XML_PART_SIZE): Promise<string> {
    const entry = this.entry(name);
    if (entry.size > maxSize) {
      throw new Error(`${entry.name} is too large to read (${formatMegabytes(entry.size)}; limit ${formatMegabytes(maxSize)})`);
    }
    let text = '';
    for await (const chunk of this.textChunks(name)) {
      text += chunk;
    }
    return text;
  }

  /** The entry as text in chunks, for parts too large to hold as one string. */
  async *textChunks(name: string, signal?: AbortSignal): AsyncGenerator<string> {
    let decoder: TextDecoder | undefined;
    for await (const chunk of this.chunks(name)) {
      signal?.throwIfAborted();
      decoder ??= new TextDecoder(chunk[0] === 0xff && chunk[1] === 0xfe ? 'utf-16le' : chunk[0] === 0xfe && chunk[1] === 0xff ? 'utf-16be' : 'utf-8');
      const text = decoder.decode(chunk, { stream: true });
      if (text) {
        yield text;
      }
    }
    const tail = decoder?.decode() ?? '';
    if (tail) {
      yield tail;
    }
  }

  /** The entry's bytes in chunks, inflated to exactly the declared size. */
  async *chunks(name: string): AsyncGenerator<Uint8Array> {
    const entry = this.entry(name);
    const data = this.compressedData(entry);
    if (entry.method === 0) {
      if (data.byteLength !== entry.size) {
        throw new Error(`Corrupt ZIP entry (size mismatch): ${entry.name}`);
      }
      yield data;
      return;
    }
    if (entry.method !== 8) {
      throw new Error(`Unsupported ZIP compression method ${entry.method}: ${entry.name}`);
    }
    // Fed in small pieces on demand, so only what the reader has asked for is inflated: a whole part handed over at
    // once would be inflated at once, all of it in memory and without the page getting a turn in between.
    let offset = 0;
    const input = new ReadableStream<Uint8Array<ArrayBuffer>>(
      {
        pull(controller) {
          const end = Math.min(data.byteLength, offset + INPUT_CHUNK);
          // The webview's bytes always sit in an ArrayBuffer (never a SharedArrayBuffer).
          controller.enqueue(data.subarray(offset, end) as Uint8Array<ArrayBuffer>);
          offset = end;
          if (offset >= data.byteLength) {
            controller.close();
          }
        },
      },
      { highWaterMark: 1 },
    );
    const reader = input.pipeThrough(new DecompressionStream('deflate-raw')).getReader();
    let total = 0;
    let lastPause = Date.now();
    try {
      for (;;) {
        let result: ReadableStreamReadResult<Uint8Array>;
        try {
          result = await reader.read();
        } catch (error) {
          throw new Error(`Corrupt ZIP entry (bad compressed data): ${entry.name}`, { cause: error });
        }
        if (result.done) {
          break;
        }
        total += result.value.byteLength;
        if (total > entry.size) {
          throw new Error(`Corrupt ZIP entry (inflates beyond its declared size): ${entry.name}`);
        }
        yield result.value;
        // Stream reads resolve as microtasks; without a real pause, reading a large part would freeze the page.
        if (Date.now() - lastPause > MAX_BUSY_MS) {
          await new Promise((resolve) => setTimeout(resolve, 0));
          lastPause = Date.now();
        }
      }
    } finally {
      reader.cancel().catch(() => undefined);
    }
    if (total !== entry.size) {
      throw new Error(`Corrupt ZIP entry (size mismatch): ${entry.name}`);
    }
  }

  private entry(name: string): ZipEntry {
    const entry = this.entries.get(name.toLowerCase());
    if (!entry) {
      throw new Error(`ZIP entry not found: ${name}`);
    }
    if (entry.flags & FLAG_ENCRYPTED) {
      throw new Error(`ZIP entry is encrypted: ${entry.name}`);
    }
    return entry;
  }

  private compressedData(entry: ZipEntry): Uint8Array {
    const header = entry.localHeaderOffset;
    if (header + 30 > this.view.byteLength || this.view.getUint32(header, true) !== LOCAL_HEADER_SIGNATURE) {
      throw new Error(`Corrupt ZIP local header: ${entry.name}`);
    }
    const start = header + 30 + this.view.getUint16(header + 26, true) + this.view.getUint16(header + 28, true);
    if (start + entry.compressedSize > this.view.byteLength) {
      throw new Error(`Corrupt ZIP entry (truncated data): ${entry.name}`);
    }
    return this.bytes.subarray(start, start + entry.compressedSize);
  }
}

export function isZip(bytes: Uint8Array): boolean {
  return bytes[0] === 0x50 && bytes[1] === 0x4b && (bytes[2] === 0x03 || bytes[2] === 0x05) && (bytes[3] === 0x04 || bytes[3] === 0x06);
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
