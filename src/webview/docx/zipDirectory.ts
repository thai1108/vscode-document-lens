// Just enough ZIP reading for the webview to check a package before docx-preview (through JSZip) unpacks it. The host
// has already validated the central directory against OOXML_ZIP_LIMITS (src/preflight.ts); this reads it again for
// names and sizes, and inflates entries as streams that stop as soon as one outgrows its declared size.

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
const LOCAL_HEADER_SIGNATURE = 0x04034b50;
const EOCD_MIN_SIZE = 22;
const MAX_COMMENT_SIZE = 0xffff;
const ZIP64_MARKER = 0xffffffff;
const UTF8_FLAG = 0x0800;
const UNICODE_PATH_FIELD = 0x7075;

export interface ZipEntry {
  /** The entry name decoded the way JSZip (and so docx-preview) decodes it. */
  readonly name: string;
  /**
   * JSZip may name this entry differently: a non-ASCII name that is not flagged as UTF-8 and carries a Unicode path
   * field (which JSZip prefers) or is not valid UTF-8.
   */
  readonly ambiguousName: boolean;
  readonly method: number;
  readonly compressedSize: number;
  /** Declared uncompressed size. */
  readonly size: number;
  readonly localHeaderOffset: number;
}

export function readZipDirectory(data: Uint8Array): ZipEntry[] {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const eocd = findEndOfCentralDirectory(view);
  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  if (offset === ZIP64_MARKER) {
    throw new Error('ZIP64 archives are not supported');
  }
  const entries: ZipEntry[] = [];
  for (let i = 0; i < count; i++) {
    if (offset + 46 > data.byteLength || view.getUint32(offset, true) !== CENTRAL_HEADER_SIGNATURE) {
      throw new Error(`Corrupt ZIP central directory at entry ${i}`);
    }
    const flags = view.getUint16(offset + 8, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const nameStart = offset + 46;
    const extraStart = nameStart + nameLength;
    if (extraStart + extraLength + commentLength > data.byteLength) {
      throw new Error(`Corrupt ZIP central directory at entry ${i}`);
    }
    const nameBytes = data.subarray(nameStart, extraStart);
    const name = decodeName(nameBytes);
    const ascii = nameBytes.every((byte) => byte < 0x80);
    const ambiguousName =
      !ascii && (name === undefined || ((flags & UTF8_FLAG) === 0 && hasExtraField(view, extraStart, extraLength, UNICODE_PATH_FIELD)));
    entries.push({
      name: name ?? '',
      ambiguousName,
      method: view.getUint16(offset + 10, true),
      compressedSize: view.getUint32(offset + 20, true),
      size: view.getUint32(offset + 24, true),
      localHeaderOffset: view.getUint32(offset + 42, true),
    });
    offset = extraStart + extraLength + commentLength;
  }
  return entries;
}

/** Thrown when an entry inflates to more than its header declares, i.e. the sizes the host checked were forged. */
export class ForgedSizeError extends Error {}

/**
 * Inflates an entry chunk by chunk, stopping with a ForgedSizeError as soon as it grows past its declared size, so
 * the declared sizes the caps are checked against can be trusted. Other inflate errors are passed on as they are.
 */
export async function inflateEntry(data: Uint8Array, entry: ZipEntry, onChunk?: (chunk: Uint8Array) => void): Promise<void> {
  const body = entryData(data, entry);
  if (entry.method === 0) {
    if (body.byteLength > entry.size) {
      throw new ForgedSizeError(`ZIP entry larger than declared: ${entry.name}`);
    }
    onChunk?.(body);
    return;
  }
  if (entry.method !== 8) {
    throw new Error(`Unsupported ZIP compression method ${entry.method}: ${entry.name}`);
  }
  const reader = new Blob([body.slice()]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      return;
    }
    total += value.byteLength;
    if (total > entry.size) {
      await reader.cancel();
      throw new ForgedSizeError(`ZIP entry larger than declared: ${entry.name}`);
    }
    onChunk?.(value);
  }
}

/** Inflates an entry and decodes it as UTF-8. */
export async function readZipText(data: Uint8Array, entry: ZipEntry): Promise<string> {
  const decoder = new TextDecoder();
  let text = '';
  await inflateEntry(data, entry, (chunk) => {
    text += decoder.decode(chunk, { stream: true });
  });
  return text + decoder.decode();
}

function entryData(data: Uint8Array, entry: ZipEntry): Uint8Array {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const header = entry.localHeaderOffset;
  if (header + 30 > data.byteLength || view.getUint32(header, true) !== LOCAL_HEADER_SIGNATURE) {
    throw new Error(`Corrupt ZIP local header: ${entry.name}`);
  }
  const start = header + 30 + view.getUint16(header + 26, true) + view.getUint16(header + 28, true);
  if (start + entry.compressedSize > data.byteLength) {
    throw new Error(`Corrupt ZIP entry (truncated data): ${entry.name}`);
  }
  return data.subarray(start, start + entry.compressedSize);
}

function findEndOfCentralDirectory(view: DataView): number {
  const last = view.byteLength - EOCD_MIN_SIZE;
  const first = Math.max(0, last - MAX_COMMENT_SIZE);
  for (let offset = last; offset >= first; offset--) {
    if (view.getUint32(offset, true) === EOCD_SIGNATURE) {
      return offset;
    }
  }
  throw new Error('Not a ZIP archive (end of central directory not found)');
}

function decodeName(bytes: Uint8Array): string | undefined {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return undefined;
  }
}

function hasExtraField(view: DataView, start: number, length: number, id: number): boolean {
  for (let offset = start; offset + 4 <= start + length; offset += 4 + view.getUint16(offset + 2, true)) {
    if (view.getUint16(offset, true) === id) {
      return true;
    }
  }
  return false;
}
