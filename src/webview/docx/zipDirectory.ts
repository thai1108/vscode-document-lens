// Just enough ZIP reading for the webview to check a package before docx-preview (through JSZip) unpacks it. The host
// has already validated the central directory against OOXML_ZIP_LIMITS (src/preflight.ts). This reader must see the
// archive exactly as JSZip 3.10 does (lib/zipEntries.js, zipEntry.js, load.js), so anything JSZip would read
// differently is refused rather than interpreted: data before the central directory or after the end record, ZIP64,
// names JSZip rewrites or decodes another way, duplicates, and local headers that disagree with the central directory.

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
const LOCAL_HEADER_SIGNATURE = 0x04034b50;
const EOCD_SIZE = 22;
const MAX_16 = 0xffff;
const MAX_32 = 0xffffffff;
const ENCRYPTED_FLAG = 0x0001;
const UTF8_FLAG = 0x0800;
const UNICODE_PATH_FIELD = 0x7075;
const DOS_DIRECTORY = 0x0010;
const UNIX_DIRECTORY = 0x4000;
const MADE_BY_UNIX = 3;

export const STORED = 0;
export const DEFLATED = 8;

export interface ZipEntry {
  /** The name docx-preview looks the entry up by (identical to JSZip's). */
  readonly name: string;
  readonly isDirectory: boolean;
  readonly method: typeof STORED | typeof DEFLATED;
  readonly compressedSize: number;
  /** Declared uncompressed size; inflateEntry holds the data to it. */
  readonly size: number;
  /** Where the entry's data starts, after its local header. */
  readonly dataOffset: number;
}

/** The archive is damaged, or built so that JSZip could read it differently from this reader. */
export class ZipFormatError extends Error {}

export function readZipDirectory(data: Uint8Array): ZipEntry[] {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const refuse = (reason: string) => new ZipFormatError(reason);
  const eocd = lastIndexOf(view, EOCD_SIGNATURE);
  if (eocd < 0 || eocd + EOCD_SIZE > data.byteLength) {
    throw refuse('no ZIP end of central directory');
  }
  const disk = view.getUint16(eocd + 4, true);
  const directoryDisk = view.getUint16(eocd + 6, true);
  const recordsOnDisk = view.getUint16(eocd + 8, true);
  const count = view.getUint16(eocd + 10, true);
  const directorySize = view.getUint32(eocd + 12, true);
  const directoryOffset = view.getUint32(eocd + 16, true);
  const commentLength = view.getUint16(eocd + 20, true);
  if ([disk, directoryDisk, recordsOnDisk, count].includes(MAX_16) || directorySize === MAX_32 || directoryOffset === MAX_32) {
    throw refuse('ZIP64 archives are not supported');
  }
  if (disk !== 0 || directoryDisk !== 0 || recordsOnDisk !== count) {
    throw refuse('multi-volume ZIP archives are not supported');
  }
  // JSZip shifts every offset by any gap here (data prepended to the archive); this reader does not.
  if (directoryOffset + directorySize !== eocd || eocd + EOCD_SIZE + commentLength !== data.byteLength) {
    throw refuse('unexpected data around the ZIP central directory');
  }

  const entries: ZipEntry[] = [];
  const names = new Set<string>();
  let offset = directoryOffset;
  for (let i = 0; i < count; i++) {
    if (offset + 46 > eocd || view.getUint32(offset, true) !== CENTRAL_HEADER_SIGNATURE) {
      throw refuse(`corrupt ZIP central directory at entry ${i}`);
    }
    const madeBy = view.getUint16(offset + 4, true) >> 8;
    const flags = view.getUint16(offset + 8, true);
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const size = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const externalAttributes = view.getUint32(offset + 38, true);
    const localHeader = view.getUint32(offset + 42, true);
    const nameStart = offset + 46;
    const extraStart = nameStart + nameLength;
    offset = extraStart + extraLength + commentLength;
    if (offset > eocd) {
      throw refuse(`corrupt ZIP central directory at entry ${i}`);
    }
    const nameBytes = data.subarray(nameStart, extraStart);
    const label = new TextDecoder().decode(nameBytes);
    if (flags & ENCRYPTED_FLAG) {
      throw refuse(`${label} is encrypted`);
    }
    if (method !== STORED && method !== DEFLATED) {
      throw refuse(`${label} uses compression method ${method}`);
    }
    if (compressedSize === MAX_32 || size === MAX_32 || localHeader === MAX_32) {
      throw refuse('ZIP64 archives are not supported');
    }
    if (!(flags & UTF8_FLAG) && hasExtraField(view, extraStart, extraLength, UNICODE_PATH_FIELD)) {
      // JSZip would name the entry from this field instead.
      throw refuse(`${label} has a second, Unicode name`);
    }
    const name = decodeName(nameBytes);
    if (name === undefined || name !== resolvePath(name) || name.includes('\\')) {
      throw refuse(`unusual entry name ${JSON.stringify(label)}`);
    }
    const unixMode = madeBy === MADE_BY_UNIX ? externalAttributes >>> 16 : 0;
    const isDirectory = Boolean(externalAttributes & DOS_DIRECTORY || unixMode & UNIX_DIRECTORY || name.endsWith('/'));
    if (isDirectory && !name.endsWith('/')) {
      throw refuse(`unusual entry name ${JSON.stringify(label)}`);
    }
    if (names.has(name)) {
      throw refuse(`${label} appears twice`);
    }
    names.add(name);

    // JSZip names the entry from its local header and reads the data after it.
    if (localHeader + 30 > data.byteLength || view.getUint32(localHeader, true) !== LOCAL_HEADER_SIGNATURE) {
      throw refuse(`corrupt ZIP local header for ${label}`);
    }
    const localNameLength = view.getUint16(localHeader + 26, true);
    const localExtraLength = view.getUint16(localHeader + 28, true);
    const localName = data.subarray(localHeader + 30, localHeader + 30 + localNameLength);
    if (localNameLength !== nameLength || !localName.every((byte, index) => byte === nameBytes[index])) {
      throw refuse(`${label} is named differently in its local header`);
    }
    const dataOffset = localHeader + 30 + localNameLength + localExtraLength;
    if (dataOffset + compressedSize > data.byteLength) {
      throw refuse(`${label} is truncated`);
    }
    entries.push({ name, isDirectory, method, compressedSize, size, dataOffset });
  }
  if (offset !== eocd) {
    throw refuse('unexpected data in the ZIP central directory');
  }
  return entries;
}

/** Thrown when an entry inflates to more than its header declares, i.e. the sizes the host checked were forged. */
export class ForgedSizeError extends ZipFormatError {}

/**
 * Inflates an entry chunk by chunk and holds it to its declared size: it stops with a ForgedSizeError as soon as the
 * data grows past it. Any other inflate error refuses the entry too, since the size cannot be known then.
 */
export async function inflateEntry(data: Uint8Array, entry: ZipEntry, onChunk?: (chunk: Uint8Array) => void): Promise<void> {
  const body = data.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize);
  if (entry.method === STORED) {
    if (body.byteLength > entry.size) {
      throw new ForgedSizeError(`${entry.name} is larger than its ZIP header says`);
    }
    if (body.byteLength < entry.size) {
      throw new ZipFormatError(`${entry.name} is truncated`);
    }
    onChunk?.(body);
    return;
  }
  const reader = new Blob([body.slice()]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
  let total = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      total += value.byteLength;
      if (total > entry.size) {
        await reader.cancel();
        throw new ForgedSizeError(`${entry.name} is larger than its ZIP header says`);
      }
      onChunk?.(value);
    }
  } catch (error) {
    if (error instanceof ZipFormatError) {
      throw error;
    }
    throw new ZipFormatError(`${entry.name} cannot be inflated`, { cause: error });
  }
  if (total !== entry.size) {
    throw new ZipFormatError(`${entry.name} is smaller than its ZIP header says`);
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

/** JSZip's utils.resolve: drops `.` and empty segments (except a leading or trailing one) and applies `..`. */
export function resolvePath(path: string): string {
  const parts = path.split('/');
  const result: string[] = [];
  parts.forEach((part, index) => {
    if (part === '.' || (part === '' && index !== 0 && index !== parts.length - 1)) {
      return;
    }
    if (part === '..') {
      result.pop();
    } else {
      result.push(part);
    }
  });
  return result.join('/');
}

/** Like JSZip's lastIndexOfSignature: the last occurrence anywhere in the data. */
function lastIndexOf(view: DataView, signature: number): number {
  for (let offset = view.byteLength - 4; offset >= 0; offset--) {
    if (view.getUint32(offset, true) === signature) {
      return offset;
    }
  }
  return -1;
}

/** UTF-8, as JSZip decodes names; undefined when the bytes are not valid UTF-8 (JSZip's decoder differs there). */
function decodeName(bytes: Uint8Array): string | undefined {
  try {
    return new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch {
    return undefined;
  }
}

/** Any extra field with this id (more lenient than JSZip, which skips a field header in the last four bytes). */
function hasExtraField(view: DataView, start: number, length: number, id: number): boolean {
  for (let offset = start; offset + 4 <= start + length; offset += 4 + view.getUint16(offset + 2, true)) {
    if (view.getUint16(offset, true) === id) {
      return true;
    }
  }
  return false;
}
