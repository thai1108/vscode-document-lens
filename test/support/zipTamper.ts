// Edits ZIP archives written by zipWriter.ts into the odd shapes a hostile file can take: patched central directory
// fields, Unicode path extra fields, deflate streams with trailing bytes, and data prepended before the archive.
import { crc32, deflateRawSync } from 'node:zlib';
import type { ZipInput } from './zipWriter.ts';

const EOCD = Buffer.from([0x50, 0x4b, 0x05, 0x06]);

export interface CentralRecord {
  /** Offset of the central directory record. */
  offset: number;
  name: string;
  localHeader: number;
}

export function centralRecords(zip: Buffer): CentralRecord[] {
  const eocd = zip.lastIndexOf(EOCD);
  const records: CentralRecord[] = [];
  for (let offset = zip.readUInt32LE(eocd + 16), i = 0; i < zip.readUInt16LE(eocd + 10); i++) {
    const nameLength = zip.readUInt16LE(offset + 28);
    records.push({ offset, name: zip.toString('utf8', offset + 46, offset + 46 + nameLength), localHeader: zip.readUInt32LE(offset + 42) });
    offset += 46 + nameLength + zip.readUInt16LE(offset + 30) + zip.readUInt16LE(offset + 32);
  }
  return records;
}

/** A copy of `zip` with `edit` applied to the central directory record (and local header) of entry `name`. */
export function patchEntry(zip: Buffer, name: string, edit: (copy: Buffer, record: CentralRecord) => void): Buffer {
  const copy = Buffer.from(zip);
  const record = centralRecords(copy).find((candidate) => candidate.name === name);
  if (!record) {
    throw new Error(`No entry ${name}`);
  }
  edit(copy, record);
  return copy;
}

/** Gives entry `name` a Unicode path extra field naming it `unicodeName` and clears its UTF-8 flag. */
export function withUnicodePath(zip: Buffer, name: string, unicodeName: string): Buffer {
  const eocd = zip.lastIndexOf(EOCD);
  const records = centralRecords(zip);
  const index = records.findIndex((record) => record.name === name);
  const record = records[index];
  if (!record) {
    throw new Error(`No entry ${name}`);
  }
  const nameLength = zip.readUInt16LE(record.offset + 28);
  const end = record.offset + 46 + nameLength + zip.readUInt16LE(record.offset + 30) + zip.readUInt16LE(record.offset + 32);
  const unicode = Buffer.from(unicodeName, 'utf8');
  const field = Buffer.alloc(9 + unicode.length);
  field.writeUInt16LE(0x7075, 0);
  field.writeUInt16LE(5 + unicode.length, 2);
  field.writeUInt8(1, 4);
  field.writeUInt32LE(crc32(zip.subarray(record.offset + 46, record.offset + 46 + nameLength)), 5);
  unicode.copy(field, 9);
  const header = Buffer.from(zip.subarray(record.offset, record.offset + 46 + nameLength));
  header.writeUInt16LE(header.readUInt16LE(8) & ~0x0800, 8);
  header.writeUInt16LE(field.length, 30);
  header.writeUInt16LE(0, 32);
  const tail = Buffer.from(zip.subarray(eocd));
  tail.writeUInt32LE(tail.readUInt32LE(12) + field.length - (end - record.offset - 46 - nameLength), 12);
  return Buffer.concat([zip.subarray(0, record.offset), header, field, zip.subarray(end, eocd), tail]);
}

/**
 * An entry holding a deflate stream of `size` zero bytes followed by a few bytes of junk. writeZip stores it as it is;
 * markDeflated then makes the archive call it deflated.
 */
export function junkDeflateEntry(name: string, size: number): ZipInput {
  return { name, data: Buffer.concat([deflateRawSync(Buffer.alloc(size)), Buffer.from('JUNK')]), store: true, declaredSize: size };
}

/** Sets entry `name`'s compression method to deflate, in both of its headers. */
export function markDeflated(zip: Buffer, name: string): Buffer {
  return patchEntry(zip, name, (copy, record) => {
    copy.writeUInt16LE(8, record.offset + 10);
    copy.writeUInt16LE(8, record.localHeader + 8);
  });
}

/**
 * Two archives in one file: `decoy` first, padded to the length of `real`'s entries, then `real`, whose end record
 * points at `decoy`'s central directory. A reader that trusts the end record's offset sees the decoy; JSZip notices
 * the gap before the end record and shifts every offset, which lands on `real`.
 */
export function prependedDecoy(decoy: Buffer, real: Buffer): Buffer {
  const split = (zip: Buffer) => {
    const eocd = zip.lastIndexOf(EOCD);
    const directory = zip.readUInt32LE(eocd + 16);
    return { entries: zip.subarray(0, directory), directory: zip.subarray(directory, eocd), count: zip.readUInt16LE(eocd + 10) };
  };
  const front = split(decoy);
  const back = split(real);
  if (front.entries.length > back.entries.length) {
    throw new Error('The decoy must be smaller than the real archive');
  }
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(back.count, 8);
  end.writeUInt16LE(back.count, 10);
  end.writeUInt32LE(back.directory.length, 12);
  end.writeUInt32LE(back.entries.length, 16);
  const padding = Buffer.alloc(back.entries.length - front.entries.length);
  return Buffer.concat([front.entries, padding, front.directory, back.entries, back.directory, end]);
}
