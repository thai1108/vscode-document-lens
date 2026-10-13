import { crc32, deflateRawSync } from 'node:zlib';

export interface ZipInput {
  name: string;
  data: Buffer | string;
  /** Store uncompressed instead of deflating. */
  store?: boolean;
  /** Lie about the uncompressed size in the headers (for hostile-archive tests). */
  declaredSize?: number;
  /** Lie about the compressed size in the headers (for hostile-archive tests). */
  declaredCompressedSize?: number;
  /** Extra field bytes, written to both headers. */
  extra?: Buffer;
  /** A different name in the local header than in the central directory (for hostile-archive tests). */
  localName?: string;
  /** Leave the UTF-8 name flag unset. */
  noUtf8Flag?: boolean;
}

export interface ZipOptions {
  /** Bytes written before the first entry without adjusting any offset, as when data is prepended to an archive. */
  prefix?: Buffer;
}

/** Minimal ZIP writer for building test archives. */
export function writeZip(files: ZipInput[], comment = '', options: ZipOptions = {}): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const file of files) {
    const raw = typeof file.data === 'string' ? Buffer.from(file.data, 'utf8') : file.data;
    const body = file.store ? raw : deflateRawSync(raw);
    const name = Buffer.from(file.name, 'utf8');
    const localName = Buffer.from(file.localName ?? file.name, 'utf8');
    const extra = file.extra ?? Buffer.alloc(0);
    const flags = file.noUtf8Flag ? 0 : 0x0800;
    const method = file.store ? 0 : 8;
    const crc = crc32(raw);
    const size = file.declaredSize ?? raw.length;
    const compressedSize = file.declaredCompressedSize ?? body.length;

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressedSize, 18);
    local.writeUInt32LE(size, 22);
    local.writeUInt16LE(localName.length, 26);
    local.writeUInt16LE(extra.length, 28);
    locals.push(local, localName, extra, body);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(flags, 8);
    central.writeUInt16LE(method, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressedSize, 20);
    central.writeUInt32LE(size, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt16LE(extra.length, 30);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, name, extra);
    offset += local.length + localName.length + extra.length + body.length;
  }
  const directory = Buffer.concat(centrals);
  const commentBytes = Buffer.from(comment, 'utf8');
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(commentBytes.length, 20);
  return Buffer.concat([options.prefix ?? Buffer.alloc(0), ...locals, directory, end, commentBytes]);
}
