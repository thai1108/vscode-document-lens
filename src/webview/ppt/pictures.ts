// Pictures ([MS-ODRAW] OfficeArtBlip records): PNG and JPEG data are shown as they are, a DIB gets the BMP file header
// browsers need, and metafiles (EMF/WMF/PICT) and TIFF become labelled placeholders. Image dimensions are read from
// the data first so a tiny file declaring a gigantic image never reaches the browser's decoder.
import type { PptLimits } from './limits.ts';
import { RT, type PptRecord, type RecordStream } from './records.ts';

export type PictureFormat = 'png' | 'jpeg' | 'bmp' | 'emf' | 'wmf' | 'pict' | 'tiff' | 'unknown';

export interface Picture {
  readonly format: PictureFormat;
  /** Image data the browser can show (with `mime`); absent when the picture is shown as a placeholder. */
  readonly data?: Uint8Array;
  readonly mime?: string;
  readonly width?: number;
  readonly height?: number;
  /** Why the picture is not shown, for its placeholder. */
  readonly note?: string;
}

const UID_SIZE = 16;
const BITMAP_TAG_SIZE = 1;
const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const DIB_HEADER_SIZES = new Set([12, 40, 52, 56, 64, 108, 124]);
const BMP_FILE_HEADER_SIZE = 14;
const MIME: Partial<Record<PictureFormat, string>> = { png: 'image/png', jpeg: 'image/jpeg', bmp: 'image/bmp' };
const METAFILES: Partial<Record<number, PictureFormat>> = { [RT.BlipEmf]: 'emf', [RT.BlipWmf]: 'wmf', [RT.BlipPict]: 'pict' };

export function isBlip(record: PptRecord): boolean {
  return record.type >= RT.BlipFirst && record.type <= RT.BlipLast;
}

export function placeholder(format: PictureFormat, note: string): Picture {
  return { format, note };
}

/** The picture held by an OfficeArtBlip record. */
export function readBlip(stream: RecordStream, blip: PptRecord, limits: PptLimits): Picture {
  const metafile = METAFILES[blip.type];
  if (metafile) {
    return placeholder(metafile, `${metafile.toUpperCase()} picture (not shown)`);
  }
  if (blip.type === RT.BlipTiff) {
    return placeholder('tiff', 'TIFF picture (not shown)');
  }
  const format: PictureFormat | undefined =
    blip.type === RT.BlipPng ? 'png' : blip.type === RT.BlipJpeg || blip.type === RT.BlipJpegCmyk ? 'jpeg' : blip.type === RT.BlipDib ? 'bmp' : undefined;
  if (!format) {
    return placeholder('unknown', 'Picture in an unknown format (not shown)');
  }
  // One or two 16-byte UIDs (an odd instance means two), then a tag byte. Some writers get the instance wrong, so the
  // other position is tried when the expected one does not start with the format's signature.
  const uids = blip.instance & 1 ? 2 : 1;
  const candidates = [uids, 3 - uids].map((count) => stream.bytes.subarray(blip.start + count * UID_SIZE + BITMAP_TAG_SIZE, blip.end));
  const raw = candidates.find((data) => hasSignature(format, data));
  if (!raw) {
    return placeholder(format, 'Damaged picture (not shown)');
  }
  if (raw.byteLength > limits.maxPictureSize) {
    return placeholder(format, 'Picture too large to show');
  }
  const data = format === 'bmp' ? dibToBmp(raw) : raw;
  const size = data && imageSize(format, data);
  if (!data || !size) {
    return placeholder(format, 'Damaged picture (not shown)');
  }
  if (size.width * size.height > limits.maxPicturePixels) {
    return placeholder(format, `Picture too large to show (${size.width} × ${size.height})`);
  }
  return { format, data, mime: MIME[format], width: size.width, height: size.height };
}

function hasSignature(format: PictureFormat, data: Uint8Array): boolean {
  switch (format) {
    case 'png':
      return PNG_SIGNATURE.every((byte, i) => data[i] === byte);
    case 'jpeg':
      return data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff;
    case 'bmp':
      return data.byteLength >= 16 && DIB_HEADER_SIZES.has(new DataView(data.buffer, data.byteOffset, 4).getUint32(0, true));
    default:
      return false;
  }
}

/** Prepends the BITMAPFILEHEADER a DIB lacks, so the browser decodes it as a .bmp file. */
export function dibToBmp(dib: Uint8Array): Uint8Array | undefined {
  const view = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const headerSize = view.getUint32(0, true);
  let paletteEntrySize = 4;
  let bitCount: number;
  let compression = 0;
  let colorsUsed = 0;
  if (headerSize === 12) {
    bitCount = view.getUint16(10, true);
    paletteEntrySize = 3;
  } else if (headerSize >= 40 && dib.byteLength >= 40) {
    bitCount = view.getUint16(14, true);
    compression = view.getUint32(16, true);
    colorsUsed = view.getUint32(32, true);
  } else {
    return undefined;
  }
  const colors = colorsUsed || (bitCount >= 1 && bitCount <= 8 ? 1 << bitCount : 0);
  // BI_BITFIELDS / BI_ALPHABITFIELDS keep their channel masks after a plain 40-byte header.
  const masks = headerSize === 40 ? (compression === 3 ? 12 : compression === 6 ? 16 : 0) : 0;
  const pixelOffset = BMP_FILE_HEADER_SIZE + headerSize + masks + colors * paletteEntrySize;
  if (pixelOffset > BMP_FILE_HEADER_SIZE + dib.byteLength) {
    return undefined;
  }
  const bmp = new Uint8Array(BMP_FILE_HEADER_SIZE + dib.byteLength);
  const header = new DataView(bmp.buffer);
  bmp[0] = 0x42;
  bmp[1] = 0x4d;
  header.setUint32(2, bmp.byteLength, true);
  header.setUint32(10, pixelOffset, true);
  bmp.set(dib, BMP_FILE_HEADER_SIZE);
  return bmp;
}

/** Pixel dimensions from the image header: PNG IHDR, JPEG SOFn, BMP info header. */
export function imageSize(format: PictureFormat, data: Uint8Array): { width: number; height: number } | undefined {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  switch (format) {
    case 'png':
      return data.byteLength >= 24 ? { width: view.getUint32(16), height: view.getUint32(20) } : undefined;
    case 'bmp': {
      const dib = BMP_FILE_HEADER_SIZE;
      if (view.getUint32(dib, true) === 12) {
        return { width: view.getUint16(dib + 4, true), height: view.getUint16(dib + 6, true) };
      }
      // A negative height means a top-down bitmap.
      return { width: Math.abs(view.getInt32(dib + 4, true)), height: Math.abs(view.getInt32(dib + 8, true)) };
    }
    case 'jpeg':
      return jpegSize(data, view);
    default:
      return undefined;
  }
}

function jpegSize(data: Uint8Array, view: DataView): { width: number; height: number } | undefined {
  let position = 2;
  while (position + 4 <= data.byteLength) {
    if (data[position] !== 0xff) {
      return undefined;
    }
    const marker = data[position + 1] ?? 0;
    if (marker === 0xff) {
      position += 1;
      continue;
    }
    // Markers without a length: TEM, RST0–7, SOI, EOI.
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9)) {
      position += 2;
      continue;
    }
    const isFrame = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isFrame) {
      return position + 9 <= data.byteLength ? { width: view.getUint16(position + 7), height: view.getUint16(position + 5) } : undefined;
    }
    position += 2 + view.getUint16(position + 2);
  }
  return undefined;
}
