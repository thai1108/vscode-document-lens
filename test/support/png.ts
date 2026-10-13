import { crc32, deflateSync } from 'node:zlib';

const SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/** A truecolor-with-alpha PNG of `width` × `height`, each pixel given by `pixel(x, y)` as [r, g, b, a]. */
export function png(width: number, height: number, pixel: (x: number, y: number) => readonly [number, number, number, number]): Buffer {
  return encode(width, height, 6, 4, pixel);
}

/** An 8-bit RGB PNG whose pixels come from `color(x, y)`. */
export function rgbPng(width: number, height: number, color: (x: number, y: number) => readonly [number, number, number]): Buffer {
  return encode(width, height, 2, 3, color);
}

/** A truecolor PNG whose pixel (x, y) is `color(x, y)` as [r, g, b]. */
export const encodePng = rgbPng;

/** 64×48 picture in four quadrants: red, green (top) and blue, yellow (bottom), so crops are easy to see. */
export function quadrantPng(): Buffer {
  return rgbPng(64, 48, (x, y) => (y < 24 ? (x < 32 ? [255, 0, 0] : [0, 255, 0]) : x < 32 ? [0, 0, 255] : [255, 255, 0]));
}

function encode(width: number, height: number, colorType: number, channels: number, pixel: (x: number, y: number) => readonly number[]): Buffer {
  const stride = width * channels + 1;
  const raw = Buffer.alloc(stride * height);
  for (let y = 0; y < height; y++) {
    // Each scanline starts with its filter type; 0 = none.
    raw[y * stride] = 0;
    for (let x = 0; x < width; x++) {
      raw.set(pixel(x, y), y * stride + 1 + x * channels);
    }
  }
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  // Bit depth 8, the colour type, default compression, filtering and no interlacing.
  header.set([8, colorType, 0, 0, 0], 8);
  return Buffer.concat([SIGNATURE, chunk('IHDR', header), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

function chunk(type: string, data: Buffer): Buffer {
  const typeAndData = Buffer.concat([Buffer.from(type, 'latin1'), data]);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(typeAndData));
  return Buffer.concat([length, typeAndData, crc]);
}
