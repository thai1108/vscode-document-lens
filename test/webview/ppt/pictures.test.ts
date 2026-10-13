import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { PPT_LIMITS, type PptLimits } from '../../../src/webview/ppt/limits.ts';
import { dibToBmp, imageSize, readBlip } from '../../../src/webview/ppt/pictures.ts';
import { RecordStream } from '../../../src/webview/ppt/records.ts';
import { rgbPng } from '../../support/png.ts';
import { atom, u16, u32, u8 } from '../../support/pptWriter.ts';

const UID = Buffer.alloc(16, 0xab);

function blip(type: number, instance: number, body: Buffer, limits: PptLimits = PPT_LIMITS) {
  const stream = new RecordStream(new Uint8Array(atom(type, body, instance)), { records: 10 });
  return readBlip(stream, stream.at(0), limits);
}

/** A bottom-up 2 × 2, 24-bit DIB (BITMAPINFOHEADER + rows padded to 4 bytes). */
function dib(width = 2, height = 2, bitCount = 24, colorsUsed = 0): Buffer {
  const header = Buffer.concat([u32(40), u32(width, height), u16(1, bitCount), u32(0, 0, 2835, 2835, colorsUsed, 0)]);
  return Buffer.concat([header, Buffer.alloc(16, 0x7f)]);
}

/** A baseline JPEG header up to its SOF0 frame (enough for the size). */
function jpegHeader(width: number, height: number): Buffer {
  const app0 = Buffer.concat([Buffer.from([0xff, 0xe0, 0x00, 0x10]), Buffer.from('JFIF\0', 'latin1'), Buffer.alloc(9)]);
  const sof = Buffer.from([0xff, 0xc0, 0x00, 0x0b, 0x08, height >> 8, height & 0xff, width >> 8, width & 0xff, 0x01, 0x01, 0x11, 0x00]);
  return Buffer.concat([Buffer.from([0xff, 0xd8]), app0, Buffer.from([0xff, 0xff]), sof, Buffer.from([0xff, 0xd9])]);
}

describe('pictures', () => {
  it('reads PNG and JPEG data after the UID and tag', () => {
    const png = rgbPng(3, 2, () => [1, 2, 3]);
    const picture = blip(0xf01e, 0x6e0, Buffer.concat([UID, u8(0xff), png]));
    assert.equal(picture.format, 'png');
    assert.equal(picture.mime, 'image/png');
    assert.deepEqual(picture.data, new Uint8Array(png));
    assert.deepEqual([picture.width, picture.height], [3, 2]);

    const jpeg = jpegHeader(640, 480);
    const twoUids = blip(0xf01d, 0x46b, Buffer.concat([UID, UID, u8(0xff), jpeg]));
    assert.equal(twoUids.mime, 'image/jpeg');
    assert.deepEqual([twoUids.width, twoUids.height], [640, 480]);
    assert.equal(blip(0xf02a, 0x6e2, Buffer.concat([UID, u8(0xff), jpeg])).format, 'jpeg');
  });

  it('finds the data when the instance claims the wrong number of UIDs', () => {
    const png = rgbPng(1, 1, () => [0, 0, 0]);
    assert.equal(blip(0xf01e, 0x6e1, Buffer.concat([UID, u8(0xff), png])).mime, 'image/png');
    assert.equal(blip(0xf01e, 0x6e0, Buffer.concat([UID, UID, u8(0xff), png])).mime, 'image/png');
  });

  it('turns a DIB into a BMP file the browser can decode', () => {
    const picture = blip(0xf01f, 0x7a8, Buffer.concat([UID, u8(0xff), dib()]));
    assert.equal(picture.mime, 'image/bmp');
    const bmp = Buffer.from(picture.data ?? []);
    assert.equal(bmp.toString('latin1', 0, 2), 'BM');
    assert.equal(bmp.readUInt32LE(2), bmp.length);
    assert.equal(bmp.readUInt32LE(10), 14 + 40);
    assert.deepEqual([picture.width, picture.height], [2, 2]);
  });

  it('places the BMP pixel data after the palette and bit masks', () => {
    const paletted = Buffer.from(dibToBmp(new Uint8Array(Buffer.concat([dib(2, 2, 8, 0), Buffer.alloc(1024)]))) ?? []);
    assert.equal(paletted.readUInt32LE(10), 14 + 40 + 256 * 4);
    const fewColors = Buffer.from(dibToBmp(new Uint8Array(Buffer.concat([dib(2, 2, 4, 3), Buffer.alloc(16)]))) ?? []);
    assert.equal(fewColors.readUInt32LE(10), 14 + 40 + 3 * 4);
    const bitfields = dib(2, 2, 32);
    bitfields.writeUInt32LE(3, 16);
    assert.equal(Buffer.from(dibToBmp(new Uint8Array(bitfields)) ?? []).readUInt32LE(10), 14 + 40 + 12);
    const core = Buffer.concat([u32(12), u16(2, 2, 1, 8), Buffer.alloc(256 * 3 + 8)]);
    assert.equal(Buffer.from(dibToBmp(new Uint8Array(core)) ?? []).readUInt32LE(10), 14 + 12 + 256 * 3);
    // A palette larger than the data is refused.
    assert.equal(dibToBmp(new Uint8Array(dib(2, 2, 8, 0x10000000))), undefined);
  });

  it('shows metafiles and TIFF as labelled placeholders', () => {
    assert.deepEqual(blip(0xf01a, 0x3d4, Buffer.alloc(60)), { format: 'emf', note: 'EMF picture (not shown)' });
    assert.deepEqual(blip(0xf01b, 0x216, Buffer.alloc(60)), { format: 'wmf', note: 'WMF picture (not shown)' });
    assert.deepEqual(blip(0xf01c, 0x542, Buffer.alloc(60)), { format: 'pict', note: 'PICT picture (not shown)' });
    assert.deepEqual(blip(0xf029, 0x6e4, Buffer.alloc(60)), { format: 'tiff', note: 'TIFF picture (not shown)' });
    assert.equal(blip(0xf020, 0, Buffer.alloc(60)).format, 'unknown');
  });

  it('refuses damaged data, oversized data and images too large to decode', () => {
    assert.equal(blip(0xf01e, 0x6e0, Buffer.concat([UID, u8(0xff), Buffer.from('not a png')])).note, 'Damaged picture (not shown)');
    assert.equal(blip(0xf01d, 0x46a, Buffer.concat([UID, u8(0xff), Buffer.from([0xff, 0xd8, 0xff, 0x00])])).note, 'Damaged picture (not shown)');
    const png = rgbPng(4, 4, () => [0, 0, 0]);
    assert.equal(blip(0xf01e, 0x6e0, Buffer.concat([UID, u8(0xff), png]), { ...PPT_LIMITS, maxPictureSize: 10 }).note, 'Picture too large to show');
    assert.equal(blip(0xf01e, 0x6e0, Buffer.concat([UID, u8(0xff), png]), { ...PPT_LIMITS, maxPicturePixels: 15 }).note, 'Picture too large to show (4 × 4)');
    // A tiny PNG header declaring 100000 × 100000 pixels.
    const bomb = Buffer.from(png);
    bomb.writeUInt32BE(100_000, 16);
    bomb.writeUInt32BE(100_000, 20);
    assert.equal(blip(0xf01e, 0x6e0, Buffer.concat([UID, u8(0xff), bomb])).note, 'Picture too large to show (100000 × 100000)');
  });

  it('reads image sizes from headers', () => {
    assert.deepEqual(imageSize('jpeg', new Uint8Array(jpegHeader(1, 65535))), { width: 1, height: 65535 });
    assert.equal(imageSize('jpeg', new Uint8Array([0xff, 0xd8, 0x00, 0x00, 0x00])), undefined);
    assert.equal(imageSize('jpeg', new Uint8Array([0xff, 0xd8, 0xff, 0xd9])), undefined);
    assert.equal(imageSize('png', new Uint8Array(10)), undefined);
    const core = Buffer.concat([Buffer.alloc(14), u32(12), u16(300, 200, 1, 24)]);
    assert.deepEqual(imageSize('bmp', new Uint8Array(core)), { width: 300, height: 200 });
    const topDown = Buffer.concat([Buffer.alloc(14), dib(5, -7)]);
    assert.deepEqual(imageSize('bmp', new Uint8Array(topDown)), { width: 5, height: 7 });
  });
});
