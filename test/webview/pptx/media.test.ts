import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import { PPTX_LIMITS } from '../../../src/webview/pptx/limits.ts';
import { imageSize, MediaStore, sniffImage } from '../../../src/webview/pptx/media.ts';
import { PptxPackage } from '../../../src/webview/pptx/package.ts';
import { encodePng } from '../../support/png.ts';
import { writeZip } from '../../support/zipWriter.ts';

const bytes = (...values: number[]) => new Uint8Array(values);
const emf = () => {
  const header = new Uint8Array(64);
  header.set([0x01, 0, 0, 0]);
  header.set([0x20, 0x45, 0x4d, 0x46], 40);
  return header;
};

describe('pictures', () => {
  afterEach(() => {
    PPTX_LIMITS.maxPictureBytes = 128 * 1024 * 1024;
    PPTX_LIMITS.maxMediaBytes = 512 * 1024 * 1024;
  });

  it('recognizes formats by their bytes, not their names', () => {
    assert.equal(sniffImage(encodePng(1, 1, () => [0, 0, 0])), 'png');
    assert.equal(sniffImage(bytes(0xff, 0xd8, 0xff, 0xe0)), 'jpeg');
    assert.equal(sniffImage(new TextEncoder().encode('GIF89a')), 'gif');
    assert.equal(sniffImage(bytes(0x42, 0x4d, 0, 0)), 'bmp');
    assert.equal(sniffImage(new TextEncoder().encode('RIFF\0\0\0\0WEBPVP8 ')), 'webp');
    assert.equal(sniffImage(emf()), 'emf');
    assert.equal(sniffImage(bytes(0xd7, 0xcd, 0xc6, 0x9a)), 'wmf');
    assert.equal(sniffImage(bytes(0x49, 0x49, 0x2a, 0)), 'tiff');
    assert.equal(sniffImage(new TextEncoder().encode('<?xml version="1.0"?>\n<!-- logo -->\n<svg xmlns="http://www.w3.org/2000/svg"/>')), 'svg');
    assert.equal(sniffImage(new TextEncoder().encode('<html><svg/></html>')), 'unknown');
  });

  it('reads picture sizes from their headers', () => {
    assert.deepEqual(imageSize(encodePng(64, 48, () => [0, 0, 0]), 'png'), { width: 64, height: 48 });
    assert.deepEqual(imageSize(new Uint8Array([...new TextEncoder().encode('GIF89a'), 3, 0, 2, 0]), 'gif'), { width: 3, height: 2 });
    const jpeg = bytes(0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 17, 8, 0, 20, 0, 30, 3);
    assert.deepEqual(imageSize(jpeg, 'jpeg'), { width: 30, height: 20 });
    assert.deepEqual(imageSize(bytes(0x89, 0x50), 'png'), {});
  });

  it('hands out blob: URLs for browser formats, labels for the rest, and revokes them', async () => {
    const pkg = PptxPackage.open(new Uint8Array(writeZip([
      { name: 'ppt/media/a.png', data: encodePng(2, 2, () => [255, 0, 0]) },
      { name: 'ppt/media/b.emf', data: Buffer.from(emf()) },
      { name: 'ppt/media/c.png', data: Buffer.from('not a picture') },
    ])));
    const media = new MediaStore(pkg);
    const png = await media.picture('ppt/media/a.png');
    assert.ok('url' in png && png.url.startsWith('blob:'));
    assert.equal(await media.picture('ppt/media/a.png'), png, 'each picture is decoded once');
    assert.deepEqual(await media.picture('ppt/media/b.emf'), { unsupported: 'EMF picture' });
    assert.deepEqual(await media.picture('ppt/media/c.png'), { unsupported: 'Picture' });
    assert.deepEqual(await media.picture('ppt/media/missing.png'), { unsupported: 'Missing picture' });
    const revoked: string[] = [];
    const original = URL.revokeObjectURL;
    URL.revokeObjectURL = (url: string) => revoked.push(url);
    try {
      media.dispose();
    } finally {
      URL.revokeObjectURL = original;
    }
    assert.deepEqual(revoked, ['url' in png ? png.url : '']);
  });

  it('stops creating pictures past the per-picture and total byte budgets', async () => {
    const picture = encodePng(16, 16, (x) => [x * 16, 0, 0]);
    const pkg = PptxPackage.open(new Uint8Array(writeZip([
      { name: 'one.png', data: picture },
      { name: 'two.png', data: picture },
    ])));
    PPTX_LIMITS.maxPictureBytes = picture.length - 1;
    assert.deepEqual(await new MediaStore(pkg).picture('one.png'), { unsupported: 'Picture not shown (too large)' });
    PPTX_LIMITS.maxPictureBytes = picture.length;
    PPTX_LIMITS.maxMediaBytes = picture.length * 1.5;
    const media = new MediaStore(pkg);
    assert.ok('url' in (await media.picture('one.png')));
    assert.deepEqual(await media.picture('two.png'), { unsupported: 'Picture not shown (too large)' });
    media.dispose();
  });
});
