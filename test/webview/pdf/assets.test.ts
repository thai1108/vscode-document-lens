import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { decodeBase64, embeddedDataFactory } from '../../../src/webview/pdf/assets.ts';

const files = {
  cMapUrl: { 'UniJIS-UCS2-H.bcmap': Buffer.from([1, 2, 3]).toString('base64') },
  standardFontDataUrl: { 'FoxitDingbats.pfb': Buffer.from('font').toString('base64') },
  wasmUrl: { 'jbig2.wasm': Buffer.from([0, 0x61, 0x73, 0x6d]).toString('base64') },
};

/** Constructs and calls the factory the way pdf.js does (getDocument's BinaryDataFactory option). */
function fetchFile(kind: string, filename: string): Promise<Uint8Array> {
  const Factory = embeddedDataFactory(files);
  return new Factory().fetch({ kind, filename });
}

describe('embeddedDataFactory', () => {
  it('serves the embedded CMaps, fonts and wasm files as bytes', async () => {
    assert.deepEqual(await fetchFile('cMapUrl', 'UniJIS-UCS2-H.bcmap'), new Uint8Array([1, 2, 3]));
    assert.deepEqual(await fetchFile('standardFontDataUrl', 'FoxitDingbats.pfb'), new TextEncoder().encode('font'));
    assert.deepEqual(await fetchFile('wasmUrl', 'jbig2.wasm'), new Uint8Array([0, 0x61, 0x73, 0x6d]));
  });

  it('refuses files it does not have, including names that look up the prototype', async () => {
    await assert.rejects(fetchFile('cMapUrl', 'Missing-H.bcmap'), /No embedded cMapUrl file named Missing-H\.bcmap/);
    await assert.rejects(fetchFile('iccUrl', 'profile.icc'), /No embedded iccUrl file/);
    for (const name of ['__proto__', 'constructor', 'toString', 'hasOwnProperty']) {
      await assert.rejects(fetchFile('cMapUrl', name), /No embedded/);
      await assert.rejects(fetchFile(name, 'UniJIS-UCS2-H.bcmap'), /No embedded/);
    }
  });
});

describe('decodeBase64', () => {
  it('round-trips every byte value', () => {
    const bytes = Uint8Array.from({ length: 256 }, (_, i) => i);
    assert.deepEqual(decodeBase64(Buffer.from(bytes).toString('base64')), bytes);
    assert.deepEqual(decodeBase64(''), new Uint8Array());
  });
});
