import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { shareMedia } from '../../../src/webview/docx/media.ts';

/** A stand-in for docx-preview's WordDocument, counting how often it inflates a part and makes a URL. */
function fakeDocument() {
  const calls = { load: 0, url: 0 };
  const document = {
    _package: {
      load(path: string, type = 'string'): Promise<unknown> {
        calls.load++;
        return Promise.resolve(type === 'blob' ? new Blob([path]) : type === 'uint8array' ? new TextEncoder().encode(path) : path);
      },
    },
    blobToURL(blob: Blob | null): string | null {
      calls.url++;
      return blob ? `blob:https://webview.test/${calls.url}` : null;
    },
  };
  return { document, calls };
}

describe('docx media loads', () => {
  it('inflates each picture once and gives it one URL, however often it is used', async () => {
    const { document, calls } = fakeDocument();
    const media = shareMedia(document, new Map([['word/media/image1.png', 4 * 1024 * 1024]]));
    const toUrl = document.blobToURL as (blob: Blob | null, path: string) => string;
    // As WordDocument.loadDocumentImage does for every reference: load the part, then make a URL for its path.
    const urls = await Promise.all(
      Array.from({ length: 250 }, async () => toUrl((await document._package.load('word/media/image1.png', 'blob')) as Blob, 'word/media/image1.png')),
    );
    assert.equal(new Set(urls).size, 1);
    assert.equal(calls.load, 1);
    assert.equal(calls.url, 1);
    assert.deepEqual(media.objectUrls, [urls[0]]);
    assert.equal(media.overBudget, false);
  });

  it('gives every use of a font its own copy of the bytes', async () => {
    const { document } = fakeDocument();
    shareMedia(document, new Map([['word/fonts/font1.odttf', 10]]));
    const first = (await document._package.load('word/fonts/font1.odttf', 'uint8array')) as Uint8Array;
    first[0] = 0;
    const second = (await document._package.load('word/fonts/font1.odttf', 'uint8array')) as Uint8Array;
    assert.notEqual(second[0], 0);
  });

  it('leaves out pictures beyond the budget and says so', async () => {
    const { document } = fakeDocument();
    const media = shareMedia(document, new Map([['a.png', 60], ['b.png', 60]]), 100);
    assert.ok(await document._package.load('a.png', 'blob'));
    assert.ok(await document._package.load('a.png', 'blob'), 'a picture already loaded stays available');
    assert.equal(await document._package.load('b.png', 'blob'), null);
    assert.equal(media.overBudget, true);
    assert.equal(await document._package.load('word/document.xml'), 'word/document.xml', 'XML parts are not affected');
  });

  it('refuses a docx-preview without the loader it relies on', () => {
    assert.throws(() => shareMedia({}, new Map()), /no longer exposes the package loader/);
  });
});
