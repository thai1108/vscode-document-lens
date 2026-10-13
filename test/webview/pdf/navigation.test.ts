import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { fitWidthZoom, parsePageNumber, parseViewState } from '../../../src/webview/pdf/navigation.ts';

describe('parseViewState', () => {
  it('keeps a well-formed saved position', () => {
    assert.deepEqual(parseViewState({ page: 3, left: 0, top: 595.5, outline: true }), { page: 3, left: 0, top: 595.5, outline: true });
    assert.deepEqual(parseViewState({ page: 1 }), { page: 1 });
  });

  it('drops a position without both coordinates and ignores unknown fields', () => {
    assert.deepEqual(parseViewState({ page: 2, top: 10 }), { page: 2 });
    assert.deepEqual(parseViewState({ page: 2, left: 5, top: Number.POSITIVE_INFINITY, outline: 'yes', extra: 1 }), { page: 2 });
  });

  it('refuses state without a valid page number', () => {
    for (const raw of [undefined, null, 'page 3', 3, [], {}, { page: 0 }, { page: -1 }, { page: 1.5 }, { page: '3' }, { page: Number.NaN }, { page: 2 ** 60 }]) {
      assert.equal(parseViewState(raw), undefined, JSON.stringify(raw));
    }
  });
});

describe('parsePageNumber', () => {
  it('reads a plain number and clamps it to the document', () => {
    assert.equal(parsePageNumber('3', 10), 3);
    assert.equal(parsePageNumber(' 7 ', 10), 7);
    assert.equal(parsePageNumber('0', 10), 1);
    assert.equal(parsePageNumber('999999999', 10), 10);
  });

  it('refuses anything else', () => {
    for (const text of ['', ' ', 'abc', '-2', '1.5', '1e2', '0x10', '3 4', '1234567890']) {
      assert.equal(parsePageNumber(text, 10), null, text);
    }
    assert.equal(parsePageNumber('1', 0), null);
  });
});

describe('fitWidthZoom', () => {
  it('fits the page between the gutters', () => {
    assert.equal(fitWidthZoom(1000, 800, 12), 976 / 800);
    assert.equal(fitWidthZoom(424, 400, 12), 1);
  });

  it('gives up without room or without a page size', () => {
    assert.equal(fitWidthZoom(24, 800, 12), undefined);
    assert.equal(fitWidthZoom(0, 800, 12), undefined);
    assert.equal(fitWidthZoom(1000, 0, 12), undefined);
    assert.equal(fitWidthZoom(1000, Number.NaN, 12), undefined);
  });
});
