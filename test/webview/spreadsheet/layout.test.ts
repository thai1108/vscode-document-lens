// Grid geometry: positions along an axis, merged ranges near a window, and which rows/columns are on screen.
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { Axis } from '../../../src/webview/spreadsheet/view/axis.ts';
import { MergeIndex } from '../../../src/webview/spreadsheet/view/merges.ts';
import { canvasExtent, contains, frozenExtent, position, scrollingSpan, scrollToShow, type AxisView } from '../../../src/webview/spreadsheet/view/viewport.ts';

// Entries 0-2 are 20px, entry 3 is hidden, entry 4 is 50px, the rest default to 20px.
const axis = new Axis({ defaultSize: 20, sizes: new Map([[3, 0], [4, 50]]) }, 1_048_576);

describe('Axis', () => {
  it('positions entries with their own sizes, then default sizes for ever', () => {
    assert.deepEqual([0, 1, 3, 4, 5, 6].map((i) => axis.start(i)), [0, 20, 60, 60, 110, 130]);
    assert.equal(axis.size(3), 0);
    assert.equal(axis.size(4), 50);
    assert.equal(axis.start(1_000_000), 110 + (1_000_000 - 5) * 20);
  });

  it('finds the entry at an offset, skipping hidden ones', () => {
    assert.equal(axis.indexAt(0, 100), 0);
    assert.equal(axis.indexAt(59.9, 100), 2);
    assert.equal(axis.indexAt(60, 100), 4, 'the hidden entry 3 is never the one shown');
    assert.equal(axis.indexAt(109, 100), 4);
    assert.equal(axis.indexAt(110, 100), 5);
    assert.equal(axis.indexAt(-5, 100), 0);
    assert.equal(axis.indexAt(1e9, 100), 99);
    assert.equal(axis.countUntil(0, 60, 100), 3);
    assert.equal(axis.countUntil(4, 50, 100), 1);
  });

  it('ignores sizes past the limit and handles everything hidden', () => {
    const limited = new Axis({ defaultSize: 10, sizes: new Map([[5_000_000, 99]]) }, 1000);
    assert.equal(limited.start(10), 100);
    const hidden = new Axis({ defaultSize: 0, sizes: new Map() }, 1000);
    assert.equal(hidden.indexAt(50, 10), 0);
  });
});

describe('MergeIndex', () => {
  const merges = [
    { top: 0, left: 0, bottom: 0, right: 3 },
    { top: 100, left: 2, bottom: 200, right: 4 },
    { top: 10, left: 0, bottom: 1_000_000, right: 0 },
  ];
  const index = new MergeIndex(merges);

  it('finds merges overlapping a window, including very tall ones', () => {
    assert.deepEqual(index.query(0, 5, 0, 10), [merges[0]]);
    assert.deepEqual(index.query(150, 160, 3, 3), [merges[1]]);
    assert.deepEqual(index.query(150, 160, 0, 10), [merges[1], merges[2]]);
    assert.deepEqual(index.query(500_000, 500_001, 0, 0), [merges[2]]);
    assert.deepEqual(index.query(150, 160, 5, 10), []);
  });

  it('finds the merge containing a cell', () => {
    assert.equal(index.at(0, 2), merges[0]);
    assert.equal(index.at(1, 2), undefined);
  });
});

describe('viewport', () => {
  const view: AxisView = { axis: new Axis({ defaultSize: 20, sizes: new Map() }, 1_048_576), count: 1000, frozen: 2, header: 20, zoom: 1 };

  it('places entries after the header, scaled by zoom', () => {
    assert.equal(position(view, 0), 20);
    assert.equal(position(view, 3), 80);
    assert.equal(frozenExtent(view), 60);
    assert.equal(canvasExtent(view), 20 + 20_000);
    assert.equal(position({ ...view, zoom: 1.5, header: 30 }, 3), 120);
  });

  it('finds the scrolling entries on screen, after the frozen ones', () => {
    assert.deepEqual(scrollingSpan(view, 0, 200), { first: 2, last: 8 });
    assert.deepEqual(scrollingSpan(view, 400, 200), { first: 22, last: 28 });
    assert.deepEqual(scrollingSpan(view, 400, 200, 100), { first: 17, last: 33 });
    // At 200% the frozen part is 120px: rows 12 and 13 (canvas 520-600) are on screen.
    assert.deepEqual(scrollingSpan({ ...view, zoom: 2, header: 40 }, 400, 200), { first: 12, last: 13 });
    assert.deepEqual(scrollingSpan(view, 1e9, 200), { first: 999, last: 999 });
    const tiny = scrollingSpan(view, 0, 50);
    assert.ok(tiny.last < tiny.first, 'a viewport smaller than the frozen part shows no scrolling entries');
  });

  it('knows when the drawn span still covers what is on screen', () => {
    assert.ok(contains({ first: 10, last: 50 }, { first: 20, last: 30 }));
    assert.ok(!contains({ first: 10, last: 50 }, { first: 45, last: 55 }));
    assert.ok(contains({ first: 10, last: 50 }, { first: 5, last: 4 }));
  });

  it('scrolls just enough to show an entry, never for frozen ones', () => {
    assert.equal(scrollToShow(view, 1, 300, 200), 300);
    assert.equal(scrollToShow(view, 5, 0, 200), 0);
    assert.equal(scrollToShow(view, 20, 0, 200), 240, 'the entry ends at the bottom of the viewport');
    assert.equal(scrollToShow(view, 10, 300, 200), 160, 'the entry starts right under the frozen part');
  });
});
