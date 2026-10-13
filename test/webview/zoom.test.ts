import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { clampZoom, formatZoom, MAX_ZOOM, MIN_ZOOM, stepZoom } from '../../src/webview/shell/zoom.ts';

describe('zoom steps', () => {
  it('steps through presets from any zoom, staying within bounds', () => {
    assert.equal(stepZoom(1, 1), 1.1);
    assert.equal(stepZoom(1, -1), 0.9);
    assert.equal(stepZoom(0.87, 1), 0.9);
    assert.equal(stepZoom(0.87, -1), 0.8);
    assert.equal(stepZoom(MAX_ZOOM, 1), MAX_ZOOM);
    assert.equal(stepZoom(MIN_ZOOM, -1), MIN_ZOOM);
  });

  it('clamps and formats', () => {
    assert.equal(clampZoom(100), MAX_ZOOM);
    assert.equal(clampZoom(0), MIN_ZOOM);
    assert.equal(clampZoom(Number.NaN), 1);
    assert.equal(formatZoom(0.333), '33%');
  });
});
