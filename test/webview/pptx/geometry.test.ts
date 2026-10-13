import assert from 'node:assert/strict';
import { afterEach, before, describe, it } from 'node:test';
import { arcSegments, buildGeometry, calculate, customGeometry, evaluateGuides, formulasOf, presetGeometry, type Formula } from '../../../src/webview/pptx/geometry.ts';
import { PPTX_LIMITS } from '../../../src/webview/pptx/limits.ts';
import { POLYGON_CORNERS, PRESET_ALIASES, PRESETS, STAR_POINTS } from '../../../src/webview/pptx/presets.ts';
import { mirror } from '../../../src/webview/pptx/shapes.ts';
import { emuToPx, mapIntoGroup, readGroupXfrm, readXfrm, type Xfrm } from '../../../src/webview/pptx/units.ts';
import { A_NS, installDom, xml } from './dom.ts';

const PX = 9525;
const formula = (op: string, ...args: (string | number)[]): Formula => ({ name: 'x', op, args: args.map(String) });
const values = new Map([
  ['a', 3],
  ['b', 4],
]);

describe('guide formulas', () => {
  it('evaluates every DrawingML operator', () => {
    assert.equal(calculate(formula('*/', 'a', 10, 2), values), 15);
    assert.equal(calculate(formula('*/', 1, 1, 0), values), 0, 'division by zero gives 0');
    assert.equal(calculate(formula('+-', 'a', 'b', 1), values), 6);
    assert.equal(calculate(formula('+/', 'a', 'b', 7), values), 1);
    assert.equal(calculate(formula('?:', 'a', 1, 2), values), 1);
    assert.equal(calculate(formula('?:', 0, 1, 2), values), 2);
    assert.equal(calculate(formula('abs', -5), values), 5);
    assert.equal(calculate(formula('at2', 1, 1), values), 45 * 60000);
    assert.ok(Math.abs(calculate(formula('cat2', 10, 1, 0), values) - 10) < 1e-9);
    assert.ok(Math.abs(calculate(formula('cos', 10, 60 * 60000), values) - 5) < 1e-9);
    assert.equal(calculate(formula('max', 'a', 'b'), values), 4);
    assert.equal(calculate(formula('min', 'a', 'b'), values), 3);
    assert.equal(calculate(formula('mod', 'a', 'b', 0), values), 5);
    assert.equal(calculate(formula('pin', 0, -5, 10), values), 0);
    assert.equal(calculate(formula('pin', 0, 15, 10), values), 10);
    assert.equal(calculate(formula('pin', 0, 5, 10), values), 5);
    assert.ok(Math.abs(calculate(formula('sat2', 10, 0, 1), values) - 10) < 1e-9);
    assert.ok(Math.abs(calculate(formula('sin', 10, 30 * 60000), values) - 5) < 1e-9);
    assert.equal(calculate(formula('sqrt', 16), values), 4);
    assert.ok(Math.abs(calculate(formula('tan', 10, 45 * 60000), values) - 10) < 1e-9);
    assert.equal(calculate(formula('val', 'b'), values), 4);
    assert.equal(calculate(formula('bogus', 1), values), 0);
    assert.equal(calculate(formula('val', 'unknown-guide'), values), 0);
  });

  it('starts from the built-in guides and lets the shape override adjust values', () => {
    const def = presetGeometry('roundRect');
    assert.ok(def);
    const defaults = evaluateGuides(def, 200, 100);
    assert.equal(defaults.get('ss'), 100);
    assert.equal(defaults.get('hc'), 100);
    assert.equal(defaults.get('cd4'), 5400000);
    assert.equal(defaults.get('x1'), 16.667);
    const overridden = evaluateGuides(def, 200, 100, [{ name: 'adj', op: 'val', args: ['50000'] }]);
    assert.equal(overridden.get('x1'), 50);
  });
});

describe('arcs', () => {
  it('draws a quarter circle from the current point', () => {
    const arc = arcSegments(0, 50, 50, 50, 10800000, 5400000);
    assert.equal(arc.d, 'A50 50 0 0 1 50 0');
    assert.deepEqual([Math.round(arc.x) + 0, Math.round(arc.y) + 0], [50, 0]);
  });

  it('uses visual angles on ellipses', () => {
    const arc = arcSegments(200, 50, 100, 50, 0, 45 * 60000);
    // The end point lies on the ray at 45° from the center (100, 50).
    assert.ok(Math.abs(arc.x - 100 - (arc.y - 50)) < 1e-9);
    assert.ok(Math.abs(((arc.x - 100) / 100) ** 2 + ((arc.y - 50) / 50) ** 2 - 1) < 1e-9);
  });

  it('splits full turns, sweeps backwards and ignores empty arcs', () => {
    assert.equal(arcSegments(0, 50, 50, 50, 10800000, 21600000).d.split('A').length - 1, 2);
    assert.match(arcSegments(100, 50, 50, 50, 0, -5400000).d, /^A50 50 0 0 0 /);
    assert.deepEqual(arcSegments(1, 2, 0, 50, 0, 5400000), { d: '', x: 1, y: 2 });
  });
});

describe('preset geometries', () => {
  const names = [...Object.keys(PRESETS), ...Object.keys(STAR_POINTS), ...Object.keys(POLYGON_CORNERS), ...Object.keys(PRESET_ALIASES)];

  it('builds every preset without invalid numbers, and its text box inside reasonable bounds', () => {
    for (const name of names) {
      const def = presetGeometry(name);
      assert.ok(def, name);
      const built = buildGeometry(def, 200 * PX, 100 * PX);
      assert.ok(built.paths.length > 0, name);
      for (const path of built.paths) {
        assert.doesNotMatch(path.d, /NaN|Infinity|undefined/, name);
        assert.match(path.d, /^M/, name);
      }
      const { x, y, w, h } = built.textRect;
      assert.ok([x, y, w, h].every(Number.isFinite), name);
      assert.ok(x >= -1 && y >= -1 && x + w <= 201 && y + h <= 101, `${name} text box ${JSON.stringify(built.textRect)}`);
    }
  });

  it('draws rectangles, rounded rectangles and arrows at their size', () => {
    const rect = buildGeometry(presetGeometry('rect') ?? assert.fail(), 100 * PX, 50 * PX);
    assert.equal(rect.paths[0]?.d, 'M0 0 L100 0 L100 50 L0 50 Z');
    assert.deepEqual(rect.textRect, { x: 0, y: 0, w: 100, h: 50 });
    const round = buildGeometry(presetGeometry('roundRect') ?? assert.fail(), 100 * PX, 50 * PX);
    assert.match(round.paths[0]?.d ?? '', /^M0 8\.33 A8\.33 8\.33 0 0 1 8\.33 0 L91\.67 0 /);
    const arrow = buildGeometry(presetGeometry('rightArrow') ?? assert.fail(), 100 * PX, 50 * PX);
    assert.equal(arrow.paths[0]?.d, 'M0 12.5 L75 12.5 L75 0 L100 25 L75 50 L75 37.5 L0 37.5 Z');
    const line = buildGeometry(presetGeometry('line') ?? assert.fail(), 100 * PX, 0);
    assert.equal(line.paths[0]?.fill, 'none');
    assert.equal(line.paths[0]?.d, 'M0 0 L100 0');
  });

  it('points callout wedges at their adjusted tip', () => {
    const callout = buildGeometry(presetGeometry('wedgeRectCallout') ?? assert.fail(), 120 * PX, 60 * PX);
    // Default tip: 20.833% of the width left of center, 62.5% of the height below it → under the bottom edge.
    assert.match(callout.paths[0]?.d ?? '', / L50 60 L35 67\.5 L20 60 /);
    const right = buildGeometry(presetGeometry('wedgeRectCallout') ?? assert.fail(), 120 * PX, 60 * PX, [{ name: 'adj1', op: 'val', args: ['80000'] }, { name: 'adj2', op: 'val', args: ['0'] }]);
    assert.match(right.paths[0]?.d ?? '', / L156 30 /);
  });

  it('knows aliases and leaves unknown presets to the caller', () => {
    assert.equal(presetGeometry('flowChartProcess'), presetGeometry('rect'));
    assert.equal(presetGeometry('noSuchShape'), undefined);
    const star = buildGeometry(presetGeometry('star8') ?? assert.fail(), 100 * PX, 100 * PX);
    assert.equal(star.paths[0]?.d.split('L').length, 16);
  });
});

describe('custom geometry', () => {
  before(() => installDom());
  afterEach(() => {
    PPTX_LIMITS.maxPathCommands = 20_000;
    PPTX_LIMITS.maxGuides = 1_000;
  });

  const house = `<a:custGeom ${A_NS}><a:avLst/><a:gdLst><a:gd name="mid" fmla="*/ w 1 2"/></a:gdLst><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>` +
    '<a:path w="200" h="150"><a:moveTo><a:pt x="0" y="60"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="200" y="60"/></a:lnTo>' +
    '<a:lnTo><a:pt x="200" y="150"/></a:lnTo><a:lnTo><a:pt x="0" y="150"/></a:lnTo><a:close/></a:path>' +
    '<a:path fill="none" stroke="0"><a:moveTo><a:pt x="mid" y="0"/></a:moveTo><a:cubicBezTo><a:pt x="0" y="0"/><a:pt x="0" y="0"/><a:pt x="0" y="0"/></a:cubicBezTo></a:path></a:pathLst></a:custGeom>';

  it('scales path coordinates from the path space to the shape and evaluates its guides', () => {
    const built = buildGeometry(customGeometry(xml(house)), 100 * PX, 75 * PX);
    assert.equal(built.paths[0]?.d, 'M0 30 L50 0 L100 30 L100 75 L0 75 Z');
    assert.equal(built.paths[1]?.d, 'M50 0 C0 0 0 0 0 0');
    assert.equal(built.paths[1]?.fill, 'none');
    assert.equal(built.paths[1]?.stroke, false);
  });

  it('caps the number of path commands and guides a hostile geometry can ask for', () => {
    PPTX_LIMITS.maxPathCommands = 10;
    PPTX_LIMITS.maxGuides = 5;
    const commands = '<a:lnTo><a:pt x="1" y="1"/></a:lnTo>'.repeat(50);
    const guides = Array.from({ length: 50 }, (_, i) => `<a:gd name="g${i}" fmla="val ${i}"/>`).join('');
    const def = customGeometry(xml(`<a:custGeom ${A_NS}><a:gdLst>${guides}</a:gdLst><a:pathLst><a:path w="1" h="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${commands}</a:path></a:pathLst></a:custGeom>`));
    assert.equal(def.paths[0]?.commands.length, 9, 'the path itself counts as one command');
    assert.equal(def.guides.length, 5);
    assert.equal(formulasOf(xml(`<a:gdLst ${A_NS}>${guides}</a:gdLst>`)).length, 5);
    const empty = customGeometry(xml(`<a:custGeom ${A_NS}><a:pathLst>${'<a:path/>'.repeat(50)}</a:pathLst></a:custGeom>`));
    assert.equal(empty.paths.length, 10, 'empty paths cannot multiply without bound');
    assert.equal(buildGeometry(empty, 952500, 952500).paths[0]?.segments, 1);
  });

  it('reads a custom geometry element once, however many shapes inherit it', () => {
    const element = xml(house);
    assert.equal(customGeometry(element), customGeometry(element));
  });
});

describe('frames', () => {
  before(() => installDom());

  it('reads offsets, sizes, rotation and flips in EMUs and degrees', () => {
    const frame = readXfrm(xml(`<a:xfrm ${A_NS} rot="1800000" flipH="1"><a:off x="952500" y="-9525"/><a:ext cx="190500" cy="95250"/></a:xfrm>`));
    assert.deepEqual(frame, { x: 952500, y: -9525, w: 190500, h: 95250, rot: 30, flipH: true, flipV: false });
    assert.equal(emuToPx(952500), 100);
    assert.equal(readXfrm(xml(`<a:xfrm ${A_NS}><a:off x="0" y="0"/></a:xfrm>`)), undefined);
    assert.equal(readXfrm(xml(`<a:xfrm ${A_NS}><a:off x="0" y="0"/><a:ext cx="-5" cy="1e40"/></a:xfrm>`))?.w, 0, 'negative extents are clamped');
  });

  it('maps group members from the child space into the group box', () => {
    const group = readGroupXfrm(xml(`<a:xfrm ${A_NS}><a:off x="1000" y="2000"/><a:ext cx="400" cy="200"/><a:chOff x="100" y="100"/><a:chExt cx="200" cy="100"/></a:xfrm>`));
    assert.ok(group);
    const member: Xfrm = { x: 200, y: 150, w: 50, h: 20, rot: 10, flipH: false, flipV: false };
    assert.deepEqual(mapIntoGroup(member, group), { ...member, x: 200, y: 100, w: 100, h: 40 });
  });

  it('mirrors members of flipped groups without mirroring their text', () => {
    const member: Xfrm = { x: 10, y: 20, w: 30, h: 40, rot: 15, flipH: false, flipV: false };
    const group: Xfrm = { x: 0, y: 0, w: 100, h: 200, rot: 0, flipH: true, flipV: true };
    assert.deepEqual(mirror(member, group), { x: 60, y: 140, w: 30, h: 40, rot: 15, flipH: true, flipV: true });
    assert.deepEqual(mirror(member, { ...group, flipV: false }), { ...member, x: 60, rot: -15, flipH: true });
  });
});
