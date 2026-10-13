import assert from 'node:assert/strict';
import { before, describe, it } from 'node:test';
import { colorIn, DEFAULT_COLOR_MAP, hexToRgba, presetColor, readColor, toCss, toHex, transformColor, type ColorContext, type Rgba } from '../../../src/webview/pptx/color.ts';
import { A_NS, installDom, xml } from './dom.ts';

const ACCENT1 = hexToRgba('4472C4') as Rgba;
const apply = (color: Rgba, ...transforms: [string, number][]) => transforms.reduce((result, [name, value]) => transformColor(result, name, value), color);

const context: ColorContext = {
  scheme: { dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', hlink: '0563C1' },
  map: { ...DEFAULT_COLOR_MAP, bg1: 'dk1', tx1: 'lt1' },
};

describe('color transforms', () => {
  it('reproduces the theme color variants PowerPoint offers for an accent', () => {
    assert.equal(toHex(apply(ACCENT1, ['lumMod', 20000], ['lumOff', 80000])), '#dae3f3');
    assert.equal(toHex(apply(ACCENT1, ['lumMod', 40000], ['lumOff', 60000])), '#b4c7e7');
    assert.equal(toHex(apply(ACCENT1, ['lumMod', 60000], ['lumOff', 40000])), '#8faadc');
    assert.equal(toHex(apply(ACCENT1, ['lumMod', 75000])), '#2f5597');
    assert.equal(toHex(apply(ACCENT1, ['lumMod', 50000])), '#203864');
  });

  it('tints toward white and shades toward black in linear RGB, as table styles use them', () => {
    assert.equal(toHex(apply(ACCENT1, ['tint', 20000])), '#e9ebf5');
    assert.equal(toHex(apply(ACCENT1, ['tint', 40000])), '#cfd5ea');
    assert.equal(toHex(apply(ACCENT1, ['tint', 0])), '#ffffff');
    assert.equal(toHex(apply(ACCENT1, ['shade', 0])), '#000000');
    assert.equal(toHex(apply(ACCENT1, ['tint', 100000])), '#4472c4');
  });

  it('changes alpha, hue, saturation and channels', () => {
    assert.equal(apply(ACCENT1, ['alpha', 50000]).a, 0.5);
    assert.equal(apply(ACCENT1, ['alpha', 50000], ['alphaMod', 50000]).a, 0.25);
    assert.equal(apply(ACCENT1, ['alphaOff', -30000]).a, 0.7);
    assert.equal(toHex(apply(hexToRgba('FF0000') as Rgba, ['hueOff', 120 * 60000])), '#00ff00');
    assert.equal(toHex(apply(hexToRgba('FF0000') as Rgba, ['comp', 0])), '#00ffff');
    assert.equal(toHex(apply(hexToRgba('336699') as Rgba, ['inv', 0])), '#cc9966');
    assert.equal(toHex(apply(hexToRgba('FF0000') as Rgba, ['sat', 0])), '#808080');
    assert.equal(toHex(apply(hexToRgba('FFFFFF') as Rgba, ['gray', 0])), '#ffffff');
    assert.equal(toHex(apply(hexToRgba('000000') as Rgba, ['red', 100000])), '#ff0000');
    // An unknown transform leaves the color alone.
    assert.equal(toHex(apply(ACCENT1, ['sparkle', 1])), '#4472c4');
  });

  it('formats CSS colors', () => {
    assert.equal(toCss({ r: 255, g: 0, b: 0, a: 1 }), '#ff0000');
    assert.equal(toCss({ r: 255.4, g: -3, b: 300, a: 0.5 }), 'rgba(255, 0, 255, 0.5)');
  });

  it('names preset colors with their dk/lt/med abbreviations', () => {
    assert.equal(presetColor('black'), '000000');
    assert.equal(presetColor('dkBlue'), '00008b');
    assert.equal(presetColor('ltGray'), 'd3d3d3');
    assert.equal(presetColor('medPurple'), '9370db');
    assert.equal(presetColor('noSuchColor'), undefined);
    assert.equal(hexToRgba('12345'), undefined);
    assert.equal(hexToRgba('zz0000'), undefined);
  });
});

describe('color elements', () => {
  before(() => installDom());

  const color = (inner: string, ctx = context) => readColor(xml(`<a:wrap ${A_NS}>${inner}</a:wrap>`).firstElementChild ?? undefined, ctx);

  it('resolves scheme colors through the slide color map', () => {
    assert.equal(toHex(color('<a:schemeClr val="bg1"/>') as Rgba), '#000000');
    assert.equal(toHex(color('<a:schemeClr val="tx1"/>') as Rgba), '#ffffff');
    assert.equal(toHex(color('<a:schemeClr val="accent2"/>') as Rgba), '#ed7d31');
    assert.equal(color('<a:schemeClr val="accent6"/>'), undefined);
    assert.equal(toHex(color('<a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr>') as Rgba), '#203864');
  });

  it('uses the style reference color for phClr', () => {
    assert.equal(color('<a:schemeClr val="phClr"/>'), undefined);
    const placeholder = { r: 1, g: 2, b: 3, a: 1 };
    assert.deepEqual(color('<a:schemeClr val="phClr"/>', { ...context, placeholder }), placeholder);
  });

  it('reads RGB, system, preset, HSL and scRGB colors', () => {
    assert.equal(toHex(color('<a:srgbClr val="00B050"/>') as Rgba), '#00b050');
    assert.equal(toHex(color('<a:sysClr val="windowText" lastClr="112233"/>') as Rgba), '#112233');
    assert.equal(toHex(color('<a:sysClr val="window"/>') as Rgba), '#ffffff');
    assert.equal(toHex(color('<a:prstClr val="ltBlue"/>') as Rgba), '#add8e6');
    assert.equal(toHex(color('<a:hslClr hue="0" sat="100000" lum="50000"/>') as Rgba), '#ff0000');
    assert.equal(toHex(color('<a:scrgbClr r="100000" g="0" b="0"/>') as Rgba), '#ff0000');
    assert.equal(color('<a:srgbClr val="not-a-color"/>'), undefined);
  });

  it('finds the color element inside a fill', () => {
    const fill = xml(`<a:solidFill ${A_NS}><a:srgbClr val="FF0000"><a:alpha val="25000"/></a:srgbClr></a:solidFill>`);
    assert.equal(toCss(colorIn(fill, context) as Rgba), 'rgba(255, 0, 0, 0.25)');
    assert.equal(colorIn(undefined, context), undefined);
  });
});
