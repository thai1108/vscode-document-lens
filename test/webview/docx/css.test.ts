import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import { Window } from 'happy-dom';
import {
  clampLengths,
  embeddedFontNames,
  isSafeCssValue,
  renameFonts,
  sanitizeDeclaration,
  sanitizeDeclarations,
  sanitizeStyleSheet,
} from '../../../src/webview/docx/css.ts';

describe('docx CSS values', () => {
  it('accepts values that load nothing or only blob:/data: URLs', () => {
    for (const value of [
      '#c00000',
      'var(--docx-accent1-color)',
      '"" counter(docx-num-2-0, decimal) "."',
      'url("blob:https://webview.test/0b1c")',
      'url(data:image/png;base64,iVBORw0KGgo=)',
      'image-set("blob:https://webview.test/1" 1x, url(data:image/png;base64,AA==) 2x)',
    ]) {
      assert.equal(isSafeCssValue(value), true, value);
    }
  });

  it('refuses any other URL, whichever function carries it', () => {
    for (const value of [
      'url("https://example.com/x.png")',
      "url('//example.com/x.png')",
      'url(x.png)',
      'url("blob:https://webview.test/1"), url("http://example.com/2")',
      'image-set("https://example.com/a.png" 1x)',
      '-webkit-image-set("https://example.com/a.png" 1x)',
      'IMAGE-SET(url(http://example.com/a.png) 1x)',
      'cross-fade(url(blob:a), url(http://example.com/b), 50%)',
      'src("https://example.com/c.png")',
    ]) {
      assert.equal(isSafeCssValue(value), false, value);
    }
  });

  it('clamps lengths beyond the limits, keeping their unit', () => {
    assert.equal(clampLengths('49999999.50pt'), '5000pt');
    assert.equal(clampLengths('-1.33333e+08px'), '-6666.67px');
    assert.equal(clampLengths('calc(100% + 99999pt)'), 'calc(100% + 5000pt)');
    assert.equal(clampLengths('1000in 99999%'), '69.44in 1000%');
    assert.equal(clampLengths('500em'), '100em');
    assert.equal(clampLengths('595.30pt 841.90pt'), '595.30pt 841.90pt');
  });

  it('leaves strings, URLs, identifiers and non-length units alone', () => {
    for (const value of ['"Font 99999pt"', 'url("blob:https://webview.test/99999pt")', 'h99999px', 'docx-num-99999-0', 'rotate(99999deg)', '99999']) {
      assert.equal(clampLengths(value), value);
    }
  });
});

describe('docx CSS declarations', () => {
  it('lets var() refer only to --docx-* properties, without a fallback, and drops env() and attr()', () => {
    assert.equal(sanitizeDeclaration('color', 'var(--docx-accent1-color)'), 'var(--docx-accent1-color)');
    assert.equal(sanitizeDeclaration('font-family', 'var(--docx-majorHAnsi-font), serif'), 'var(--docx-majorHAnsi-font), serif');
    for (const [name, value] of [
      ['border-image', 'url("https://example.com/literal.png") var(--s)'],
      ['border-image-source', 'image-set(var(--u) 1x)'],
      ['margin-left', 'calc(1pt * var(--k))'],
      ['color', 'var(--docx-accent1-color, red)'],
      ['width', 'env(safe-area-inset-left)'],
      ['content', 'attr(title)'],
    ]) {
      assert.equal(sanitizeDeclaration(name ?? '', value ?? ''), undefined, `${name}: ${value}`);
    }
  });

  it('keeps only --docx-* custom properties holding a colour, a blob: URL or names', () => {
    assert.equal(sanitizeDeclaration('--docx-accent1-color', '#4472c4'), '#4472c4');
    assert.equal(sanitizeDeclaration('--docx-rid5', 'url(blob:https://webview.test/1)'), 'url(blob:https://webview.test/1)');
    assert.equal(sanitizeDeclaration('--docx-minorHAnsi-font', 'Yu Mincho'), 'Yu Mincho');
    for (const [name, value] of [
      ['--u', '#000000'],
      ['--docx-u', '"https://example.com/beacon.png"'],
      ['--docx-k', '1000000'],
      ['--docx-k', '1000000px'],
      ['--docx-bullet', 'url(https://example.com/x.png)'],
      ['--docx-bullet', 'u\\72 l(https://example.com/x.png)'],
    ]) {
      assert.equal(sanitizeDeclaration(name ?? '', value ?? ''), undefined, `${name}: ${value}`);
    }
  });

  it('drops products and quotients inside math functions, and zoom', () => {
    assert.equal(sanitizeDeclaration('margin-top', 'calc(-36pt)'), 'calc(-36pt)');
    assert.equal(sanitizeDeclaration('line-height', 'calc(100% + 12pt)'), 'calc(100% + 12pt)');
    assert.equal(sanitizeDeclaration('width', 'calc((1px * sibling-index()) * 1e+09)'), undefined);
    assert.equal(sanitizeDeclaration('width', 'min(1px / 0.000001, 2px)'), undefined);
    assert.equal(sanitizeDeclaration('zoom', '40'), undefined);
  });

  it('renames embedded fonts in font-family lists', () => {
    const fonts = new Map([['lens embedded', 'dl-docx-font-1']]);
    assert.equal(renameFonts(`'Lens Embedded', "Yu Mincho", serif`, fonts), 'dl-docx-font-1, "Yu Mincho", serif');
    assert.equal(sanitizeDeclaration('font-family', 'LENS   embedded', fonts), 'dl-docx-font-1');
    assert.equal(sanitizeDeclaration('--docx-majorHAnsi-font', 'Lens Embedded', fonts), 'dl-docx-font-1');
  });
});

describe('docx style sheets', () => {
  let window: Window;
  before(() => {
    window = new Window({ url: 'https://webview.test/' });
    (globalThis as Record<string, unknown>)['window'] = window;
  });
  after(() => window.happyDOM.close());

  it('scopes style rules to the document and drops what could load or escape', () => {
    const css = sanitizeStyleSheet(
      [
        '.docx p { color: red; background-image: url("https://example.com/x.png"); width: 99999pt; }',
        '@media print { .docx { color: blue; } }',
        'body { background-image: url("https://example.com/y.png"); }',
      ].join('\n'),
      '.dl-docx',
    );
    assert.equal(css, '@scope (.dl-docx) {\n.docx p { color: red; width: 5000pt; }\n}');
  });

  it('renames embedded fonts and keeps only their family, source, weight and style', () => {
    const sheet = [
      '@font-face { font-family: Embedded; src: url("blob:https://webview.test/font"); font-weight: bold; size-adjust: 1000%; }',
      '@font-face { font-family: "Remote"; src: url("https://example.com/font.woff"); }',
      '@font-face { font-family: "Missing"; src: url(null); }',
    ].join('\n');
    const fonts = embeddedFontNames([sheet]);
    assert.deepEqual([...fonts], [['embedded', 'dl-docx-font-1']]);
    assert.equal(
      sanitizeStyleSheet(sheet, '.dl-docx', fonts),
      '@font-face { font-family: dl-docx-font-1; src: url("blob:https://webview.test/font"); font-weight: bold; }',
    );
  });

  it('returns nothing for a sheet with nothing left', () => {
    assert.equal(sanitizeStyleSheet('@media print { p { color: red; } }', '.dl-docx'), '');
  });

  it('cleans inline declarations: no zoom, no unsafe custom properties, clamped lengths', () => {
    const element = window.document.createElement('span');
    element.style.setProperty('zoom', '40');
    element.style.setProperty('--docx-bullet', 'u\\72 l(https://example.com/x.png)');
    element.style.setProperty('--docx-accent1-color', '#4472c4');
    element.style.setProperty('font-size', '49999999.5pt');
    element.style.setProperty('line-height', '416666.25');
    element.style.setProperty('column-count', '100000');
    sanitizeDeclarations(element.style as unknown as CSSStyleDeclaration);
    assert.equal(element.style.getPropertyValue('zoom'), '');
    assert.equal(element.style.getPropertyValue('--docx-bullet'), '');
    assert.equal(element.style.getPropertyValue('--docx-accent1-color'), '#4472c4');
    assert.equal(element.style.getPropertyValue('font-size'), '5000pt');
    assert.equal(element.style.getPropertyValue('line-height'), '20');
    assert.equal(element.style.getPropertyValue('column-count'), '45');
  });
});
