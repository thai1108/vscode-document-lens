import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import { Window } from 'happy-dom';
import { IMAGE_PLACEHOLDER_CLASS, sanitizeTree } from '../../../src/webview/docx/sanitize.ts';

const SVG = 'http://www.w3.org/2000/svg';
const MATHML = 'http://www.w3.org/1998/Math/MathML';

describe('docx DOM sanitizer', () => {
  let window: Window;
  let owner: Document;
  before(() => {
    window = new Window({ url: 'https://webview.test/' });
    owner = window.document.implementation.createHTMLDocument() as unknown as Document;
  });
  after(() => window.happyDOM.close());

  /** Parses `html` into a container in an inert document and sanitizes it. */
  function sanitized(html: string): HTMLElement {
    const root = owner.createElement('div');
    root.innerHTML = html;
    sanitizeTree(root, { idPrefix: 'dl-docx-' });
    return root;
  }

  it('removes elements that could run code, load content or change the page', () => {
    const root = sanitized(
      '<section class="docx"><p>kept</p>' +
        '<script>alert(1)</script><iframe src="https://example.com/"></iframe><object data="x.swf"></object><embed src="x">' +
        '<base href="https://example.com/"><meta http-equiv="refresh" content="0"><link rel="stylesheet" href="x.css">' +
        '<style>body { display: none }</style><form><input></form><video src="x.mp4"></video><template><p>t</p></template></section>',
    );
    assert.equal(root.innerHTML, '<section class="docx"><p>kept</p></section>');

    const math = owner.createElementNS(MATHML, 'math');
    math.append(owner.createElementNS(MATHML, 'maction'), owner.createElementNS(MATHML, 'mrow'));
    const container = owner.createElement('div');
    container.append(math);
    sanitizeTree(container, { idPrefix: 'dl-docx-' });
    assert.deepEqual([...math.children].map((child) => child.localName), ['mrow']);
  });

  it('keeps only the attributes docx-preview sets, dropping event handlers', () => {
    const root = sanitized('<p class="docx_normal" lang="ja-JP" onclick="alert(1)" data-x="1" title="t"><span id="mark" onmouseover="x()">a</span></p>');
    assert.equal(root.innerHTML, '<p class="docx_normal" lang="ja-JP" title="t"><span id="dl-docx-mark">a</span></p>');

    // docx-preview appends rows straight to the table (no tbody).
    const table = owner.createElement('table');
    const row = table.appendChild(owner.createElement('tr'));
    const cell = row.appendChild(owner.createElement('td'));
    for (const [name, value] of [['colspan', '2'], ['rowspan', '3'], ['background', 'https://example.com/b.png'], ['onload', 'x()']]) {
      cell.setAttribute(name ?? '', value ?? '');
    }
    const container = owner.createElement('div');
    container.append(table);
    sanitizeTree(container, { idPrefix: 'dl-docx-' });
    assert.equal(container.innerHTML, '<table><tr><td colspan="2" rowspan="3"></td></tr></table>');
  });

  it('keeps http(s) and in-document links only', () => {
    const root = sanitized(
      [
        'https://example.com/page',
        'http://example.com/',
        '#bookmark',
        'javascript:alert(1)',
        'JaVaScRiPt:alert(1)',
        'vbscript:msgbox(1)',
        'data:text/html,<script>alert(1)</script>',
        'file:///etc/passwd',
        'https://user:secret@example.com/',
        '#',
        '',
      ]
        .map((href) => `<a href="${href.replaceAll('"', '&quot;')}">link</a>`)
        .join(''),
    );
    assert.deepEqual(
      [...root.querySelectorAll('a')].map((link) => link.getAttribute('href')),
      ['https://example.com/page', 'http://example.com/', '#dl-docx-bookmark', null, null, null, null, null, null, null, null],
    );
    assert.equal(root.querySelector('a')?.getAttribute('title'), 'https://example.com/page');
  });

  it('replaces pictures that are not stored in the file with a labelled placeholder of the same size', () => {
    const root = sanitized(
      '<img src="blob:https://webview.test/1"><img src="data:image/png;base64,iVBORw0KGgo=">' +
        '<img src="https://example.com/tracker.png" style="width: 72pt; height: 36pt"><img src="null"><img>' +
        '<img src="data:text/html,<script>alert(1)</script>">',
    );
    assert.deepEqual(
      [...root.querySelectorAll('img')].map((image) => image.getAttribute('src')),
      ['blob:https://webview.test/1', 'data:image/png;base64,iVBORw0KGgo='],
    );
    const placeholders = [...root.querySelectorAll(`.${IMAGE_PLACEHOLDER_CLASS}`)] as HTMLElement[];
    assert.equal(placeholders.length, 4);
    assert.equal(placeholders[0]?.textContent, 'Picture not shown');
    assert.equal(placeholders[0]?.style.getPropertyValue('width'), '72pt');
    assert.equal(placeholders[0]?.style.getPropertyValue('height'), '36pt');
  });

  it('cleans SVG shapes: no external paint or picture, clamped coordinates, no SMIL or scripts', () => {
    const root = owner.createElement('div');
    const svg = owner.createElementNS(SVG, 'svg');
    const shape = owner.createElementNS(SVG, 'rect');
    shape.setAttribute('fill', 'url(https://example.com/fill.svg#p)');
    shape.setAttribute('stroke', '#ff0000');
    shape.setAttribute('width', '99999999');
    shape.setAttribute('height', '100%');
    shape.setAttribute('onclick', 'alert(1)');
    const picture = owner.createElementNS(SVG, 'image');
    picture.setAttribute('href', 'https://example.com/picture.png');
    const set = owner.createElementNS(SVG, 'set');
    set.setAttribute('attributeName', 'href');
    svg.append(shape, picture, set, owner.createElementNS(SVG, 'script'), owner.createElementNS(SVG, 'foreignObject'));
    root.append(svg);
    sanitizeTree(root, { idPrefix: 'dl-docx-' });
    assert.equal(shape.getAttribute('fill'), null);
    assert.equal(shape.getAttribute('stroke'), '#ff0000');
    assert.equal(shape.getAttribute('width'), '6666.67');
    assert.equal(shape.getAttribute('height'), '100%');
    assert.equal(shape.getAttribute('onclick'), null);
    assert.equal(picture.getAttribute('href'), null);
    assert.deepEqual([...svg.children].map((child) => child.localName), ['rect', 'image', 'foreignObject']);
  });

  it('applies the CSS rules to inline styles', () => {
    const root = sanitized('<span style="font-size: 99999pt; background-image: url(https://example.com/x.png); color: red">x</span>');
    const span = root.querySelector('span') as HTMLElement;
    assert.equal(span.style.getPropertyValue('font-size'), '5000pt');
    assert.equal(span.style.getPropertyValue('background-image'), '');
    assert.equal(span.style.getPropertyValue('color'), 'red');
  });
});
