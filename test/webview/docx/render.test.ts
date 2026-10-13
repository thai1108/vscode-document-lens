import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import { Window } from 'happy-dom';
import { checkModel, elementFactory } from '../../../src/webview/docx/render.ts';
import { DOCX_LIMITS } from '../../../src/webview/docx/limits.ts';

interface Model {
  type: string;
  break?: string;
  children?: Model[];
}

const text = (): Model => ({ type: 'text' });
const paragraph = (...children: Model[]): Model => ({ type: 'paragraph', children: [{ type: 'run', children }] });
const pageBreak = (): Model => paragraph({ type: 'break', break: 'page' });
const parsed = (body: Model[], parts: unknown[] = []) => ({ documentPart: { body: { type: 'document', children: body } }, parts });

describe('docx model limits', () => {
  it('accepts an ordinary document', () => {
    checkModel(parsed([paragraph(text()), pageBreak(), paragraph(text())], [{ rootElement: { type: 'header', children: [paragraph(text())] } }]));
  });

  it('refuses a document without a main part', () => {
    assert.throws(() => checkModel({ parts: [] }), /the document is damaged \(its main part is missing\)/);
  });

  it('refuses nesting deeper than the limit, without recursing itself', () => {
    const root: Model = { type: 'table', children: [] };
    let node = root;
    for (let depth = 0; depth < 100_000; depth++) {
      const child: Model = { type: 'table', children: [] };
      node.children?.push(child);
      node = child;
    }
    assert.throws(() => checkModel(parsed([root])), new RegExp(`nested too deeply to preview \\(more than ${DOCX_LIMITS.maxDepth} levels\\)`));
  });

  it('counts a header and a footer on every page', () => {
    const header = { rootElement: { type: 'header', children: Array.from({ length: 1000 }, () => paragraph(text())) } };
    const footer = { rootElement: { type: 'footer', children: Array.from({ length: 1000 }, () => paragraph(text())) } };
    const pages = (count: number) => Array.from({ length: count - 1 }, pageBreak);
    // Each header/footer paragraph is 3 nodes and the root 1: 3001 per page and part.
    checkModel(parsed(pages(300), [header, footer]));
    assert.throws(() => checkModel(parsed(pages(400), [header, footer])), /the document is too large to preview \(about 2,40\d,\d{3} elements; limit 2,000,000\)/);
  });
});

describe('docx element factory', () => {
  let window: Window;
  before(() => {
    window = new Window({ url: 'https://webview.test/' });
  });
  after(() => window.happyDOM.close());

  it('builds elements in the given document, styling them through the CSSOM', () => {
    const owner = window.document.implementation.createHTMLDocument() as unknown as Document;
    const h = elementFactory(owner);
    const link = h({
      tagName: 'a',
      className: 'docx_hyperlink',
      style: { textDecoration: 'underline', 'text-align': 'center', $lang: 'ja-JP' },
      href: '#details',
      children: ['Jump', h({ tagName: 'span', children: [' there'] }), null as unknown as string],
    }) as HTMLAnchorElement;
    assert.equal(link.ownerDocument, owner);
    assert.equal(link.getAttribute('class'), 'docx_hyperlink');
    assert.equal(link.getAttribute('href'), '#details');
    assert.equal(link.style.getPropertyValue('text-decoration'), 'underline');
    assert.equal(link.style.getPropertyValue('text-align'), 'center');
    assert.equal(link.textContent, 'Jump there');

    const shape = h({ ns: 'http://www.w3.org/2000/svg', tagName: 'svg', style: 'position: absolute; width: 20pt' }) as SVGSVGElement;
    assert.equal(shape.namespaceURI, 'http://www.w3.org/2000/svg');
    assert.equal(shape.style.getPropertyValue('width'), '20pt');

    const fragment = h({ tagName: '#fragment', children: [h({ tagName: '#comment', children: ['comment by an author'] }), 'text'] });
    assert.equal(fragment.nodeType, 11);
    assert.equal(fragment.childNodes[0]?.nodeType, 8);
    assert.equal(fragment.childNodes[0]?.textContent, '');
    assert.equal(fragment.textContent, 'text');
  });
});
