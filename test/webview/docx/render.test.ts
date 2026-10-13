import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import { Window } from 'happy-dom';
import { elementFactory } from '../../../src/webview/docx/render.ts';
import { DOCX_LIMITS } from '../../../src/webview/docx/limits.ts';
import { checkModel } from '../../../src/webview/docx/model.ts';

interface Model {
  type: string;
  id?: string;
  break?: string;
  styleName?: string;
  pageBreakBefore?: boolean;
  children?: Model[];
}

const text = (): Model => ({ type: 'text' });
const paragraph = (...children: Model[]): Model => ({ type: 'paragraph', children: [{ type: 'run', children }] });
const pageBreak = (): Model => paragraph({ type: 'break', break: 'page' });
const paragraphs = (count: number) => Array.from({ length: count }, () => paragraph(text()));
const parsed = (body: Model[], more: Record<string, unknown> = {}) => ({ documentPart: { body: { type: 'document', children: body } }, parts: [], ...more });
// Each paragraph of paragraphs() is 3 nodes and a header's root 1: 1000 of them make 3001 per page.
const header = { rootElement: { type: 'header', children: paragraphs(1000) } };
const footer = { rootElement: { type: 'footer', children: paragraphs(1000) } };
const tooLarge = /the document is too large to preview \(about [\d,]+ elements; limit 2,000,000\)/;

describe('docx model limits', () => {
  it('accepts an ordinary document', () => {
    checkModel(parsed([paragraph(text()), pageBreak(), paragraph(text())], { parts: [{ rootElement: { type: 'header', children: [paragraph(text())] } }] }));
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

  it('counts a header and a footer on every page started by a page break', () => {
    const pages = (count: number) => Array.from({ length: count - 1 }, pageBreak);
    checkModel(parsed(pages(300), { parts: [header, footer] }));
    assert.throws(() => checkModel(parsed(pages(400), { parts: [header, footer] })), /about 2,40\d,\d{3} elements/);
  });

  it('counts pages started by paragraphs that break the page before them, directly or through their style', () => {
    const direct = Array.from({ length: 400 }, (): Model => ({ type: 'paragraph', pageBreakBefore: true, children: [] }));
    assert.throws(() => checkModel(parsed(direct, { parts: [header, footer] })), tooLarge);
    // Heading2 is based on Heading1, which breaks the page; Loop is based on itself.
    const styles = [
      { id: 'Heading1', paragraphProps: { pageBreakBefore: true } },
      { id: 'Heading2', basedOn: 'Heading1', paragraphProps: {} },
      { id: 'Loop', basedOn: 'Loop', paragraphProps: {} },
    ];
    const styled = (styleName: string) => Array.from({ length: 400 }, (): Model => ({ type: 'paragraph', styleName, children: [] }));
    assert.throws(() => checkModel(parsed(styled('Heading2'), { parts: [header, footer], stylesPart: { styles } })), tooLarge);
    checkModel(parsed(styled('Loop'), { parts: [header, footer], stylesPart: { styles } }));
  });

  it('counts a note once for every reference to it', () => {
    // A 300-paragraph note is 901 nodes.
    const note = { id: '1', children: paragraphs(300) };
    const references = (type: string, count: number) => [paragraph(...Array.from({ length: count }, (): Model => ({ type, id: '1' })))];
    checkModel(parsed(references('footnoteReference', 2000), { footnotesPart: { notes: [note] } }));
    assert.throws(() => checkModel(parsed(references('footnoteReference', 3000), { footnotesPart: { notes: [note] } })), tooLarge);
    assert.throws(() => checkModel(parsed(references('endnoteReference', 3000), { endnotesPart: { notes: [note] } })), tooLarge);
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
