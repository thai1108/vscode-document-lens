import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import type { HostMessage } from '../../../src/protocol.ts';
import { LAYOUT_LIMITS } from '../../../src/webview/doc/layout.ts';
import { docViewer } from '../../../src/webview/doc/viewer.ts';
import { DOC_LIMITS } from '../../../src/webview/doc/wordDocument.ts';
import { SAMPLE_DOC } from '../../fixtures/doc.ts';
import { writeWordDocument, type WordDocumentInput } from '../../support/wordWriter.ts';
import { bootShell, type FakeHost } from '../dom/env.ts';

let host: FakeHost | undefined;
afterEach(() => host?.close());

function docMessage(input: WordDocumentInput, fileName = 'sample.doc'): HostMessage {
  return { type: 'document', format: 'doc', fileName, data: new Uint8Array(writeWordDocument(input)) };
}

const texts = (selector: string) => [...(host?.window.document.querySelectorAll(selector) ?? [])].map((node) => node.textContent);

async function rendered(): Promise<void> {
  await host?.until(() => host?.window.document.getElementById('doc-paper') !== null || host?.$('#error').hidden === false, 'render');
}

describe('.doc viewer', () => {
  it('renders the paragraphs, table, link, page break and footnotes of the sample with the simplified notice', async () => {
    host = await bootShell(docViewer());
    host.deliver(docMessage(SAMPLE_DOC));
    await rendered();

    assert.equal(host.$('#error').hidden, true);
    assert.equal(host.$('#notice').hidden, false);
    assert.equal(host.$('#notice').textContent, 'Simplified preview: only the text of this document is shown; formatting, pictures and page layout are left out.');
    assert.deepEqual(texts('#doc-paper > .doc-p').slice(0, 5), [
      'Document Lens sample',
      'This synthetic Word 97–2003 document has paragraphs, a table, a field and a footnote1.',
      'Prices: 5 € – “quoted” text.',
      'Visit the example site for more.',
      'Tracked changes: only the kept text shows.',
    ]);
    assert.ok(texts('#doc-paper > .doc-p').includes('日本語のテキストも表示されます。'));
    assert.deepEqual(texts('.doc-note-ref'), ['1', '1']);

    const rows = [...host.window.document.querySelectorAll('.doc-table tr')].map((row) => [...row.querySelectorAll('td')].map((cell) => [...cell.querySelectorAll('p')].map((p) => p.textContent)));
    assert.deepEqual(rows, [
      [['Item'], ['Qty'], ['Remarks']],
      [['Apples'], ['3'], ['Fresh']],
      [['Pears'], ['5'], ['Two', 'paragraphs']],
    ]);

    const link = host.$<HTMLAnchorElement>('.doc-link');
    assert.equal(link.getAttribute('href'), 'https://example.com/');
    assert.equal(link.textContent, 'the example site');
    host.click(link);
    assert.deepEqual(host.posted.at(-1), { type: 'openLink', url: 'https://example.com/' });

    const pageBreak = host.$('.doc-page-break');
    assert.equal(pageBreak.getAttribute('role'), 'separator');
    assert.equal(pageBreak.textContent, '');
    assert.deepEqual(texts('.doc-story-label'), ['Footnotes']);
    assert.deepEqual(texts('[data-story="footnotes"] .doc-p'), ['1A footnote at the end of the text.']);

    // The CSP refuses inline styles, so the viewer must not generate any.
    assert.equal(host.$('#content').querySelectorAll('[style]').length, 0);
    assert.equal(host.$('#fit-width').hidden, false);
  });

  it('renders the new file on reload', async () => {
    host = await bootShell(docViewer());
    host.deliver(docMessage({ body: ['First version\r'] }));
    await rendered();
    assert.deepEqual(texts('.doc-p'), ['First version']);
    host.deliver(docMessage({ body: ['Second version\r', 'with two paragraphs\r'] }));
    await host.until(() => texts('.doc-p').length === 2, 'reload');
    assert.deepEqual(texts('.doc-p'), ['Second version', 'with two paragraphs']);
  });

  it('shows errors without stale content or notice, and recovers on the next good file', async () => {
    host = await bootShell(docViewer());
    host.deliver(docMessage({ body: ['Readable\r'] }));
    await rendered();
    host.deliver(docMessage({ body: ['Secret\r'], encrypted: true }, 'locked.doc'));
    await host.until(() => host?.$('#error').hidden === false, 'error banner');
    assert.equal(host.$('#error').textContent, 'Could not display locked.doc: This .doc file is password-protected; Document Lens cannot open encrypted files');
    assert.equal(host.$('#content').childElementCount, 0);
    assert.equal(host.$('#notice').hidden, true);

    host.deliver(docMessage({ body: ['Readable again\r'] }));
    await host.until(() => host?.$('#error').hidden === true, 'error cleared');
    assert.deepEqual(texts('.doc-p'), ['Readable again']);
  });

  it('reports a corrupt file in the error banner', async () => {
    host = await bootShell(docViewer());
    const data = new Uint8Array(writeWordDocument({ body: ['x\r'] }));
    data.fill(0, 512);
    host.deliver({ type: 'document', format: 'doc', fileName: 'broken.doc', data });
    await host.until(() => host?.$('#error').hidden === false, 'error banner');
    assert.match(host.$('#error').textContent ?? '', /^Could not display broken\.doc: Corrupt compound file: /);
  });

  it('says when a long document is cut', async () => {
    host = await bootShell(docViewer({ ...DOC_LIMITS, ...LAYOUT_LIMITS, maxCharacters: 25 }));
    host.deliver(docMessage(SAMPLE_DOC));
    await rendered();
    assert.match(host.$('#notice').textContent ?? '', /It is very long, so only its beginning is shown\.$/);
    assert.deepEqual(texts('.doc-p'), ['Document Lens sample', 'This']);
  });

  it('says when a document has no text', async () => {
    host = await bootShell(docViewer());
    // Word's empty document is one paragraph mark; a piece table without pieces is accepted too.
    for (const body of [['\r', '\x01\r'], []]) {
      host.deliver(docMessage({ body }));
      await host.until(() => texts('#doc-paper > *').length === 1, 'empty document');
      assert.deepEqual(texts('#doc-paper > *'), ['This document has no text.']);
      assert.equal(host.$('#error').hidden, true);
    }
  });
});
