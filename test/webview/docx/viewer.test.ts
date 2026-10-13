// The .docx viewer inside the shell, in happy-dom. happy-dom's XML parser drops namespaced attributes, so docx-preview
// renders the samples only roughly here; what they look like is checked in Chrome (test/e2e/docx.spec.ts). These
// tests cover how the viewer wires the render into the page.
import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import type { HostMessage } from '../../../src/protocol.ts';
import { docxViewer } from '../../../src/webview/docx/viewer.ts';
import { samples } from '../../fixtures/docx.ts';
import { bootShell, type FakeHost } from '../dom/env.ts';

let host: FakeHost | undefined;
afterEach(() => host?.close());

/** The shell plus the browser globals docx-preview and the viewer use beyond those env.ts provides. */
async function bootDocxShell(): Promise<FakeHost> {
  const booted = await bootShell(docxViewer());
  const source = booted.window as unknown as Record<string, unknown>;
  for (const name of ['DOMParser', 'Node', 'XMLSerializer', 'requestAnimationFrame']) {
    (globalThis as Record<string, unknown>)[name] = source[name];
  }
  return booted;
}

function docx(fileName: string, data = sample(fileName)): HostMessage {
  return { type: 'document', format: 'docx', fileName, data };
}

function sample(fileName: string): Uint8Array {
  const found = samples.find((candidate) => candidate.fileName === fileName);
  assert.ok(found, fileName);
  // A plain Uint8Array, as the host sends it (payload.ts): a Node Buffer would arrive as JSON.
  return new Uint8Array(found.build());
}

const viewerStyles = (shell: FakeHost) => [...shell.window.document.head.querySelectorAll('style')] as unknown as HTMLStyleElement[];

describe('.docx viewer in the shell', () => {
  it('renders pages into the content area and its style sheets into <head> with the nonce', async () => {
    host = await bootDocxShell();
    const shell = host;
    host.deliver(docx('sample.docx'));
    await host.until(() => shell.window.document.querySelector('#content > .dl-docx-frame > .dl-docx > .docx-wrapper > section.docx') !== null, 'pages');
    assert.equal(host.$('#error').hidden, true);
    const styles = viewerStyles(host);
    assert.ok(styles.length > 0);
    for (const style of styles) {
      assert.equal(style.getAttribute('nonce'), 'style-nonce');
      assert.match(style.textContent ?? '', /^(@font-face|@scope \(\.dl-docx\))/);
    }
    assert.equal(host.$('#content').querySelectorAll('style, script').length, 0);
  });

  it('replaces the previous render and its style sheets on reload', async () => {
    host = await bootDocxShell();
    const shell = host;
    host.deliver(docx('sample.docx'));
    await host.until(() => viewerStyles(shell).length > 0, 'first render');
    const first = viewerStyles(host);
    const firstPage = host.$('.docx-wrapper');
    host.deliver(docx('sample.docx'));
    await host.until(() => !firstPage.isConnected, 'second render');
    const second = viewerStyles(host);
    assert.equal(second.length, first.length);
    assert.ok(first.every((style) => !style.isConnected));
    assert.equal(host.window.document.querySelectorAll('.dl-docx').length, 1);
  });

  it('keeps the last good render when a reload cannot be shown', async () => {
    host = await bootDocxShell();
    const shell = host;
    host.deliver(docx('sample.docx'));
    await host.until(() => shell.window.document.querySelector('.docx-wrapper') !== null, 'first render');
    const page = host.$('.docx-wrapper');
    const styles = viewerStyles(host);

    host.deliver(docx('sample.docx', new TextEncoder().encode('not a zip file')));
    await host.until(() => !shell.$('#error').hidden, 'error banner');
    assert.equal(host.$('#error').textContent, 'Could not display sample.docx: Not a ZIP archive (end of central directory not found)');
    assert.equal(page.isConnected, true);
    assert.deepEqual(viewerStyles(host), styles);
  });

  it('shows nothing half-rendered when the first render fails', async () => {
    host = await bootDocxShell();
    const shell = host;
    host.deliver(docx('broken.docx', new TextEncoder().encode('not a zip file')));
    await host.until(() => !shell.$('#error').hidden, 'error banner');
    assert.equal(host.$('#content').childElementCount, 0);
    assert.equal(viewerStyles(host).length, 0);
    assert.equal(host.$('#fit-width').hidden, false);
    host.key('w');
    assert.equal(host.$('#zoom-value').textContent, '100%', 'nothing to fit yet');
  });
});
