import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import type { LoadedDocument, Viewer, ViewerContext } from '../../../src/webview/shell/app.ts';
import { placeholderViewer } from '../../../src/webview/shell/placeholder.ts';
import { bootShell, documentMessage, type FakeHost } from './env.ts';

let host: FakeHost | undefined;
afterEach(() => host?.close());

/** A viewer that records what the shell asks of it. */
function recordingViewer(extra: Partial<Viewer> = {}) {
  const renders: { document: LoadedDocument; reload: boolean }[] = [];
  let context: ViewerContext | undefined;
  const viewer: Viewer = {
    render(document, ctx, reload) {
      renders.push({ document, reload });
      context = ctx;
      ctx.content.textContent = new TextDecoder().decode(document.data);
    },
    ...extra,
  };
  return { viewer, renders, context: () => context };
}

describe('viewer shell', () => {
  it('asks for the document, then renders it and every reload as plain bytes', async () => {
    const { viewer, renders } = recordingViewer();
    host = await bootShell(viewer);
    assert.deepEqual(host.posted, [{ type: 'ready' }]);

    host.deliver(documentMessage('report.pdf', 'first'));
    await host.until(() => renders.length === 1, 'first render');
    assert.equal(host.$('#file-name').textContent, 'report.pdf');
    assert.equal(host.$('#content').textContent, 'first');
    assert.equal(renders[0]?.document.data.constructor, Uint8Array);
    assert.equal(renders[0]?.reload, false);

    host.deliver(documentMessage('report.pdf', 'second'));
    await host.until(() => renders.length === 2, 'reload');
    assert.equal(renders[1]?.reload, true);
    assert.equal(host.$('#content').textContent, 'second');
  });

  it('renders one document at a time', async () => {
    const order: string[] = [];
    host = await bootShell({
      async render(document) {
        const text = new TextDecoder().decode(document.data);
        order.push(`start ${text}`);
        await new Promise((resolve) => setTimeout(resolve, text === 'slow' ? 30 : 0));
        order.push(`end ${text}`);
      },
    });
    host.deliver(documentMessage('a.pdf', 'slow'));
    host.deliver(documentMessage('a.pdf', 'fast'));
    await host.until(() => order.length === 4, 'both renders');
    assert.deepEqual(order, ['start slow', 'end slow', 'start fast', 'end fast']);
  });

  it('shows host errors and render failures in the error banner until a render succeeds', async () => {
    let fail = true;
    host = await bootShell({
      render(document, context) {
        if (fail) {
          throw new Error('bad table');
        }
        context.content.textContent = document.fileName;
      },
    });
    assert.equal(host.$('#error').hidden, true);
    host.deliver({ type: 'error', fileName: 'huge.pdf', message: 'huge.pdf is too large to preview' });
    assert.equal(host.$('#error').hidden, false);
    assert.equal(host.$('#error').textContent, 'huge.pdf is too large to preview');
    assert.equal(host.$('#file-name').textContent, 'huge.pdf');

    host.deliver(documentMessage('broken.pdf', ''));
    await host.until(() => host?.$('#error').textContent?.includes('bad table') === true, 'render failure');
    assert.equal(host.$('#error').textContent, 'Could not display broken.pdf: bad table');

    fail = false;
    host.deliver(documentMessage('broken.pdf', ''));
    await host.until(() => host?.$('#error').hidden === true, 'error cleared');
  });

  it('zooms with the buttons and keys, scales the content with CSS zoom and remembers the zoom', async () => {
    host = await bootShell(recordingViewer().viewer);
    const content = host.$('#content');
    assert.equal(host.$('#zoom-value').textContent, '100%');
    host.click(host.$('#zoom-in'));
    assert.equal(host.$('#zoom-value').textContent, '110%');
    assert.equal(content.style.getPropertyValue('zoom'), '1.1');
    host.key('-');
    host.key('-');
    assert.equal(host.$('#zoom-value').textContent, '90%');
    host.key('+');
    assert.equal(host.$('#zoom-value').textContent, '100%');
    host.key('=');
    assert.equal(host.state?.zoom, 1.1);
    host.click(host.$('#zoom-value'));
    assert.equal(host.$('#zoom-value').textContent, '100%');

    // Shortcuts with modifiers belong to VS Code; typing in a field is not zooming.
    host.key('+', { metaKey: true });
    const input = host.window.document.createElement('input');
    host.$('#viewer-tools').append(input as unknown as Node);
    host.key('+', {}, input as unknown as EventTarget);
    assert.equal(host.$('#zoom-value').textContent, '100%');
  });

  it('restores the saved zoom and viewer state', async () => {
    const recorded = recordingViewer();
    host = await bootShell(recorded.viewer, { state: { zoom: 2, viewer: { page: 3 } } });
    assert.equal(host.$('#zoom-value').textContent, '200%');
    host.deliver(documentMessage('a.pdf', 'x'));
    await host.until(() => recorded.renders.length === 1, 'render');
    const context = recorded.context();
    assert.deepEqual(context?.viewerState(), { page: 3 });
    context?.saveViewerState({ page: 4 });
    assert.deepEqual(host.state, { zoom: 2, viewer: { page: 4 } });
  });

  it('lets a viewer apply zoom itself and offer fit width', async () => {
    const applied: number[] = [];
    host = await bootShell(recordingViewer({ applyZoom: (zoom) => applied.push(zoom), fitWidthZoom: () => 0.87 }).viewer);
    assert.equal(host.$('#fit-width').hidden, false);
    host.click(host.$('#fit-width'));
    assert.equal(host.$('#zoom-value').textContent, '87%');
    host.key('+');
    assert.equal(host.$('#zoom-value').textContent, '90%');
    host.key('0');
    host.key('w');
    assert.deepEqual(applied, [0.87, 0.9, 1, 0.87]);
    assert.equal(host.$('#content').style.getPropertyValue('zoom'), '');
  });

  it('hides fit width when the viewer cannot fit', async () => {
    host = await bootShell(recordingViewer().viewer);
    assert.equal(host.$('#fit-width').hidden, true);
    host.key('w');
    assert.equal(host.$('#zoom-value').textContent, '100%');
  });

  it('opens http(s) links through the host and never navigates the webview', async () => {
    host = await bootShell({
      render(_document, context) {
        context.content.innerHTML =
          '<a id="web" href="https://example.com/x"><span id="inner">site</span></a><a id="anchor" href="#part-2">part</a>' +
          '<a id="script" href="javascript:alert(1)">x</a><a id="relative" href="other.docx">y</a>';
      },
    });
    host.deliver(documentMessage('a.docx', ''));
    await host.until(() => host?.window.document.getElementById('web') !== null, 'links');

    const web = host.click(host.$('#inner'));
    assert.equal(web.defaultPrevented, true);
    assert.deepEqual(host.posted.at(-1), { type: 'openLink', url: 'https://example.com/x' });

    const count = host.posted.length;
    assert.equal(host.click(host.$('#anchor')).defaultPrevented, false, 'in-document anchors scroll as usual');
    assert.equal(host.click(host.$('#script')).defaultPrevented, true);
    assert.equal(host.click(host.$('#relative')).defaultPrevented, true);
    assert.equal(host.posted.length, count, 'nothing but http(s) reaches the host');
  });

  it('shows and clears notices, including the simplified preview banner', async () => {
    const recorded = recordingViewer();
    host = await bootShell(recorded.viewer);
    host.deliver(documentMessage('old.doc', 'x', 'doc'));
    await host.until(() => recorded.renders.length === 1, 'render');
    const context = recorded.context();
    context?.showSimplifiedNotice('only the text is shown.');
    assert.equal(host.$('#notice').hidden, false);
    assert.equal(host.$('#notice').textContent, 'Simplified preview: only the text is shown.');
    context?.showNotice(null);
    assert.equal(host.$('#notice').hidden, true);
  });

  it('exposes the script nonce for viewer <style> elements', async () => {
    const recorded = recordingViewer();
    host = await bootShell(recorded.viewer);
    host.deliver(documentMessage('a.pdf', 'x'));
    await host.until(() => recorded.renders.length === 1, 'render');
    assert.equal(recorded.context()?.styleNonce, 'style-nonce');
  });

  it('placeholder viewer shows the file name, kind and size', async () => {
    host = await bootShell(placeholderViewer('.pdf document'));
    host.deliver(documentMessage('scan.pdf', 'x'.repeat(2048)));
    await host.until(() => host?.window.document.getElementById('placeholder') !== null, 'placeholder');
    assert.equal(host.$('#placeholder h1').textContent, 'scan.pdf');
    assert.equal(host.$('#placeholder-size').textContent, '.pdf document · 2.0 KB');
  });
});
