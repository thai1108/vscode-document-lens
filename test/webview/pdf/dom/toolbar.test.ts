// The PDF toolbar inside the real shell, in happy-dom. pdf.js itself needs a canvas and a worker, so rendering, find
// results and links are covered by the Playwright suite (test/e2e/pdf.spec.ts).
import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import type { ViewerContext } from '../../../../src/webview/shell/app.ts';
import { matchesLabel, PdfToolbar, type ToolbarActions } from '../../../../src/webview/pdf/toolbar.ts';
import { bootShell, documentMessage, type FakeHost } from '../../dom/env.ts';

let host: FakeHost | undefined;
let toolbar: PdfToolbar | undefined;
afterEach(async () => {
  toolbar?.dispose();
  await host?.close();
  host = toolbar = undefined;
});

/** Boots the shell with a viewer that only mounts the toolbar, recording the actions it asks for. */
async function bootToolbar() {
  const calls: string[] = [];
  const actions: ToolbarActions = {
    goToPage: (page) => calls.push(`page ${page}`),
    find: (query, { again, previous }) => calls.push(`find ${JSON.stringify(query)}${again ? ' again' : ''}${previous ? ' previous' : ''}`),
    closeFind: () => calls.push('close find'),
    toggleOutline: () => calls.push('toggle outline'),
    focusDocument: () => calls.push('focus'),
  };
  let context: ViewerContext | undefined;
  host = await bootShell({
    render(_document, ctx) {
      context = ctx;
    },
  });
  host.deliver(documentMessage('a.pdf', '%PDF-'));
  await host.until(() => context !== undefined, 'render');
  toolbar = new PdfToolbar((context as ViewerContext).tools, actions);
  const shell = host;
  const input = (selector: string) => shell.$<HTMLInputElement>(selector);
  const type = (selector: string, value: string, init: { isComposing?: boolean } = {}) => {
    input(selector).value = value;
    input(selector).dispatchEvent(new shell.window.InputEvent('input', { bubbles: true, ...init }) as unknown as Event);
  };
  return { host: shell, toolbar, calls, input, type };
}

describe('PDF toolbar: pages', () => {
  it('shows the page and page count in the shell toolbar and steps with the buttons', async () => {
    const { host, toolbar, calls, input } = await bootToolbar();
    assert.ok(host.$('#viewer-tools').contains(input('#pdf-page-number')));
    assert.equal(input('#pdf-page-number').disabled, true, 'nothing to navigate before the page count is known');

    toolbar.setPageCount(12);
    assert.equal(input('#pdf-page-number').value, '1');
    assert.equal(host.$('#pdf-page-count').textContent, 'of 12');
    assert.equal(host.$<HTMLButtonElement>('#pdf-previous-page').disabled, true);
    host.click(host.$('#pdf-next-page'));
    assert.deepEqual(calls, ['page 2']);

    toolbar.setPage(12);
    assert.equal(input('#pdf-page-number').value, '12');
    assert.equal(host.$<HTMLButtonElement>('#pdf-previous-page').disabled, false);
    assert.equal(host.$<HTMLButtonElement>('#pdf-next-page').disabled, true);
    host.click(host.$('#pdf-previous-page'));
    assert.deepEqual(calls, ['page 2', 'page 11']);
  });

  it('goes to a typed page number, clamped, and puts the current page back otherwise', async () => {
    const { host, toolbar, calls, input } = await bootToolbar();
    toolbar.setPageCount(12);
    const field = input('#pdf-page-number');
    for (const [text, expected] of [
      ['7', 'page 7'],
      ['40', 'page 12'],
    ]) {
      field.focus();
      field.value = text as string;
      host.key('Enter', {}, field as unknown as EventTarget);
      assert.equal(calls.at(-1), expected);
    }
    calls.length = 0;
    field.focus();
    field.value = 'seven';
    host.key('Enter', {}, field as unknown as EventTarget);
    assert.deepEqual(calls, ['focus']);
    assert.equal(field.value, '1');

    // Typing in the field is not a zoom shortcut, and Escape gives the document the focus back.
    field.focus();
    host.key('+', {}, field as unknown as EventTarget);
    assert.equal(host.$('#zoom-value').textContent, '100%');
    field.value = '3';
    host.key('Escape', {}, field as unknown as EventTarget);
    assert.equal(field.value, '1');
    assert.equal(calls.at(-1), 'focus');
  });

  it('shows the outline button only when there is an outline, pressed while it is open', async () => {
    const { host, toolbar, calls } = await bootToolbar();
    const button = host.$('#pdf-outline-toggle');
    assert.equal(button.hidden, true);
    toolbar.setOutline(true, false);
    assert.equal(button.hidden, false);
    assert.equal(button.getAttribute('aria-pressed'), 'false');
    host.click(button);
    assert.deepEqual(calls, ['toggle outline']);
    toolbar.setOutline(true, true);
    assert.equal(button.getAttribute('aria-pressed'), 'true');
    toolbar.setOutline(false, true);
    assert.equal(button.hidden, true);
    assert.equal(button.getAttribute('aria-pressed'), 'false');
  });
});

describe('PDF toolbar: find', () => {
  it('opens with Ctrl+F or Cmd+F, searches as you type, steps with Enter / Shift+Enter and closes with Escape', async () => {
    const { host, calls, input, type } = await bootToolbar();
    const bar = host.$('#pdf-find-bar');
    assert.equal(bar.hidden, true);
    host.key('f', { ctrlKey: true });
    assert.equal(bar.hidden, false);
    assert.equal(host.$('#pdf-find-toggle').hidden, true);
    assert.equal(host.window.document.activeElement, input('#pdf-find-input') as unknown as Element);

    type('#pdf-find-input', 'needle');
    host.key('Enter', {}, input('#pdf-find-input') as unknown as EventTarget);
    host.key('Enter', { shiftKey: true }, input('#pdf-find-input') as unknown as EventTarget);
    host.click(host.$('#pdf-find-next'));
    host.click(host.$('#pdf-find-previous'));
    host.key('g', { metaKey: true });
    host.key('F3', { shiftKey: true });
    assert.deepEqual(calls, [
      'find "needle"',
      'find "needle" again',
      'find "needle" again previous',
      'find "needle" again',
      'find "needle" again previous',
      'find "needle" again',
      'find "needle" again previous',
    ]);

    calls.length = 0;
    host.key('Escape');
    assert.equal(bar.hidden, true);
    assert.equal(host.$('#pdf-find-toggle').hidden, false);
    assert.deepEqual(calls, ['close find', 'focus']);

    // Reopening (here with Cmd+F) keeps the query and shows its matches again.
    calls.length = 0;
    host.key('f', { metaKey: true });
    assert.equal(input('#pdf-find-input').value, 'needle');
    assert.deepEqual(calls, ['find "needle"']);
  });

  it('waits for an IME composition to be committed before searching or stepping', async () => {
    const { host, calls, input, type } = await bootToolbar();
    host.click(host.$('#pdf-find-toggle'));
    type('#pdf-find-input', 'にほ', { isComposing: true });
    assert.deepEqual(calls, []);
    // A keydown that belongs to the composition (e.g. Enter committing it) is not a find shortcut.
    input('#pdf-find-input').dispatchEvent(
      new host.window.KeyboardEvent('keydown', { key: 'Enter', isComposing: true, bubbles: true, cancelable: true }) as unknown as Event,
    );
    input('#pdf-find-input').value = '日本';
    input('#pdf-find-input').dispatchEvent(new host.window.Event('compositionend', { bubbles: true }) as unknown as Event);
    assert.deepEqual(calls, ['find "日本"']);
  });

  it('shows the match count, "No results" when the whole document has none, and nothing while searching', async () => {
    const { host, toolbar, type } = await bootToolbar();
    host.click(host.$('#pdf-find-toggle'));
    type('#pdf-find-input', 'needle');
    const count = host.$('#pdf-find-count');
    toolbar.setMatches({ current: 0, total: 0, notFound: false });
    assert.equal(count.textContent, '');
    toolbar.setMatches({ current: 3, total: 12, notFound: false });
    assert.equal(count.textContent, '3 of 12');
    assert.equal(count.classList.contains('no-results'), false);
    toolbar.setMatches({ current: 0, total: 0, notFound: true });
    assert.equal(count.textContent, 'No results');
    assert.equal(count.classList.contains('no-results'), true);
  });

  it('ignores the step shortcuts while there is no query, and Escape while the bar is closed', async () => {
    const { host, calls } = await bootToolbar();
    host.key('g', { ctrlKey: true });
    host.key('F3');
    host.key('Escape');
    assert.deepEqual(calls, []);
    assert.equal(host.$('#pdf-find-bar').hidden, true);
  });

  it('stops listening to the keyboard once disposed', async () => {
    const { host, toolbar } = await bootToolbar();
    toolbar.dispose();
    host.key('f', { ctrlKey: true });
    assert.equal(host.$('#pdf-find-bar').hidden, true);
  });
});

describe('matchesLabel', () => {
  it('needs a query and a selected match to count', () => {
    assert.equal(matchesLabel('', { current: 1, total: 2, notFound: false }), '');
    assert.equal(matchesLabel('x', { current: 1, total: 2, notFound: false }), '1 of 2');
    assert.equal(matchesLabel('x', { current: 0, total: 2, notFound: false }), '');
    assert.equal(matchesLabel('x', { current: 0, total: 0, notFound: true }), 'No results');
  });
});
