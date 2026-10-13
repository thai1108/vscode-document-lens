import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { externalUrl, MESSAGE_LIMITS, parseWebviewMessage } from '../src/protocol.ts';
import { receivedFromWebview } from './support/vscodeWebviewSerializer.ts';

describe('parseWebviewMessage', () => {
  it('accepts ready, open as text and http(s) links as the host receives them', () => {
    assert.deepEqual(parseWebviewMessage(receivedFromWebview({ type: 'ready' })), { type: 'ready' });
    assert.deepEqual(parseWebviewMessage(receivedFromWebview({ type: 'openAsText', url: 'file:///etc/passwd' })), { type: 'openAsText' });
    assert.deepEqual(parseWebviewMessage(receivedFromWebview({ type: 'openLink', url: 'https://example.com/a b?q=1#x' })), {
      type: 'openLink',
      url: 'https://example.com/a%20b?q=1#x',
    });
  });

  it('drops unknown, malformed and unsafe messages', () => {
    for (const raw of [
      null,
      'ready',
      42,
      {},
      { type: 'save' },
      { type: 'openLink' },
      { type: 'openLink', url: 7 },
      { type: 'openLink', url: 'javascript:alert(1)' },
      { type: 'openLink', url: 'command:workbench.action.reloadWindow' },
      { type: 'openLink', url: 'file:///etc/passwd' },
      { type: 'openLink', url: 'vscode://settings' },
      { type: 'openLink', url: '/relative/path' },
      { type: 'openLink', url: 'https://user:secret@example.com/' },
      { type: 'openLink', url: `https://example.com/${'a'.repeat(MESSAGE_LIMITS.urlLength)}` },
    ]) {
      assert.equal(parseWebviewMessage(raw), null, JSON.stringify(raw)?.slice(0, 80));
    }
  });
});

describe('externalUrl', () => {
  it('normalizes http(s) URLs', () => {
    assert.equal(externalUrl('HTTP://Example.COM'), 'http://example.com/');
    assert.equal(externalUrl('mailto:someone@example.com'), null);
  });
});
