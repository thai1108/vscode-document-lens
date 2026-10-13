import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { contentSecurityPolicy, webviewHtml } from '../src/webviewHtml.ts';

function directives(csp: string): Map<string, string[]> {
  return new Map(csp.split('; ').map((part) => { const [name = '', ...sources] = part.split(' '); return [name, sources]; }));
}

describe('contentSecurityPolicy', () => {
  it('allows no network, inline script or eval by default', () => {
    const policy = directives(contentSecurityPolicy('https://cdn.test', 'abc'));
    assert.deepEqual(policy.get('default-src'), ["'none'"]);
    assert.deepEqual(policy.get('script-src'), ["'nonce-abc'"]);
    assert.deepEqual(policy.get('style-src'), ['https://cdn.test', "'nonce-abc'"]);
    assert.deepEqual(policy.get('img-src'), ['https://cdn.test', 'blob:', 'data:']);
    assert.deepEqual(policy.get('font-src'), ['https://cdn.test', 'blob:', 'data:']);
    assert.equal(policy.get('connect-src'), undefined);
    assert.equal(policy.get('worker-src'), undefined);
    assert.deepEqual(policy.get('base-uri'), ["'none'"]);
    assert.deepEqual(policy.get('form-action'), ["'none'"]);
  });

  it('adds a format’s extra sources without duplicates', () => {
    const policy = directives(contentSecurityPolicy('https://cdn.test', 'abc', { 'worker-src': ['blob:'], 'img-src': ['blob:'], 'script-src': ["'wasm-unsafe-eval'"] }));
    assert.deepEqual(policy.get('worker-src'), ['blob:']);
    assert.deepEqual(policy.get('img-src'), ['https://cdn.test', 'blob:', 'data:']);
    assert.deepEqual(policy.get('script-src'), ["'nonce-abc'", "'wasm-unsafe-eval'"]);
  });
});

describe('webviewHtml', () => {
  it('loads the viewer bundle with the nonce under the policy', () => {
    const html = webviewHtml({ cspSource: 'https://cdn.test', nonce: 'n0nce', scriptUri: 'https://cdn.test/webview-pdf.js', styleUri: 'https://cdn.test/webview-pdf.css' });
    assert.match(html, /<meta http-equiv="Content-Security-Policy" content="default-src 'none'; [^"]+">/);
    assert.match(html, /<script nonce="n0nce" src="https:\/\/cdn\.test\/webview-pdf\.js"><\/script>/);
    assert.match(html, /<link rel="stylesheet" href="https:\/\/cdn\.test\/webview-pdf\.css">/);
    assert.doesNotMatch(html, /style="/);
  });
});
