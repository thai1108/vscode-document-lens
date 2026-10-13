// The PDF viewer inside a real VS Code extension host: every PDF sample opens in the Document Lens editor (the
// viewer itself explains the ones it cannot display), and the bundle VS Code loads carries everything pdf.js fetches
// at runtime, since the webview may not fetch anything. Rendering is covered by test/e2e/pdf.spec.ts.
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readDocument } from '../../src/documentProvider.ts';
import { pdf } from '../../src/formats/pdf.ts';
import { viewTypeOf } from '../../src/formats/types.ts';
import { samples } from '../fixtures/pdf.ts';
import { openByDefault, workspaceUri } from '../support/vscodeEditors.ts';

// Mocha's TDD globals, provided by @vscode/test-cli inside the extension host.
declare function suite(name: string, body: () => void): void;
declare function test(name: string, body: () => Promise<void> | void): void;
declare function teardown(body: () => Thenable<unknown>): void;

function extensionDist(): string {
  const extension = vscode.extensions.getExtension('thai1108.document-lens');
  assert.ok(extension, 'extension is installed in the test host');
  return path.join(extension.extensionPath, 'dist');
}

suite('PDF viewer in VS Code', () => {
  teardown(() => vscode.commands.executeCommand('workbench.action.closeAllEditors'));

  for (const { fileName } of samples) {
    test(`opens ${fileName} in the ${viewTypeOf(pdf)} editor`, async () => {
      const tab = await openByDefault(workspaceUri(fileName));
      assert.equal(tab.viewType, viewTypeOf(pdf));
    });
  }

  test('the host hands every PDF sample to the viewer unchanged, including encrypted and broken ones', async () => {
    for (const { fileName, build } of samples) {
      const bytes = await readDocument(workspaceUri(fileName), pdf);
      assert.ok(Buffer.from(bytes).equals(build()), fileName);
    }
  });

  test('the bundle embeds the pdf.js worker, CMaps, standard fonts and wasm decoders, with their licenses', () => {
    const dist = extensionDist();
    assert.ok(existsSync(path.join(dist, 'webview-pdf.css')));
    const bundle = readFileSync(path.join(dist, 'webview-pdf.js'), 'utf8');
    for (const name of ['WorkerMessageHandler', 'UniJIS-UCS2-H.bcmap', 'Adobe-Japan1-UCS2.bcmap', 'FoxitDingbats.pfb', 'jbig2.wasm', 'openjpeg.wasm']) {
      assert.ok(bundle.includes(name), `${name} in webview-pdf.js`);
    }
    const licenses = readFileSync(path.join(dist, 'webview-pdf.licenses.txt'), 'utf8');
    for (const text of ['Apache License', 'Adobe Systems Incorporated', 'The PDFium Authors', 'Universite catholique de Louvain']) {
      assert.ok(licenses.includes(text), `${text} in webview-pdf.licenses.txt`);
    }
  });

  test('the PDF editor runs its worker and wasm decoders without network access', () => {
    assert.deepEqual(pdf.csp, { 'worker-src': ['blob:'], 'script-src': ["'wasm-unsafe-eval'"] });
    assert.equal(pdf.findWidget, false);
  });
});
