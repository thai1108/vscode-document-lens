// The .pptx editor inside a real VS Code extension host: files are read with the size cap and the OOXML preflight
// applied, open in the presentation editor, and the editor survives the file changing on disk.
import assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readDocument } from '../../src/documentProvider.ts';
import { FORMATS } from '../../src/formats/index.ts';
import { sampleDeck } from '../fixtures/pptx.ts';
import { ooxmlPackage } from '../support/ooxml.ts';
import { buildPptx } from '../support/pptxBuilder.ts';

// Mocha's TDD globals, provided by @vscode/test-cli inside the extension host.
declare function suite(name: string, body: () => void): void;
declare function test(name: string, body: () => Promise<void>): void;
declare function teardown(body: () => Thenable<unknown>): void;

const pptx = FORMATS.find((format) => format.id === 'pptx');

function workspaceUri(file: string): vscode.Uri {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, 'the test workspace folder is open');
  return vscode.Uri.file(path.join(folder.uri.fsPath, file));
}

async function waitFor<T>(read: () => T | undefined, what: string, timeoutMs = 15_000): Promise<T> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const value = read();
    if (value !== undefined) {
      return value;
    }
    if (Date.now() > deadline) {
      throw new Error(`Timed out waiting for ${what}`);
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
}

function presentationTab(uri: vscode.Uri): vscode.TabInputCustom | undefined {
  const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
  return input instanceof vscode.TabInputCustom && input.uri.fsPath === uri.fsPath && input.viewType === 'documentLens.pptx' ? input : undefined;
}

suite('Document Lens: .pptx', () => {
  teardown(() => vscode.commands.executeCommand('workbench.action.closeAllEditors'));

  test('reads the sample presentation through the workspace file system', async () => {
    assert.ok(pptx);
    const bytes = await readDocument(workspaceUri('sample.pptx'), pptx);
    assert.equal(new TextDecoder().decode(bytes.subarray(0, 2)), 'PK');
    assert.equal(pptx.findWidget ?? false, false, 'the viewer has no full-deck search for the find widget');
    assert.equal(pptx.csp, undefined, 'pictures use blob: URLs, which the base policy allows');
  });

  test('refuses a Word document saved with a .pptx extension', async () => {
    assert.ok(pptx);
    const uri = workspaceUri('renamed.pptx');
    const word = ooxmlPackage({
      name: 'word/document.xml',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      xml: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    });
    await vscode.workspace.fs.writeFile(uri, word);
    await assert.rejects(readDocument(uri, pptx), /This \.pptx file is actually a \.docx file; rename it to open it/);
  });

  test('opens in the presentation editor and stays open when the file changes on disk', async () => {
    assert.ok(pptx);
    const uri = workspaceUri('live.pptx');
    await vscode.workspace.fs.writeFile(uri, sampleDeck());
    await vscode.commands.executeCommand('vscode.open', uri);
    await waitFor(() => presentationTab(uri), 'the presentation editor');

    // The editor reloads on the file system watcher's change event; the webview's redraw itself is covered by the
    // happy-dom and Playwright suites, as extension host APIs cannot look into a webview.
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(uri, '..'), 'live.pptx'));
    try {
      let seen = false;
      watcher.onDidChange(() => (seen = true));
      watcher.onDidCreate(() => (seen = true));
      const changed = buildPptx({ slides: [{ layout: 2, shapes: '' }] });
      await vscode.workspace.fs.writeFile(uri, changed);
      await waitFor(() => (seen ? true : undefined), 'the change event the editor reloads on');
      const reloaded = await readDocument(uri, pptx);
      assert.equal(reloaded.byteLength, changed.byteLength, 'the editor’s reload reads the new version and it passes the checks');
    } finally {
      watcher.dispose();
    }
    assert.ok(presentationTab(uri), 'still showing the presentation editor after the change');
  });
});
