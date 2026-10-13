// The .doc viewer inside a real VS Code extension host: files are read through VS Code's file system with the host
// check applied, and the viewer's parser (browser code, but free of DOM APIs) reads them in the same runtime.
import assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readDocument } from '../../src/documentProvider.ts';
import { doc } from '../../src/formats/doc.ts';
import { layoutDocument, type Block } from '../../src/webview/doc/layout.ts';
import { readWordDocument } from '../../src/webview/doc/wordDocument.ts';
import { writeWordDocument } from '../support/wordWriter.ts';
import { writeZip } from '../support/zipWriter.ts';

// Mocha's TDD globals, provided by @vscode/test-cli inside the extension host.
declare function suite(name: string, body: () => void): void;
declare function test(name: string, body: () => Promise<void>): void;
declare function teardown(body: () => Thenable<unknown>): void;

function workspaceUri(file: string): vscode.Uri {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, 'the test workspace folder is open');
  return vscode.Uri.file(path.join(folder.uri.fsPath, file));
}

function paragraphs(blocks: Block[]): string[] {
  return blocks.flatMap((block) => (block.type === 'paragraph' ? [block.runs.map((run) => run.text).join('')] : []));
}

/** Opens `uri` with VS Code's default editor; a workbench still restoring on a cold start may swallow the first try. */
async function openedViewType(uri: vscode.Uri): Promise<string | undefined> {
  for (let attempt = 0; attempt < 2; attempt++) {
    await vscode.commands.executeCommand('vscode.open', uri);
    const deadline = Date.now() + 15_000;
    while (Date.now() < deadline) {
      const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
      if (input instanceof vscode.TabInputCustom && input.uri.fsPath === uri.fsPath) {
        return input.viewType;
      }
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
  }
  return undefined;
}

suite('.doc preview in VS Code', () => {
  teardown(() => vscode.commands.executeCommand('workbench.action.closeAllEditors'));

  test('sample.doc is read through VS Code and parses to its text, tables and Japanese', async () => {
    const bytes = await readDocument(workspaceUri('sample.doc'), doc);
    const layout = layoutDocument(readWordDocument(bytes));
    const main = layout.stories.find((story) => story.kind === 'main');
    assert.ok(main);
    const text = paragraphs(main.blocks);
    assert.equal(text[0], 'Document Lens sample');
    assert.ok(text.includes('Prices: 5 € – “quoted” text.'));
    assert.ok(text.includes('Tracked changes: only the kept text shows.'));
    assert.ok(text.includes('日本語のテキストも表示されます。'));
    assert.equal(main.blocks.filter((block) => block.type === 'table').length, 1);
    assert.equal(layout.truncated, false);
  });

  test('a .docx renamed to .doc is refused before it reaches the viewer', async () => {
    const uri = workspaceUri('renamed.doc');
    await vscode.workspace.fs.writeFile(uri, writeZip([{ name: '[Content_Types].xml', data: '<Types/>' }]));
    await assert.rejects(readDocument(uri, doc), /This \.doc file is a ZIP package, probably a \.docx file with the wrong extension/);
  });

  test('a password-protected .doc passes the host check and opens in the Document Lens editor', async () => {
    const uri = workspaceUri('locked.doc');
    await vscode.workspace.fs.writeFile(uri, writeWordDocument({ body: ['Secret\r'], encrypted: true }));
    // The host only checks the signature; the viewer explains the encryption in its error banner.
    const bytes = await readDocument(uri, doc);
    assert.throws(() => readWordDocument(bytes), /password-protected/);
    assert.equal(await openedViewType(uri), 'documentLens.doc');
  });
});
