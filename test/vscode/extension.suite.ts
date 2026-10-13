// Runs inside a real VS Code extension host (pnpm test:vscode runs every test/vscode/*.suite.ts): every format opens
// in its Document Lens editor by default, and files are read through VS Code's file system with the size cap and
// preflight check applied.
import assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readDocument } from '../../src/documentProvider.ts';
import { FORMATS } from '../../src/formats/index.ts';
import { viewTypeOf } from '../../src/formats/types.ts';
import { documentMessage } from '../../src/payload.ts';

// Mocha's TDD globals, provided by @vscode/test-cli inside the extension host.
declare function suite(name: string, body: () => void): void;
declare function test(name: string, body: () => Promise<void>): void;
declare function teardown(body: () => Thenable<unknown>): void;

/** .vscode-test.mjs opens dist-test/vscode/workspace, written by make-fixtures.ts. */
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
      throw new Error(`Timed out waiting for ${what}; open tabs: ${describeTabs()}`);
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
}

function describeTabs(): string {
  return JSON.stringify(
    vscode.window.tabGroups.all.flatMap((group) =>
      group.tabs.map((tab) => ({
        label: tab.label,
        active: tab.isActive && group.isActive,
        input: tab.input instanceof vscode.TabInputCustom ? `custom:${tab.input.viewType}` : tab.input instanceof vscode.TabInputText ? 'text' : String((tab.input as { constructor?: { name?: string } } | undefined)?.constructor?.name),
      })),
    ),
  );
}

function activeCustomTab(uri: vscode.Uri): vscode.TabInputCustom | undefined {
  const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
  return input instanceof vscode.TabInputCustom && input.uri.fsPath === uri.fsPath ? input : undefined;
}

/**
 * Opens `uri` with VS Code's default editor resolution. On a cold first start the workbench can still be restoring
 * and swallow the first open, so a second attempt is made before failing.
 */
async function openByDefault(uri: vscode.Uri): Promise<vscode.TabInputCustom> {
  for (let attempt = 1; ; attempt++) {
    await vscode.commands.executeCommand('vscode.open', uri);
    try {
      return await waitFor(() => activeCustomTab(uri), `custom editor tab for ${path.basename(uri.fsPath)}`, 15_000);
    } catch (error) {
      if (attempt >= 2) {
        throw error;
      }
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
    }
  }
}

suite('Document Lens in VS Code', () => {
  teardown(() => vscode.commands.executeCommand('workbench.action.closeAllEditors'));

  for (const format of FORMATS) {
    const fileName = `sample${format.extensions[0]}`;
    test(`opens ${fileName} in the ${viewTypeOf(format)} editor by default`, async () => {
      const uri = workspaceUri(fileName);
      const tab = await openByDefault(uri);
      assert.equal(tab.viewType, viewTypeOf(format));
      const extension = vscode.extensions.getExtension('thai1108.document-lens');
      assert.ok(extension, 'extension is installed in the test host');
      await waitFor(() => (extension.isActive ? true : undefined), 'extension activation');
    });
  }

  test('reads files through the VS Code file system and sends them as plain Uint8Arrays', async () => {
    const pdf = FORMATS.find((format) => format.id === 'pdf');
    assert.ok(pdf);
    const bytes = await readDocument(workspaceUri('sample.pdf'), pdf);
    const message = documentMessage(pdf, 'sample.pdf', bytes);
    assert.equal(message.type, 'document');
    if (message.type === 'document') {
      // A Node Buffer would reach the webview as { type: 'Buffer', data: [...] } (see payload.ts).
      assert.equal(message.data.constructor, Uint8Array);
      assert.equal(message.data.buffer.byteLength, message.data.byteLength);
      assert.equal(new TextDecoder().decode(message.data.subarray(0, 5)), '%PDF-');
    }
  });

  test('refuses oversized files before reading them and files that fail the preflight check', async () => {
    const byId = (id: string) => {
      const format = FORMATS.find((candidate) => candidate.id === id);
      assert.ok(format, id);
      return format;
    };
    await assert.rejects(readDocument(workspaceUri('oversized.pdf'), byId('pdf')), /oversized\.pdf is too large to preview \(256 MB; limit 256 MB\)/);
    await assert.rejects(readDocument(workspaceUri('broken.docx'), byId('docx')), /Not a valid \.docx file/);
    await assert.rejects(readDocument(workspaceUri('missing.pdf'), byId('pdf')));
  });

  test('opens a broken file in the custom editor instead of failing the editor', async () => {
    const uri = workspaceUri('broken.docx');
    const tab = await openByDefault(uri);
    assert.equal(tab.viewType, 'documentLens.docx');
  });
});
