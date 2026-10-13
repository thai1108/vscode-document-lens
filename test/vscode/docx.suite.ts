// The .docx format inside a real VS Code (pnpm test:vscode): both synthetic samples, including the hostile one, open
// in the Document Lens editor and pass the host's reading path. What the webview shows is covered by
// test/e2e/docx.spec.ts.
import assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readDocument } from '../../src/documentProvider.ts';
import { docx } from '../../src/formats/docx.ts';
import { viewTypeOf } from '../../src/formats/types.ts';
import { documentMessage } from '../../src/payload.ts';

// Mocha's TDD globals, provided by @vscode/test-cli inside the extension host.
declare function suite(name: string, body: () => void): void;
declare function test(name: string, body: () => Promise<void>): void;
declare function teardown(body: () => Thenable<unknown>): void;

const SAMPLES = ['sample.docx', 'hostile.docx'];

/** .vscode-test.mjs opens dist-test/vscode/workspace, written by make-fixtures.ts. */
function workspaceUri(file: string): vscode.Uri {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, 'the test workspace folder is open');
  return vscode.Uri.file(path.join(folder.uri.fsPath, file));
}

async function activeCustomEditor(uri: vscode.Uri, timeoutMs = 15_000): Promise<vscode.TabInputCustom> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
    if (input instanceof vscode.TabInputCustom && input.uri.fsPath === uri.fsPath) {
      return input;
    }
    if (Date.now() > deadline) {
      throw new Error(`Timed out waiting for the custom editor of ${path.basename(uri.fsPath)}`);
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
}

suite('.docx in VS Code', () => {
  teardown(() => vscode.commands.executeCommand('workbench.action.closeAllEditors'));

  for (const fileName of SAMPLES) {
    test(`opens ${fileName} in the ${viewTypeOf(docx)} editor`, async () => {
      const uri = workspaceUri(fileName);
      // On a cold start the workbench can swallow the first open while it restores; try once more before failing.
      for (let attempt = 1; ; attempt++) {
        await vscode.commands.executeCommand('vscode.open', uri);
        try {
          const editor = await activeCustomEditor(uri);
          assert.equal(editor.viewType, viewTypeOf(docx));
          return;
        } catch (error) {
          if (attempt >= 2) {
            throw error;
          }
        }
      }
    });
  }

  test('reads both samples through the size cap and package check into plain Uint8Arrays', async () => {
    for (const fileName of SAMPLES) {
      const bytes = await readDocument(workspaceUri(fileName), docx);
      const message = documentMessage(docx, fileName, bytes);
      assert.equal(message.type, 'document');
      if (message.type === 'document') {
        assert.equal(message.data.constructor, Uint8Array);
        assert.deepEqual([...message.data.subarray(0, 2)], [0x50, 0x4b], `${fileName} is a ZIP package`);
      }
    }
  });

  test('registers the .docx editor with VS Code’s find widget', async () => {
    // The pages are DOM text, so the webview's find widget (Cmd/Ctrl+F) searches them; the provider passes this flag
    // as enableFindWidget, which the extension API does not expose for inspection.
    assert.equal(docx.findWidget, true);
  });
});
