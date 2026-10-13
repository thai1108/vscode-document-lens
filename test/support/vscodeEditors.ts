// Helpers for suites that run inside the VS Code extension host (test/vscode/*.suite.ts).
import assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';

/** A file in dist-test/vscode/workspace, which .vscode-test.mjs opens and make-fixtures.ts fills. */
export function workspaceUri(file: string): vscode.Uri {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, 'the test workspace folder is open');
  return vscode.Uri.file(path.join(folder.uri.fsPath, file));
}

export async function waitFor<T>(read: () => T | undefined, what: string, timeoutMs = 15_000): Promise<T> {
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

function activeCustomTab(uri: vscode.Uri): vscode.TabInputCustom | undefined {
  const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
  return input instanceof vscode.TabInputCustom && input.uri.fsPath === uri.fsPath ? input : undefined;
}

/** Opens `uri` with VS Code's default editor resolution, retrying once for a workbench that is still restoring. */
export async function openByDefault(uri: vscode.Uri): Promise<vscode.TabInputCustom> {
  for (let attempt = 1; ; attempt++) {
    await vscode.commands.executeCommand('vscode.open', uri);
    try {
      return await waitFor(() => activeCustomTab(uri), `custom editor tab for ${path.basename(uri.fsPath)}`);
    } catch (error) {
      if (attempt >= 2) {
        throw error;
      }
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
    }
  }
}
