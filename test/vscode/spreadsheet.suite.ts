// The spreadsheet formats inside a real VS Code: .xlsx and .xls open in their Document Lens editors, and the host reads
// them with their checks applied (what the webview then shows is covered by test/e2e/spreadsheet.spec.ts).
import assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readDocument } from '../../src/documentProvider.ts';
import { FORMATS } from '../../src/formats/index.ts';
import { documentMessage } from '../../src/payload.ts';

// Mocha's TDD globals, provided by @vscode/test-cli inside the extension host.
declare function suite(name: string, body: () => void): void;
declare function test(name: string, body: () => Promise<void>): void;
declare function suiteTeardown(body: () => Thenable<unknown>): void;
declare function teardown(body: () => Thenable<unknown>): void;

function workspaceUri(file: string): vscode.Uri {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, 'the test workspace folder is open');
  return vscode.Uri.file(path.join(folder.uri.fsPath, file));
}

function format(id: string) {
  const found = FORMATS.find((candidate) => candidate.id === id);
  assert.ok(found, id);
  return found;
}

async function activeCustomTab(uri: vscode.Uri, timeoutMs = 15_000): Promise<vscode.TabInputCustom> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
    if (input instanceof vscode.TabInputCustom && input.uri.fsPath === uri.fsPath) {
      return input;
    }
    if (Date.now() > deadline) {
      throw new Error(`Timed out waiting for the editor of ${path.basename(uri.fsPath)}`);
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
}

const written: vscode.Uri[] = [];
async function writeWorkspaceFile(name: string, bytes: Uint8Array): Promise<vscode.Uri> {
  const uri = workspaceUri(name);
  await vscode.workspace.fs.writeFile(uri, bytes);
  written.push(uri);
  return uri;
}

suite('Document Lens spreadsheets in VS Code', () => {
  teardown(() => vscode.commands.executeCommand('workbench.action.closeAllEditors'));
  suiteTeardown(async () => {
    for (const uri of written) {
      await vscode.workspace.fs.delete(uri);
    }
  });

  for (const [file, id] of [
    ['sample.xlsx', 'xlsx'],
    ['sample.xls', 'xls'],
  ] as const) {
    test(`opens ${file} in the spreadsheet viewer and sends its bytes as they are`, async () => {
      const uri = workspaceUri(file);
      await vscode.commands.executeCommand('vscode.open', uri);
      assert.equal((await activeCustomTab(uri)).viewType, `documentLens.${id}`);
      const bytes = await readDocument(uri, format(id));
      const message = documentMessage(format(id), file, bytes);
      assert.ok(message.type === 'document' && message.data.constructor === Uint8Array);
      const signature = id === 'xlsx' ? [0x50, 0x4b, 0x03, 0x04] : [0xd0, 0xcf, 0x11, 0xe0];
      assert.deepEqual([...bytes.subarray(0, 4)], signature);
    });
  }

  test('refuses files that are not what their .xls/.xlsx extension says, with a message saying what they are', async () => {
    const webPage = await writeWorkspaceFile('export-as-web-page.xls', new TextEncoder().encode('<html><body><table><tr><td>1</td></tr></table></body></html>'));
    await assert.rejects(readDocument(webPage, format('xls')), /This \.xls file is a web page or XML spreadsheet with an \.xls extension/);
    const xlsxNamedXls = await writeWorkspaceFile('renamed.xls', await vscode.workspace.fs.readFile(workspaceUri('sample.xlsx')));
    await assert.rejects(readDocument(xlsxNamedXls, format('xls')), /This \.xls file is a ZIP package, probably a \.xlsx file/);
    const xlsNamedXlsx = await writeWorkspaceFile('renamed.xlsx', await vscode.workspace.fs.readFile(workspaceUri('sample.xls')));
    await assert.rejects(readDocument(xlsNamedXlsx, format('xlsx')), /older binary file, probably a \.xls file/);
  });

  test('still opens a refused file in its editor, to show the message there', async () => {
    const uri = await writeWorkspaceFile('broken-export.xls', new TextEncoder().encode('<?xml version="1.0"?><Workbook/>'));
    await vscode.commands.executeCommand('vscode.open', uri);
    assert.equal((await activeCustomTab(uri)).viewType, 'documentLens.xls');
  });
});
