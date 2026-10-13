// .csv and .tsv inside a real VS Code: they open in the Document Lens grid by default, "Reopen as Text" hands the
// file to VS Code's text editor, and the host reads them with their checks applied (what the webview shows is covered
// by test/e2e/csv.spec.ts).
import assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { OPEN_AS_TEXT_COMMAND, readDocument } from '../../src/documentProvider.ts';
import { FORMATS } from '../../src/formats/index.ts';
import { openByDefault, waitFor, workspaceUri } from '../support/vscodeEditors.ts';

// Mocha's TDD globals, provided by @vscode/test-cli inside the extension host.
declare function suite(name: string, body: () => void): void;
declare function test(name: string, body: () => Promise<void>): void;
declare function suiteTeardown(body: () => Thenable<unknown>): void;
declare function teardown(body: () => Thenable<unknown>): void;

function format(id: string) {
  const found = FORMATS.find((candidate) => candidate.id === id);
  assert.ok(found, id);
  return found;
}

function activeTextEditorOf(uri: vscode.Uri): vscode.TextEditor | undefined {
  const editor = vscode.window.activeTextEditor;
  return editor?.document.uri.fsPath === uri.fsPath ? editor : undefined;
}

const written: vscode.Uri[] = [];
async function writeWorkspaceFile(name: string, bytes: Uint8Array): Promise<vscode.Uri> {
  const uri = workspaceUri(name);
  await vscode.workspace.fs.writeFile(uri, bytes);
  written.push(uri);
  return uri;
}

suite('Document Lens .csv and .tsv in VS Code', () => {
  teardown(() => vscode.commands.executeCommand('workbench.action.closeAllEditors'));
  suiteTeardown(async () => {
    for (const uri of written) {
      await vscode.workspace.fs.delete(uri);
    }
  });

  for (const [file, id] of [
    ['sample.csv', 'csv'],
    ['shift-jis.csv', 'csv'],
    ['semicolon.csv', 'csv'],
    ['sample.tsv', 'tsv'],
  ] as const) {
    test(`opens ${file} in the grid by default`, async () => {
      const tab = await openByDefault(workspaceUri(file));
      assert.equal(tab.viewType, `documentLens.${id}`);
      const bytes = await readDocument(workspaceUri(file), format(id));
      assert.ok(bytes.byteLength > 0);
    });
  }

  test('Reopen as Text opens the file in the text editor, with its text decoded', async () => {
    const uri = workspaceUri('sample.csv');
    await openByDefault(uri);
    await vscode.commands.executeCommand(OPEN_AS_TEXT_COMMAND);
    const editor = await waitFor(() => activeTextEditorOf(uri), 'the text editor of sample.csv');
    assert.equal(editor.document.lineAt(0).text, 'ID,Item,Price,Note,Code');
    assert.ok(vscode.window.tabGroups.activeTabGroup.activeTab?.input instanceof vscode.TabInputText);
  });

  test('Reopen as Text takes the file from the Explorer and is only for text formats', async () => {
    const uri = workspaceUri('sample.tsv');
    await vscode.commands.executeCommand(OPEN_AS_TEXT_COMMAND, uri);
    await waitFor(() => activeTextEditorOf(uri), 'the text editor of sample.tsv');

    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
    const pdf = workspaceUri('sample.pdf');
    await openByDefault(pdf);
    await vscode.commands.executeCommand(OPEN_AS_TEXT_COMMAND);
    await new Promise((resolve) => setTimeout(resolve, 500));
    const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
    assert.ok(input instanceof vscode.TabInputCustom && input.viewType === 'documentLens.pdf', 'a PDF stays in its viewer');
  });

  test('refuses workbooks and binary files under a text extension, with a message saying what they are', async () => {
    const xlsx = await writeWorkspaceFile('renamed-workbook.csv', await vscode.workspace.fs.readFile(workspaceUri('sample.xlsx')));
    await assert.rejects(readDocument(xlsx, format('csv')), /This \.csv file is a ZIP package, probably an \.xlsx workbook/);
    const xls = await writeWorkspaceFile('renamed-workbook.tsv', await vscode.workspace.fs.readFile(workspaceUri('sample.xls')));
    await assert.rejects(readDocument(xls, format('tsv')), /This \.tsv file is an Office binary file/);
    const binary = await writeWorkspaceFile('picture.csv', new Uint8Array(4096).map((_byte, i) => (i % 3 ? 7 : 0)));
    await assert.rejects(readDocument(binary, format('csv')), /contains binary data, not text/);
    assert.equal(path.extname(binary.fsPath), '.csv');
  });
});
