// The .ppt format inside a real VS Code extension host: the file opens in its Document Lens editor, and the bytes
// VS Code reads (with the size cap and check applied) are exactly what the viewer's parser turns into slides.
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readDocument } from '../../src/documentProvider.ts';
import { FORMATS } from '../../src/formats/index.ts';
import { documentMessage } from '../../src/payload.ts';
import { PptError } from '../../src/webview/ppt/errors.ts';
import { parsePresentation } from '../../src/webview/ppt/presentation.ts';
import { SAMPLE_PPT } from '../fixtures/ppt.ts';
import { writePpt } from '../support/pptWriter.ts';
import { writeZip } from '../support/zipWriter.ts';

// Mocha's TDD globals, provided by @vscode/test-cli inside the extension host.
declare function suite(name: string, body: () => void): void;
declare function test(name: string, body: () => Promise<void>): void;
declare function teardown(body: () => Thenable<unknown>): void;

const ppt = FORMATS.find((format) => format.id === 'ppt');

function workspaceUri(file: string): vscode.Uri {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, 'the test workspace folder is open');
  return vscode.Uri.file(path.join(folder.uri.fsPath, file));
}

/** A file outside the workspace, written for one test. */
function scratchUri(name: string, bytes: Uint8Array): vscode.Uri {
  const file = path.join(mkdtempSync(path.join(tmpdir(), 'dl-ppt-')), name);
  writeFileSync(file, bytes);
  return vscode.Uri.file(file);
}

async function openCustomEditor(uri: vscode.Uri): Promise<vscode.TabInputCustom> {
  for (let attempt = 1; ; attempt++) {
    await vscode.commands.executeCommand('vscode.open', uri);
    const deadline = Date.now() + 15_000;
    while (Date.now() < deadline) {
      const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
      if (input instanceof vscode.TabInputCustom && input.uri.fsPath === uri.fsPath) {
        return input;
      }
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    // On a cold start the workbench can swallow the first open.
    if (attempt >= 2) {
      throw new Error(`No custom editor tab for ${path.basename(uri.fsPath)}`);
    }
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
  }
}

suite('.ppt presentations in VS Code', () => {
  teardown(() => vscode.commands.executeCommand('workbench.action.closeAllEditors'));

  test('opens sample.ppt in the Document Lens .ppt editor', async () => {
    const tab = await openCustomEditor(workspaceUri('sample.ppt'));
    assert.equal(tab.viewType, 'documentLens.ppt');
    assert.equal(ppt?.findWidget, true);
  });

  test('hands the viewer bytes that parse into the sample slides', async () => {
    assert.ok(ppt);
    const bytes = await readDocument(workspaceUri('sample.ppt'), ppt);
    const message = documentMessage(ppt, 'sample.ppt', bytes);
    assert.equal(message.type, 'document');
    if (message.type === 'document') {
      assert.equal(message.data.constructor, Uint8Array);
      const presentation = parsePresentation(message.data);
      assert.deepEqual(
        presentation.slides.map((slide) => slide.texts[0]?.paragraphs[0]?.text),
        ['Quarterly review', 'Highlights', 'Backup slide'],
      );
      assert.deepEqual(
        presentation.slides.map((slide) => slide.hidden),
        [false, false, true],
      );
      assert.equal(presentation.slides[1]?.pictures[0]?.mime, 'image/png');
    }
  });

  test('lets a password-protected .ppt through to the viewer, which explains it', async () => {
    assert.ok(ppt);
    const uri = scratchUri('locked.ppt', writePpt({ ...SAMPLE_PPT, encrypted: { token: true } }));
    const bytes = await readDocument(uri, ppt);
    assert.throws(() => parsePresentation(bytes), (error: Error) => error instanceof PptError && error.kind === 'encrypted');
    const tab = await openCustomEditor(uri);
    assert.equal(tab.viewType, 'documentLens.ppt');
  });

  test('refuses a .pptx package saved as .ppt before it reaches the viewer', async () => {
    assert.ok(ppt);
    const uri = scratchUri('renamed.ppt', writeZip([{ name: '[Content_Types].xml', data: '<Types/>' }]));
    await assert.rejects(readDocument(uri, ppt), /This \.ppt file is a ZIP package, probably a \.pptx file/);
  });
});
