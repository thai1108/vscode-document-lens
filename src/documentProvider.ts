import { randomBytes } from 'node:crypto';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { FORMATS } from './formats/index.ts';
import { viewTypeOf, type FormatDefinition } from './formats/types.ts';
import { formatMegabytes } from './limits.ts';
import { documentMessage } from './payload.ts';
import { parseWebviewMessage, type HostMessage } from './protocol.ts';
import { webviewHtml } from './webviewHtml.ts';

const RELOAD_DEBOUNCE_MS = 500;

function fileNameOf(uri: vscode.Uri): string {
  return path.posix.basename(uri.path);
}

/** Reads a file for `format`, refusing it above the format's size cap or when it fails the format's preflight check. */
export async function readDocument(uri: vscode.Uri, format: FormatDefinition): Promise<Uint8Array> {
  const tooLarge = (size: number) =>
    new Error(`${fileNameOf(uri)} is too large to preview (${formatMegabytes(size)}; limit ${formatMegabytes(format.maxFileSize)})`);
  const { size } = await vscode.workspace.fs.stat(uri);
  if (size > format.maxFileSize) {
    throw tooLarge(size);
  }
  const bytes = await vscode.workspace.fs.readFile(uri);
  // The file can grow between stat and read.
  if (bytes.byteLength > format.maxFileSize) {
    throw tooLarge(bytes.byteLength);
  }
  format.check(bytes);
  return bytes;
}

type Contents = { bytes: Uint8Array } | { error: string };

async function load(uri: vscode.Uri, format: FormatDefinition): Promise<Contents> {
  try {
    return { bytes: await readDocument(uri, format) };
  } catch (error) {
    return { error: (error as Error).message };
  }
}

/** An opened file; re-reads itself when the file changes on disk. */
class LensDocument implements vscode.CustomDocument {
  readonly uri: vscode.Uri;
  private readonly format: FormatDefinition;
  private contents: Contents;
  private readonly changed = new vscode.EventEmitter<void>();
  readonly onDidChange = this.changed.event;
  private readonly watcher: vscode.FileSystemWatcher;
  private reloadTimer: NodeJS.Timeout | undefined;

  private constructor(uri: vscode.Uri, format: FormatDefinition, contents: Contents) {
    this.uri = uri;
    this.format = format;
    this.contents = contents;
    this.watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(uri, '..'), fileNameOf(uri)));
    this.watcher.onDidChange(() => this.scheduleReload());
    this.watcher.onDidCreate(() => this.scheduleReload());
  }

  static async open(uri: vscode.Uri, format: FormatDefinition): Promise<LensDocument> {
    return new LensDocument(uri, format, await load(uri, format));
  }

  message(): HostMessage {
    const fileName = fileNameOf(this.uri);
    return 'bytes' in this.contents
      ? documentMessage(this.format, fileName, this.contents.bytes)
      : { type: 'error', fileName, message: this.contents.error };
  }

  private scheduleReload(): void {
    clearTimeout(this.reloadTimer);
    this.reloadTimer = setTimeout(() => void this.reload(), RELOAD_DEBOUNCE_MS);
  }

  private async reload(): Promise<void> {
    const contents = await load(this.uri, this.format);
    // A file that stops opening (e.g. half-written) keeps showing its last good version.
    if ('error' in contents && 'bytes' in this.contents) {
      void vscode.window.showWarningMessage(`Could not reload ${fileNameOf(this.uri)}: ${contents.error}`);
      return;
    }
    this.contents = contents;
    this.changed.fire();
  }

  dispose(): void {
    clearTimeout(this.reloadTimer);
    this.watcher.dispose();
    this.changed.dispose();
  }
}

/** Read-only custom editor for one format; the viewer itself runs in the webview. */
export class DocumentLensProvider implements vscode.CustomReadonlyEditorProvider<LensDocument> {
  private readonly extensionUri: vscode.Uri;
  private readonly format: FormatDefinition;

  private constructor(extensionUri: vscode.Uri, format: FormatDefinition) {
    this.extensionUri = extensionUri;
    this.format = format;
  }

  static register(context: vscode.ExtensionContext): vscode.Disposable[] {
    return FORMATS.map((format) =>
      vscode.window.registerCustomEditorProvider(viewTypeOf(format), new DocumentLensProvider(context.extensionUri, format), {
        webviewOptions: { retainContextWhenHidden: true, enableFindWidget: format.findWidget ?? false },
        supportsMultipleEditorsPerDocument: true,
      }),
    );
  }

  openCustomDocument(uri: vscode.Uri): Promise<LensDocument> {
    return LensDocument.open(uri, this.format);
  }

  resolveCustomEditor(document: LensDocument, panel: vscode.WebviewPanel): void {
    const { webview } = panel;
    const dist = vscode.Uri.joinPath(this.extensionUri, 'dist');
    const asset = (extension: string) => webview.asWebviewUri(vscode.Uri.joinPath(dist, `webview-${this.format.webview}.${extension}`)).toString();
    webview.options = { enableScripts: true, localResourceRoots: [dist] };
    webview.html = webviewHtml({
      cspSource: webview.cspSource,
      nonce: randomBytes(16).toString('base64'),
      scriptUri: asset('js'),
      styleUri: asset('css'),
      csp: this.format.csp,
    });

    const post = () => void webview.postMessage(document.message());
    const subscriptions: vscode.Disposable[] = [
      document.onDidChange(post),
      webview.onDidReceiveMessage((raw: unknown) => {
        const message = parseWebviewMessage(raw);
        switch (message?.type) {
          case 'ready':
            post();
            break;
          case 'openLink':
            void vscode.env.openExternal(vscode.Uri.parse(message.url, true));
            break;
        }
      }),
    ];
    panel.onDidDispose(() => subscriptions.forEach((subscription) => subscription.dispose()));
  }
}
