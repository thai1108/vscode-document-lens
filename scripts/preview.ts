// Builds a static copy of a viewer webview for one file so the UI can be checked in a normal browser:
// `pnpm preview <file> <out-dir>`, then serve <out-dir> over HTTP and open index.html. The output contains the file,
// so write it outside the repository.
//
// The page is the real webview page (same HTML and CSP, with 'self' in place of VS Code's resource origin) plus a
// mock host. The host message is built by the extension's own code and stored exactly as VS Code's serializer sends
// it; the mock revives it with the same deserializer, so payload bugs that only show up inside VS Code (e.g. Node
// Buffers) show up here too.
import { randomBytes } from 'node:crypto';
import { cpSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import * as path from 'node:path';
import { formatForFile } from '../src/formats/index.ts';
import { documentMessage } from '../src/payload.ts';
import type { HostMessage } from '../src/protocol.ts';
import { webviewHtml } from '../src/webviewHtml.ts';
import { deserializeInWebview, serializeForWebview } from '../test/support/vscodeWebviewSerializer.ts';

const [file, outDir] = process.argv.slice(2);
if (!file || !outDir) {
  console.error('usage: pnpm preview <file> <out-dir>');
  process.exit(2);
}
const format = formatForFile(file);
if (!format) {
  console.error(`No Document Lens format for ${file}`);
  process.exit(2);
}

const root = path.resolve(import.meta.dirname, '..');
const fileName = path.basename(file);
const bytes = readFileSync(file);
let message: HostMessage;
try {
  if (bytes.byteLength > format.maxFileSize) {
    throw new Error(`${fileName} is too large to preview`);
  }
  format.check(bytes);
  message = documentMessage(format, fileName, bytes);
} catch (error) {
  message = { type: 'error', fileName, message: (error as Error).message };
}

mkdirSync(outDir, { recursive: true });
// The viewer bundle and anything it loads, at the same relative paths as inside VS Code.
cpSync(path.join(root, 'dist'), path.join(outDir, 'dist'), { recursive: true });

const serialized = serializeForWebview(message);
writeFileSync(
  path.join(outDir, 'data.js'),
  `window.previewMessage = ${JSON.stringify({ message: serialized.message, buffers: serialized.buffers.map((buffer) => Buffer.from(buffer).toString('base64')) })};\n`,
);

const mock = `
${serializeForWebview.toString()}
${deserializeInWebview.toString()}
function decodeBase64(text) {
  if (Uint8Array.fromBase64) {
    return Uint8Array.fromBase64(text);
  }
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}
function hostMessage() {
  const stored = window.previewMessage;
  return deserializeInWebview({ message: stored.message, buffers: stored.buffers.map(decodeBase64) });
}
// Re-sends the document as the host does when the file changes on disk.
window.previewReload = () => window.dispatchEvent(new MessageEvent('message', { data: hostMessage() }));
window.previewLinks = [];
window.acquireVsCodeApi = () => ({
  postMessage(message) {
    // What the extension host would receive (structured clone out of the webview, then VS Code's serializer).
    const received = deserializeInWebview(serializeForWebview(structuredClone(message)));
    if (received.type === 'ready') {
      setTimeout(window.previewReload);
    } else if (received.type === 'openLink') {
      window.previewLinks.push(received.url);
      console.log('[openLink]', received.url);
    } else if (received.type === 'openAsText') {
      window.previewOpenedAsText = (window.previewOpenedAsText ?? 0) + 1;
      console.log('[openAsText]');
    }
  },
  getState: () => JSON.parse(sessionStorage.getItem('document-lens-state') || 'null') ?? undefined,
  setState: (state) => sessionStorage.setItem('document-lens-state', JSON.stringify(state)),
});
`;
writeFileSync(path.join(outDir, 'mock.js'), mock);

const nonce = randomBytes(16).toString('base64');
const page = webviewHtml({
  cspSource: "'self'",
  nonce,
  scriptUri: `dist/webview-${format.webview}.js`,
  styleUri: `dist/webview-${format.webview}.css`,
  csp: format.csp,
}).replace(`<script nonce="${nonce}"`, `<script nonce="${nonce}" src="data.js"></script>\n<script nonce="${nonce}" src="mock.js"></script>\n<script nonce="${nonce}"`);
writeFileSync(path.join(outDir, 'index.html'), page);
console.log(`Wrote a ${format.id} preview of ${fileName} to ${outDir}${message.type === 'error' ? ` (host error: ${message.message})` : ''}`);
