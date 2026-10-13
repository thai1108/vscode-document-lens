// Builds the browser preview (scripts/preview.ts) for a file and serves it over HTTP for Playwright.
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createServer, type Server } from 'node:http';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import { expect, type Page } from '@playwright/test';
import { loadSamples } from '../support/samples.ts';

const ROOT = path.resolve(import.meta.dirname, '../..');
const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.wasm': 'application/wasm',
};

export interface PreviewSite {
  url: string;
  close(): Promise<void>;
}

let built = false;

/** Builds the extension once per test run. */
export function buildOnce(): void {
  if (!built) {
    execFileSync(process.execPath, [path.join(ROOT, 'esbuild.mjs')], { cwd: ROOT, stdio: 'ignore' });
    built = true;
  }
}

/** Writes the synthetic sample `fileName` (see test/fixtures) to a temporary folder and returns its path. */
export async function writeSample(fileName: string): Promise<string> {
  const sample = (await loadSamples()).find((candidate) => candidate.fileName === fileName);
  if (!sample) {
    throw new Error(`No sample named ${fileName} in test/fixtures`);
  }
  const file = path.join(mkdtempSync(path.join(tmpdir(), 'document-lens-sample-')), fileName);
  writeFileSync(file, sample.build());
  return file;
}

export async function servePreview(file: string): Promise<PreviewSite> {
  buildOnce();
  const dir = mkdtempSync(path.join(tmpdir(), 'document-lens-e2e-'));
  execFileSync(process.execPath, [path.join(ROOT, 'scripts/preview.ts'), file, dir], { cwd: ROOT, stdio: 'ignore' });
  const server: Server = createServer((request, response) => {
    const pathname = decodeURIComponent(new URL(request.url ?? '/', 'http://localhost').pathname);
    if (pathname === '/favicon.ico') {
      // Chrome asks for it on its own; a 404 would show up as a console error in every test.
      response.writeHead(204).end();
      return;
    }
    const target = path.join(dir, pathname === '/' ? 'index.html' : pathname);
    if (!target.startsWith(dir)) {
      response.writeHead(403).end();
      return;
    }
    try {
      response.writeHead(200, { 'content-type': TYPES[path.extname(target)] ?? 'application/octet-stream' }).end(readFileSync(target));
    } catch {
      response.writeHead(404).end();
    }
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  const port = typeof address === 'object' && address ? address.port : 0;
  return {
    url: `http://127.0.0.1:${port}/index.html`,
    close: () =>
      new Promise<void>((resolve) => {
        server.close(() => {
          rmSync(dir, { recursive: true, force: true });
          resolve();
        });
      }),
  };
}

/** Console errors, CSP violations, uncaught exceptions and failed requests seen by the page. */
export function watchProblems(page: Page): string[] {
  const problems: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error' || /Content Security Policy|Refused to/i.test(message.text())) {
      problems.push(`console: ${message.text()}`);
    }
  });
  page.on('pageerror', (error) => problems.push(`pageerror: ${error.message}`));
  page.on('response', (response) => {
    if (response.status() >= 400) {
      problems.push(`http ${response.status()}: ${response.url()}`);
    }
  });
  return problems;
}

/** Applies the default styles VS Code injects into every webview, so layout checks see what VS Code shows. */
export async function adoptVsCodeDefaults(page: Page): Promise<void> {
  await page.evaluate(() => {
    const sheet = new CSSStyleSheet();
    sheet.replaceSync('@layer vscode-default { body { margin: 0; padding: 0 20px; } img, video { max-width: 100%; max-height: 100%; } }');
    document.adoptedStyleSheets = [...document.adoptedStyleSheets, sheet];
  });
}

/** The frame every viewer shares: file name shown, no error banner, full-width layout under VS Code's defaults. */
export async function expectShellLoaded(page: Page, fileName: string): Promise<void> {
  await expect(page.locator('#file-name')).toHaveText(fileName);
  await expect(page.locator('#error')).toBeHidden();
  await adoptVsCodeDefaults(page);
  const app = await page.locator('#app').boundingBox();
  expect(app?.x).toBe(0);
  expect(app?.width).toBe(page.viewportSize()?.width);
}
