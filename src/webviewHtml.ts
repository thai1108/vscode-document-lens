import type { CspDirective, FormatDefinition } from './formats/types.ts';

export interface WebviewPage {
  /** webview.cspSource in VS Code; 'self' in the browser preview. */
  cspSource: string;
  nonce: string;
  scriptUri: string;
  styleUri: string;
  csp?: FormatDefinition['csp'];
}

/**
 * The strict base policy every viewer runs under: no network, scripts only from the nonce-tagged bundle, styles from
 * the extension (or a <style nonce> a viewer creates), images and fonts also from blob:/data: URLs a viewer builds out
 * of the file. A format adds what its viewer needs through FormatDefinition.csp.
 */
export function contentSecurityPolicy(cspSource: string, nonce: string, extra: WebviewPage['csp'] = {}): string {
  const base: Record<string, string[]> = {
    'default-src': ["'none'"],
    'base-uri': ["'none'"],
    'form-action': ["'none'"],
    'script-src': [`'nonce-${nonce}'`],
    'style-src': [cspSource, `'nonce-${nonce}'`],
    'img-src': [cspSource, 'blob:', 'data:'],
    'font-src': [cspSource, 'blob:', 'data:'],
  };
  for (const [directive, sources] of Object.entries(extra) as [CspDirective, readonly string[]][]) {
    base[directive] = [...new Set([...(base[directive] ?? []), ...sources])];
  }
  return Object.entries(base)
    .map(([directive, sources]) => `${directive} ${sources.join(' ')}`)
    .join('; ');
}

export function webviewHtml(page: WebviewPage): string {
  const csp = contentSecurityPolicy(page.cspSource, page.nonce, page.csp);
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="${page.styleUri}">
<title>Document Lens</title>
</head>
<body>
<div id="app"></div>
<script nonce="${page.nonce}" src="${page.scriptUri}"></script>
</body>
</html>`;
}
