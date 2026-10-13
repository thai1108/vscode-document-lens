/** CSP directives a viewer may extend beyond the strict base policy (see webviewHtml.ts). */
export type CspDirective = 'script-src' | 'style-src' | 'img-src' | 'font-src' | 'media-src' | 'worker-src' | 'connect-src';

/** One file type the extension previews, opened by its own custom editor `documentLens.<id>`. */
export interface FormatDefinition {
  readonly id: string;
  /** Lower-case file extensions including the dot; package.json declares the same selectors. */
  readonly extensions: readonly string[];
  /** Larger files are refused before they are read. */
  readonly maxFileSize: number;
  /** Viewer bundle: src/webview/<webview>/main.ts, built to dist/webview-<webview>.js and .css. */
  readonly webview: string;
  /** Sources added to the base CSP for this viewer only, e.g. pdf.js needs `worker-src blob:`. */
  readonly csp?: Readonly<Partial<Record<CspDirective, readonly string[]>>>;
  /** VS Code's find widget (Cmd/Ctrl+F) over the webview's text; leave off when the viewer has its own search. */
  readonly findWidget?: boolean;
  /** A text format: the viewer may ask to reopen the file in VS Code's text editor. */
  readonly text?: boolean;
  /** Structural check of the file before it reaches the webview; throws an Error with a message for the user. */
  check(bytes: Uint8Array): void;
}

export function viewTypeOf(format: FormatDefinition): string {
  return `documentLens.${format.id}`;
}
