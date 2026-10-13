// JSZip's browser build, the same file docx-preview's `import JSZip from 'jszip'` resolves to in the bundle. Imported
// by this path, with only what the viewer uses declared here, because the package's own typings pull Node's types
// (Buffer, NodeJS.Timeout, ...) into the webview program.
declare module 'jszip/dist/jszip.min.js' {
  interface JSZipEntry {
    readonly name: string;
    readonly dir: boolean;
  }
  interface JSZipArchive {
    /** Every entry by its name as JSZip normalized it; folders end with a slash. */
    readonly files: Readonly<Record<string, JSZipEntry>>;
  }
  const JSZip: {
    /** Reads the central directory and local headers without inflating anything. */
    loadAsync(data: Uint8Array): Promise<JSZipArchive>;
  };
  export default JSZip;
}
