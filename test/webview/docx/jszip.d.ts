// The declaration of src/webview/docx/jszip.d.ts for this program, which only includes the files under test/webview
// (the viewer's modules import JSZip's browser build by path). Keep the two in step.
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
