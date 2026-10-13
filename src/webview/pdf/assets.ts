// The worker and binary data pdf.js needs at runtime, embedded in the bundle by build.mjs: the webview may not fetch
// anything and cannot start a worker from an extension file.

/** Base64 file contents by pdf.js option name (cMapUrl, standardFontDataUrl, wasmUrl) and file name. */
export type EmbeddedFiles = Readonly<Record<string, Readonly<Record<string, string>>>>;

/**
 * A BinaryDataFactory class for getDocument(): pdf.js constructs it and calls fetch() whenever the worker needs a
 * CMap, standard font or wasm decoder. File names come from the document (e.g. a CMap name), so only own keys match.
 */
export function embeddedDataFactory(files: EmbeddedFiles) {
  return class EmbeddedDataFactory {
    async fetch({ kind, filename }: { kind: string; filename: string }): Promise<Uint8Array> {
      const folder = Object.hasOwn(files, kind) ? files[kind] : undefined;
      const base64 = folder && Object.hasOwn(folder, filename) ? folder[filename] : undefined;
      if (base64 === undefined) {
        throw new Error(`No embedded ${kind} file named ${filename}`);
      }
      return decodeBase64(base64);
    }
  };
}

export function decodeBase64(text: string): Uint8Array {
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/** Starts pdf.js's worker module from its source; the CSP allows workers from blob: URLs only. */
export function startWorker(source: string): Worker {
  const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
  return new Worker(url, { type: 'module', name: 'pdf.js' });
}
