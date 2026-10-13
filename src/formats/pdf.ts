import { MAX_FILE_SIZE } from '../limits.ts';
import { checkPdf } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const pdf: FormatDefinition = {
  id: 'pdf',
  extensions: ['.pdf'],
  maxFileSize: MAX_FILE_SIZE.pdf,
  webview: 'pdf',
  csp: {
    // pdf.js parses the file in a worker started from the source embedded in the bundle (a blob: URL).
    'worker-src': ['blob:'],
    // JBIG2, CCITT fax (common in scans) and JPEG 2000 images are decoded by WebAssembly modules embedded in the
    // bundle; this allows compiling WebAssembly, not JavaScript eval.
    'script-src': ["'wasm-unsafe-eval'"],
  },
  // The viewer has its own find bar, which searches every page and not only the rendered ones.
  findWidget: false,
  check: (bytes) => checkPdf(bytes),
};
