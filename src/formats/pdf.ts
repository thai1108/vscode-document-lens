import { MAX_FILE_SIZE } from '../limits.ts';
import { checkPdf } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const pdf: FormatDefinition = {
  id: 'pdf',
  extensions: ['.pdf'],
  maxFileSize: MAX_FILE_SIZE.pdf,
  webview: 'pdf',
  check: (bytes) => checkPdf(bytes),
};
