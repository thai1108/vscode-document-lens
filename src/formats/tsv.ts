import { MAX_FILE_SIZE } from '../limits.ts';
import { checkText } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const tsv: FormatDefinition = {
  id: 'tsv',
  extensions: ['.tsv'],
  maxFileSize: MAX_FILE_SIZE.tsv,
  webview: 'spreadsheet',
  text: true,
  check: (bytes) => checkText(bytes, '.tsv'),
};
