import { MAX_FILE_SIZE } from '../limits.ts';
import { checkCompoundFile } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const xls: FormatDefinition = {
  id: 'xls',
  extensions: ['.xls'],
  maxFileSize: MAX_FILE_SIZE.xls,
  webview: 'spreadsheet',
  check: (bytes) => checkCompoundFile(bytes, '.xls'),
};
