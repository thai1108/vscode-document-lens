import { MAX_FILE_SIZE } from '../limits.ts';
import { checkText } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const csv: FormatDefinition = {
  id: 'csv',
  extensions: ['.csv'],
  maxFileSize: MAX_FILE_SIZE.csv,
  webview: 'spreadsheet',
  text: true,
  check: (bytes) => checkText(bytes, '.csv'),
};
