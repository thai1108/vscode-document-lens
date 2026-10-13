import { MAX_FILE_SIZE } from '../limits.ts';
import { checkOoxmlPackage } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const xlsx: FormatDefinition = {
  id: 'xlsx',
  extensions: ['.xlsx'],
  maxFileSize: MAX_FILE_SIZE.xlsx,
  webview: 'spreadsheet',
  check: (bytes) => checkOoxmlPackage(bytes, '.xlsx'),
};
