import { MAX_FILE_SIZE } from '../limits.ts';
import { checkOoxmlPackage } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const docx: FormatDefinition = {
  id: 'docx',
  extensions: ['.docx'],
  maxFileSize: MAX_FILE_SIZE.docx,
  webview: 'docx',
  check: (bytes) => checkOoxmlPackage(bytes, '.docx'),
};
