import { MAX_FILE_SIZE } from '../limits.ts';
import { checkCompoundFile } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const ppt: FormatDefinition = {
  id: 'ppt',
  extensions: ['.ppt'],
  maxFileSize: MAX_FILE_SIZE.ppt,
  webview: 'ppt',
  findWidget: true,
  check: (bytes) => checkCompoundFile(bytes, '.ppt'),
};
