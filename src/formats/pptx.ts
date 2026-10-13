import { MAX_FILE_SIZE } from '../limits.ts';
import { checkOoxmlPackage } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const pptx: FormatDefinition = {
  id: 'pptx',
  extensions: ['.pptx'],
  maxFileSize: MAX_FILE_SIZE.pptx,
  webview: 'pptx',
  check: (bytes) => checkOoxmlPackage(bytes, '.pptx'),
};
