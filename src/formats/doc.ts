import { MAX_FILE_SIZE } from '../limits.ts';
import { checkCompoundFile } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const doc: FormatDefinition = {
  id: 'doc',
  extensions: ['.doc'],
  maxFileSize: MAX_FILE_SIZE.doc,
  webview: 'doc',
  // The preview is plain text in the DOM, so VS Code's find widget searches it.
  findWidget: true,
  check: (bytes) => checkCompoundFile(bytes, '.doc'),
};
