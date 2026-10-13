import { MAX_FILE_SIZE } from '../limits.ts';
import { checkCompoundFile, isCompoundFile, isZip } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const xls: FormatDefinition = {
  id: 'xls',
  extensions: ['.xls'],
  maxFileSize: MAX_FILE_SIZE.xls,
  webview: 'spreadsheet',
  check(bytes) {
    // Many systems export web pages or SpreadsheetML XML under .xls; Excel opens them, but they are not workbooks.
    if (!isCompoundFile(bytes) && !isZip(bytes) && startsWithMarkup(bytes)) {
      throw new Error('This .xls file is a web page or XML spreadsheet with an .xls extension; Document Lens opens Excel 97-2003 workbooks');
    }
    checkCompoundFile(bytes, '.xls');
  },
};

/** First non-blank character `<`, after an optional UTF-8 byte order mark. */
function startsWithMarkup(bytes: Uint8Array): boolean {
  let i = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (i < Math.min(bytes.length, 1024) && (bytes[i] === 0x20 || bytes[i] === 0x09 || bytes[i] === 0x0a || bytes[i] === 0x0d)) {
    i++;
  }
  return bytes[i] === 0x3c;
}
