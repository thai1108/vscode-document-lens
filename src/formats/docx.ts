import { MAX_FILE_SIZE } from '../limits.ts';
import { checkOoxmlPackage } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const docx: FormatDefinition = {
  id: 'docx',
  extensions: ['.docx'],
  maxFileSize: MAX_FILE_SIZE.docx,
  webview: 'docx',
  // The pages are DOM text, so VS Code's find widget can search them. No CSP extras: pictures and embedded fonts are
  // blob: URLs, and the document's style sheets are <style nonce> elements.
  findWidget: true,
  check: (bytes) => checkOoxmlPackage(bytes, '.docx'),
};
