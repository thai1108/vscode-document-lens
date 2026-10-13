import { MAX_FILE_SIZE } from '../limits.ts';
import { checkOoxmlPackage } from '../preflight.ts';
import type { FormatDefinition } from './types.ts';

export const pptx: FormatDefinition = {
  id: 'pptx',
  extensions: ['.pptx'],
  maxFileSize: MAX_FILE_SIZE.pptx,
  webview: 'pptx',
  // Pictures are blob: URLs, which the base policy allows; the viewer needs nothing more.
  // Only the current slide's text is in the page, so the find widget would miss every other slide.
  findWidget: false,
  check: (bytes) => checkOoxmlPackage(bytes, '.pptx'),
};
