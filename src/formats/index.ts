import { doc } from './doc.ts';
import { docx } from './docx.ts';
import { pdf } from './pdf.ts';
import { ppt } from './ppt.ts';
import { pptx } from './pptx.ts';
import type { FormatDefinition } from './types.ts';
import { xls } from './xls.ts';
import { xlsx } from './xlsx.ts';

export const FORMATS: readonly FormatDefinition[] = [pdf, docx, doc, xlsx, xls, pptx, ppt];

export function formatForFile(fileName: string): FormatDefinition | undefined {
  const name = fileName.toLowerCase();
  return FORMATS.find((format) => format.extensions.some((extension) => name.endsWith(extension)));
}
