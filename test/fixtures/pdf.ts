import type { Sample } from '../support/samples.ts';
import { textPdf } from '../support/pdfWriter.ts';

export const samples: Sample[] = [
  {
    fileName: 'sample.pdf',
    build: () => textPdf([['Document Lens sample', 'Page one'], ['Page two']]),
  },
];
