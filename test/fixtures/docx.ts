import { ooxmlPackage } from '../support/ooxml.ts';
import type { Sample } from '../support/samples.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export const samples: Sample[] = [
  {
    fileName: 'sample.docx',
    build: () =>
      ooxmlPackage({
        name: 'word/document.xml',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        xml: `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Document Lens sample</w:t></w:r></w:p></w:body></w:document>`,
      }),
  },
];
