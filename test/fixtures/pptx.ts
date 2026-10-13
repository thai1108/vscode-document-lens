import { ooxmlPackage } from '../support/ooxml.ts';
import type { Sample } from '../support/samples.ts';

// A package skeleton without slides: enough for the host's preflight, not for a slide renderer.
export const samples: Sample[] = [
  {
    fileName: 'sample.pptx',
    build: () =>
      ooxmlPackage({
        name: 'ppt/presentation.xml',
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
        xml: '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
      }),
  },
];
