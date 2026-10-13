import { rgbPng } from '../support/png.ts';
import { writePpt, type PptSpec } from '../support/pptWriter.ts';
import type { Sample } from '../support/samples.ts';

/** A 120 × 90 colour gradient with deterministic noise, so the PNG is large enough to need regular sectors. */
export function samplePicture(): Buffer {
  let seed = 7;
  const noise = () => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return (seed >> 16) & 0x1f;
  };
  return rgbPng(120, 90, (x, y) => [Math.min(255, x * 2 + noise()), Math.min(255, y * 2 + noise()), 160 + noise()]);
}

/**
 * Three 16:9 slides: a title slide whose placeholder text sits in the slide list (as PowerPoint 97 saves it), a slide
 * with its text in the shapes (as later versions save it) with bullets, a text box, a footer and a picture, and a
 * hidden slide in 8-bit text.
 */
export const SAMPLE_PPT: PptSpec = {
  slideSize: { width: 5760, height: 3240 },
  pictures: [{ format: 'png', data: samplePicture() }],
  slides: [
    {
      placeholders: [
        { type: 'centerTitle', text: 'Quarterly review' },
        { type: 'subtitle', text: '四半期レビュー\rDocument Lens sample' },
      ],
    },
    {
      placeholders: [
        { type: 'title', text: 'Highlights', inShape: true },
        { type: 'body', text: 'Revenue grew 12%\rNew customers in Osaka\rRoadmap\vsecond line\r東京オフィス開設', levels: [0, 1, 0, 1], inShape: true },
      ],
      textBoxes: [{ type: 'other', text: '注: 速報値です' }],
      pictures: [1],
      footer: 'Confidential',
      hidden: false,
    },
    {
      hidden: true,
      placeholders: [
        { type: 'title', text: 'Backup slide', bytes: true },
        { type: 'body', text: 'Only for questions', bytes: true },
      ],
    },
  ],
};

export const samples: Sample[] = [{ fileName: 'sample.ppt', build: () => writePpt(SAMPLE_PPT) }];
