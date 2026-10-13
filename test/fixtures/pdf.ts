// Synthetic PDFs, written object by object. sample.pdf exercises what the viewer must handle: pages of different
// sizes and a rotated page, text to select and find (Latin and Japanese), external and internal links, an outline,
// notes, and the data pdf.js loads on demand: CMaps (Japanese text in a non-embedded font), standard font data
// (ZapfDingbats) and the wasm image decoders (a CCITT fax image).
import { pdfStream, pdfString, PdfObjects, textPdf, writePdf } from '../support/pdfWriter.ts';
import type { Sample } from '../support/samples.ts';

/** What sample.pdf and the other samples contain, for the tests that open them. */
export const SAMPLE_PDF = {
  pageCount: 4,
  title: 'Document Lens sample',
  findTerm: 'needle',
  /** Occurrences of findTerm: two on page 1, one on page 2 and one on page 4. */
  findMatches: 4,
  japanese: '日本語のテキスト検索',
  externalUrl: 'https://example.com/document-lens',
  outlineUrl: 'https://example.com/outline',
  outlineTitles: ['Introduction', 'Images and symbols', 'Landscape page', 'Rotated page', 'Project website'],
  /** [width, height] in points before /Rotate; page 4 is rotated by 90°. */
  pageSizes: [
    [612, 792],
    [595, 842],
    [842, 595],
    [612, 792],
  ],
  /** The CCITT fax image: a square on page 2, left half white and right half black (PDF coordinates). */
  image: { page: 2, x: 72, y: 400, size: 120 },
  longPageCount: 600,
};

function text(font: string, size: number, x: number, y: number, string: string): string {
  return `BT /${font} ${size} Tf ${x} ${y} Td ${string} Tj ET`;
}

/** UTF-16BE hex string, the code space of the UniJIS-UCS2-H CMap. */
function ucs2(string: string): string {
  return `<${[...string].map((character) => (character.codePointAt(0) ?? 0).toString(16).padStart(4, '0')).join('')}>`;
}

/**
 * CCITT group 3 one-dimensional data (K = 0) for 16 × `rows` pixels, each row a white run of 8 followed by a black
 * run of 8 (T.4 terminating codes 10011 and 000101), padded to whole bytes.
 */
function ccittHalfBlack(rows: number): string {
  const bits = '10011000101'.repeat(rows);
  const bytes = bits.padEnd(Math.ceil(bits.length / 8) * 8, '0').match(/.{8}/g) ?? [];
  return bytes.map((byte) => String.fromCharCode(parseInt(byte, 2))).join('');
}

function link(rect: number[], target: string): string {
  return `<< /Type /Annot /Subtype /Link /Rect [${rect.join(' ')}] /Border [0 0 0] ${target} >>`;
}

function samplePdf(): Buffer {
  const pdf = new PdfObjects();
  const catalog = pdf.add();
  const pages = pdf.add();
  const helvetica = pdf.add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
  const dingbats = pdf.add('<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>');
  const gothicDescriptor = pdf.add(
    '<< /Type /FontDescriptor /FontName /HeiseiKakuGo-W5 /Flags 4 /FontBBox [-92 -250 1010 922] /ItalicAngle 0 /Ascent 752 /Descent -221 /CapHeight 737 /StemV 114 >>',
  );
  const gothicCid = pdf.add(
    `<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiKakuGo-W5 /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 2 >> /FontDescriptor ${gothicDescriptor} 0 R /DW 1000 >>`,
  );
  const gothic = pdf.add(`<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiKakuGo-W5 /Encoding /UniJIS-UCS2-H /DescendantFonts [${gothicCid} 0 R] >>`);
  const image = pdf.add(
    pdfStream(
      '/Type /XObject /Subtype /Image /Width 16 /Height 16 /ColorSpace /DeviceGray /BitsPerComponent 1 /Filter /CCITTFaxDecode /DecodeParms << /K 0 /Columns 16 /Rows 16 /EndOfBlock false >>',
      ccittHalfBlack(16),
    ),
  );
  const resources = `<< /Font << /F1 ${helvetica} 0 R /F2 ${gothic} 0 R /F3 ${dingbats} 0 R >> /XObject << /Im1 ${image} 0 R >> >>`;
  const [page1, page2, page3, page4] = [pdf.add(), pdf.add(), pdf.add(), pdf.add()] as [number, number, number, number];
  const page = (id: number, size: number[], content: string, extra = '') =>
    pdf.set(id, `<< /Type /Page /Parent ${pages} 0 R /MediaBox [0 0 ${size.join(' ')}] /Resources ${resources} /Contents ${pdf.add(pdfStream('', content))} 0 R ${extra}>>`);
  const [size1, size2, size3, size4] = SAMPLE_PDF.pageSizes as [number[], number[], number[], number[]];

  const annotations = [
    pdf.add(link([70, 615, 250, 634], `/A << /S /URI /URI ${pdfString(SAMPLE_PDF.externalUrl)} >>`)),
    pdf.add(link([70, 591, 250, 610], '/Dest /landscape')),
    pdf.add(`<< /Type /Annot /Subtype /Text /Rect [520 700 542 722] /Name /Comment /T (Reviewer) /Contents (A note on the first page) >>`),
    // The icon name comes from the file; it must never turn into a request.
    pdf.add(`<< /Type /Annot /Subtype /Text /Rect [520 660 542 682] /Name /#2E#2E#2Fsecret /Contents (A note with an odd icon name) >>`),
  ];
  page(
    page1,
    size1,
    [
      text('F1', 22, 72, 720, pdfString(SAMPLE_PDF.title)),
      text('F1', 14, 72, 680, pdfString('Page one: the first needle is here.')),
      text('F1', 14, 72, 656, pdfString('A second needle on the same page.')),
      text('F1', 14, 72, 620, pdfString('Open the project website')),
      text('F1', 14, 72, 596, pdfString('Go to the landscape page')),
      text('F2', 16, 72, 560, ucs2(SAMPLE_PDF.japanese)),
    ].join('\n'),
    `/Annots [${annotations.map((id) => `${id} 0 R`).join(' ')}] `,
  );
  page(
    page2,
    size2,
    [
      text('F1', 14, 72, 770, pdfString('Page two (A4) with a needle and an image.')),
      text('F3', 18, 72, 740, pdfString('4 4 4')),
      `q ${SAMPLE_PDF.image.size} 0 0 ${SAMPLE_PDF.image.size} ${SAMPLE_PDF.image.x} ${SAMPLE_PDF.image.y} cm /Im1 Do Q`,
    ].join('\n'),
  );
  page(
    page3,
    size3,
    [text('F1', 18, 72, 520, pdfString('Landscape page')), text('F1', 14, 72, 490, pdfString('Go to the rotated page'))].join('\n'),
    `/Annots [${pdf.add(link([70, 485, 240, 504], `/Dest [${page4} 0 R /Fit]`))} 0 R] `,
  );
  page(page4, size4, text('F1', 14, 72, 720, pdfString('Rotated page with the last needle.')), '/Rotate 90 ');
  pdf.set(pages, `<< /Type /Pages /Kids [${[page1, page2, page3, page4].map((id) => `${id} 0 R`).join(' ')}] /Count 4 >>`);

  const outline = pdf.add();
  const [intro, images, landscape, rotated, website] = [pdf.add(), pdf.add(), pdf.add(), pdf.add(), pdf.add()];
  const [introTitle, imagesTitle, landscapeTitle, rotatedTitle, websiteTitle] = SAMPLE_PDF.outlineTitles.map(pdfString);
  pdf.set(outline, `<< /Type /Outlines /First ${intro} 0 R /Last ${website} 0 R /Count 5 >>`);
  pdf.set(intro, `<< /Title ${introTitle} /Parent ${outline} 0 R /Next ${images} 0 R /Dest [${page1} 0 R /XYZ 0 792 null] >>`);
  pdf.set(images, `<< /Title ${imagesTitle} /Parent ${outline} 0 R /Prev ${intro} 0 R /Next ${landscape} 0 R /Dest [${page2} 0 R /Fit] >>`);
  pdf.set(
    landscape,
    `<< /Title ${landscapeTitle} /Parent ${outline} 0 R /Prev ${images} 0 R /Next ${website} 0 R /First ${rotated} 0 R /Last ${rotated} 0 R /Count 1 /Dest /landscape >>`,
  );
  pdf.set(rotated, `<< /Title ${rotatedTitle} /Parent ${landscape} 0 R /Dest [${page4} 0 R /Fit] >>`);
  pdf.set(website, `<< /Title ${websiteTitle} /Parent ${outline} 0 R /Prev ${landscape} 0 R /A << /S /URI /URI ${pdfString(SAMPLE_PDF.outlineUrl)} >> >>`);

  pdf.set(catalog, `<< /Type /Catalog /Pages ${pages} 0 R /Outlines ${outline} 0 R /Dests << /landscape [${page3} 0 R /XYZ 0 595 null] >> >>`);
  return pdf.write();
}

/** Standard security handler with a user password: the stored check value matches no password, so none opens it. */
function encryptedPdf(): Buffer {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>',
    `<< /Filter /Standard /V 1 /R 2 /Length 40 /P -44 /O <${'aa'.repeat(32)}> /U <${'55'.repeat(32)}> >>`,
  ];
  const id = `<${'01'.repeat(16)}>`;
  return writePdf(objects, { trailer: `/Encrypt 4 0 R /ID [${id} ${id}]` });
}

export const samples: Sample[] = [
  { fileName: 'sample.pdf', build: samplePdf },
  {
    fileName: 'long.pdf',
    build: () => textPdf(Array.from({ length: SAMPLE_PDF.longPageCount }, (_, i) => [`Long document page ${i + 1}`])),
  },
  { fileName: 'encrypted.pdf', build: encryptedPdf },
  // Passes the host's header check but has no PDF structure behind it.
  { fileName: 'broken.pdf', build: () => Buffer.from('%PDF-1.7\nThis is not a PDF body.\n', 'latin1') },
];
