import type { Sample } from '../support/samples.ts';
import { writeWordDocument, type WordDocumentInput } from '../support/wordWriter.ts';

/**
 * Paragraphs in cp1252 (compressed) and UTF-16 (Japanese) pieces, a HYPERLINK field, a footnote, a table with a
 * two-paragraph cell and a page break. Each item is one piece of the piece table.
 */
export const SAMPLE_DOC: WordDocumentInput = {
  body: [
    'Document Lens sample\r',
    'This synthetic Word 97–2003 document has paragraphs, a table, a field and a footnote\x02.\r',
    'Prices: 5 € – “quoted” text.\r',
    'Visit \x13 HYPERLINK "https://example.com/" \x14the example site\x15 for more.\r',
    { row: ['Item', 'Qty', 'Remarks'] },
    { row: ['Apples', '3', 'Fresh'] },
    { row: ['Pears', '5', 'Two\rparagraphs'] },
    'After the table.\r',
    '\x0c',
    { text: '日本語のテキストも表示されます。\r', unicode: true },
    'Second page, back to cp1252.\r',
  ],
  footnotes: ['A footnote at the end of the text.'],
};

export const samples: Sample[] = [{ fileName: 'sample.doc', build: () => writeWordDocument(SAMPLE_DOC) }];
