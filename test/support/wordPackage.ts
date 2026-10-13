// Builds .docx packages for tests: word/document.xml with its relationships, further XML parts (styles, numbering,
// headers, ...) and binary media. ooxmlPackage (ooxml.ts) covers single-part packages; Word documents need part
// relationships and pictures.
import { OFFICE_DOCUMENT_RELATIONSHIP } from './ooxml.ts';
import { writeZip, type ZipInput } from './zipWriter.ts';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const RELATIONSHIP_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CONTENT_TYPE_BASE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';

/** Namespace declarations for the root element of every WordprocessingML part. */
export const WORD_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
].join(' ');

export const RELATIONSHIP = {
  styles: `${RELATIONSHIP_BASE}styles`,
  numbering: `${RELATIONSHIP_BASE}numbering`,
  header: `${RELATIONSHIP_BASE}header`,
  footer: `${RELATIONSHIP_BASE}footer`,
  footnotes: `${RELATIONSHIP_BASE}footnotes`,
  hyperlink: `${RELATIONSHIP_BASE}hyperlink`,
  image: `${RELATIONSHIP_BASE}image`,
  fontTable: `${RELATIONSHIP_BASE}fontTable`,
  font: `${RELATIONSHIP_BASE}font`,
  altChunk: `${RELATIONSHIP_BASE}aFChunk`,
} as const;

export const CONTENT_TYPE = {
  document: `${CONTENT_TYPE_BASE}document.main+xml`,
  styles: `${CONTENT_TYPE_BASE}styles+xml`,
  numbering: `${CONTENT_TYPE_BASE}numbering+xml`,
  header: `${CONTENT_TYPE_BASE}header+xml`,
  footer: `${CONTENT_TYPE_BASE}footer+xml`,
  footnotes: `${CONTENT_TYPE_BASE}footnotes+xml`,
  fontTable: `${CONTENT_TYPE_BASE}fontTable+xml`,
  obfuscatedFont: 'application/vnd.openxmlformats-officedocument.obfuscatedFont',
} as const;

export interface WordRelationship {
  id: string;
  type: string;
  /** Relative to the folder of the part it belongs to, or an absolute URL when `external`. */
  target: string;
  external?: boolean;
}

export interface WordPart {
  /** Path in the package, e.g. word/styles.xml. */
  name: string;
  contentType: string;
  /** The part's XML without the XML declaration. */
  xml: string;
  /** Adds a relationship from word/document.xml to this part. */
  relationship?: { id: string; type: string };
  /** The part's own relationships (e.g. the font table's fonts). */
  relationships?: WordRelationship[];
}

export interface BinaryPart {
  name: string;
  contentType: string;
  data: Uint8Array;
  relationship?: { id: string; type: string };
}

export interface WordPackage {
  /** Contents of <w:body>. */
  body: string;
  /** Relationships of word/document.xml besides those its parts declare (links, external pictures, ...). */
  relationships?: WordRelationship[];
  parts?: WordPart[];
  media?: BinaryPart[];
  /** Raw ZIP entries added as they are, for broken or hostile packages. */
  extraEntries?: ZipInput[];
}

export function wordPackage(word: WordPackage): Buffer {
  const parts = word.parts ?? [];
  const media = word.media ?? [];
  const relationships: WordRelationship[] = [
    ...[...parts, ...media].flatMap((part) => (part.relationship ? [{ ...part.relationship, target: part.name.replace(/^word\//, '') }] : [])),
    ...(word.relationships ?? []),
  ];
  const overrides = [{ name: 'word/document.xml', contentType: CONTENT_TYPE.document }, ...parts, ...media]
    .map((part) => `<Override PartName="/${escapeXml(part.name)}" ContentType="${escapeXml(part.contentType)}"/>`)
    .join('');
  return writeZip([
    {
      name: '[Content_Types].xml',
      data:
        `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        `<Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
    },
    { name: '_rels/.rels', data: relationshipsXml([{ id: 'rId1', type: OFFICE_DOCUMENT_RELATIONSHIP, target: 'word/document.xml' }]) },
    { name: 'word/document.xml', data: `${XML_HEADER}<w:document ${WORD_NAMESPACES}><w:body>${word.body}</w:body></w:document>` },
    { name: 'word/_rels/document.xml.rels', data: relationshipsXml(relationships) },
    ...parts.flatMap((part) => [
      { name: part.name, data: XML_HEADER + part.xml },
      ...(part.relationships ? [{ name: relationshipsName(part.name), data: relationshipsXml(part.relationships) }] : []),
    ]),
    ...media.map((part) => ({ name: part.name, data: Buffer.from(part.data) })),
    ...(word.extraEntries ?? []),
  ]);
}

/** word/header1.xml → word/_rels/header1.xml.rels */
function relationshipsName(partName: string): string {
  const slash = partName.lastIndexOf('/') + 1;
  return `${partName.slice(0, slash)}_rels/${partName.slice(slash)}.rels`;
}

export function relationshipsXml(relationships: readonly WordRelationship[]): string {
  const items = relationships
    .map(
      ({ id, type, target, external }) =>
        `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(type)}" Target="${escapeXml(target)}"${external ? ' TargetMode="External"' : ''}/>`,
    )
    .join('');
  return `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}

export function escapeXml(text: string): string {
  return text.replace(/[&<>"']/g, (char) => `&#${char.charCodeAt(0)};`);
}

/** A paragraph: `properties` is the inner XML of <w:pPr>, `content` the runs and other inline content. */
export function paragraph(content: string, properties = ''): string {
  return `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}${content}</w:p>`;
}

/** A run of text: `properties` is the inner XML of <w:rPr>. */
export function run(text: string, properties = ''): string {
  return `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

/** An inline picture of `width` × `height` EMU (914400 per inch) showing the image relationship `relationshipId`. */
export function inlinePicture(relationshipId: string, width: number, height: number, name = 'Picture'): string {
  const size = `cx="${width}" cy="${height}"`;
  return (
    `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent ${size}/>` +
    `<wp:docPr id="1" name="${escapeXml(name)}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="${escapeXml(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${escapeXml(relationshipId)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${size}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  );
}
