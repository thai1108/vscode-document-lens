import { writeZip } from './zipWriter.ts';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
export const OFFICE_DOCUMENT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

export interface OoxmlPart {
  name: string;
  contentType?: string;
  xml: string;
}

/** An OOXML package: content types, the package relationship to `main`, and the given parts. */
export function ooxmlPackage(main: OoxmlPart, others: OoxmlPart[] = []): Buffer {
  const parts = [main, ...others];
  const overrides = parts
    .filter((part) => part.contentType)
    .map((part) => `<Override PartName="/${part.name}" ContentType="${part.contentType}"/>`)
    .join('');
  return writeZip([
    {
      name: '[Content_Types].xml',
      data:
        `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        `<Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
    },
    {
      name: '_rels/.rels',
      data:
        `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="${main.name}"/></Relationships>`,
    },
    ...parts.map((part) => ({ name: part.name, data: XML_HEADER + part.xml })),
  ]);
}
