// Checks a .docx package in the webview before docx-preview unpacks it:
// - every part docx-preview parses into an XML DOM is at most MAX_XML_PART_SIZE and well-formed. A malformed part
//   would otherwise reach DOMParser, whose error report carries inline styles the CSP refuses (console errors);
// - no entry inflates past its declared size, so the host's ZIP checks (which rely on those sizes) hold.
import { formatMegabytes, MAX_XML_PART_SIZE } from '../../limits.ts';
import { ForgedSizeError, inflateEntry, readZipDirectory, readZipText, type ZipEntry } from './zipDirectory.ts';

const PACKAGE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/';
const OFFICE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/** The package-level parts docx-preview loads, with the path it falls back to (WordDocument.load). */
const TOP_LEVEL_PARTS: readonly (readonly [type: string, fallback: string])[] = [
  [`${OFFICE_RELATIONSHIPS}officeDocument`, 'word/document.xml'],
  [`${OFFICE_RELATIONSHIPS}extended-properties`, 'docProps/app.xml'],
  [`${PACKAGE_RELATIONSHIPS}core-properties`, 'docProps/core.xml'],
  [`${PACKAGE_RELATIONSHIPS}custom-properties`, 'docProps/custom.xml'],
];

/** Relationship types whose target docx-preview parses as XML (WordDocument.loadRelationshipPart). */
const XML_PART_TYPES = new Set([
  ...TOP_LEVEL_PARTS.map(([type]) => type),
  ...['fontTable', 'numbering', 'styles', 'theme', 'settings', 'footnotes', 'endnotes', 'footer', 'header', 'comments'].map(
    (name) => OFFICE_RELATIONSHIPS + name,
  ),
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
]);

interface Relationship {
  type: string | null;
  target: string | null;
}

export async function checkPackage(data: Uint8Array, maxXmlPartSize = MAX_XML_PART_SIZE): Promise<void> {
  const entries = readZipDirectory(data);
  const tooLarge = (name: string) =>
    new Error(`its part ${name} is larger than ${formatMegabytes(maxXmlPartSize)}, the most Document Lens reads from one part`);
  // JSZip may name these differently from this reader, so a large one could be an XML part under another name.
  const ambiguous = entries.find((entry) => entry.ambiguousName && entry.size > maxXmlPartSize);
  if (ambiguous) {
    throw tooLarge(ambiguous.name);
  }

  // Of entries sharing a name, JSZip keeps the last one.
  const byName = new Map(entries.map((entry) => [entry.name, entry]));
  // OpenXmlPackage.get: a leading slash is dropped, and backslash separators are accepted.
  const lookup = (path: string) => {
    const name = path.startsWith('/') ? path.slice(1) : path;
    return byName.get(name) ?? byName.get(name.replaceAll('/', '\\'));
  };
  const xml = document.implementation.createDocument(null, null, null);
  const visited = new Set<ZipEntry>();
  /**
   * Reads and checks one XML part: undefined when it is missing (docx-preview skips it too) or already read, null when
   * it cannot be inflated here (an unusual deflate stream is left to JSZip to accept or reject).
   */
  const readPart = async (path: string): Promise<Element | null | undefined> => {
    const entry = lookup(path);
    if (!entry || visited.has(entry)) {
      return undefined;
    }
    visited.add(entry);
    if (entry.size > maxXmlPartSize) {
      throw tooLarge(entry.name);
    }
    let text: string;
    try {
      text = await readZipText(data, entry);
    } catch (error) {
      if (error instanceof ForgedSizeError) {
        throw forged(entry, error);
      }
      return null;
    }
    return parseXmlPart(xml, entry.name, text);
  };
  const relationships = async (path: string): Promise<Relationship[]> => {
    const root = (await readPart(relationshipsPath(path)))?.firstElementChild;
    return root ? [...root.children].map((element) => ({ type: attribute(element, 'Type'), target: attribute(element, 'Target') })) : [];
  };

  await readPart('[Content_Types].xml');
  const packageRelationships = await relationships('');
  const queue = TOP_LEVEL_PARTS.map(([type, fallback]) => packageRelationships.find((relationship) => relationship.type === type)?.target ?? fallback);
  for (let path = queue.shift(); path !== undefined; path = queue.shift()) {
    if ((await readPart(path)) === undefined) {
      continue;
    }
    const folder = path.slice(0, path.lastIndexOf('/') + 1);
    for (const relationship of await relationships(path)) {
      if (relationship.type && relationship.target && XML_PART_TYPES.has(relationship.type)) {
        queue.push(resolvePath(relationship.target, folder));
      }
    }
  }

  // Everything else docx-preview may inflate (pictures, fonts) must not grow past its declared size either.
  for (const entry of entries) {
    if (visited.has(entry)) {
      continue;
    }
    try {
      await inflateEntry(data, entry);
    } catch (error) {
      if (error instanceof ForgedSizeError) {
        throw forged(entry, error);
      }
    }
  }
}

function forged(entry: ZipEntry, cause: Error): Error {
  return new Error(`the file is damaged: ${entry.name} is larger than its ZIP header says`, { cause });
}

/**
 * Parses a part as an XML fragment, which (unlike DOMParser) fails with an exception instead of a styled error
 * document. The text is first trimmed the way docx-preview trims it (trimXmlDeclaration), plus a leading XML
 * declaration that spans lines, which DOMParser accepts but a fragment may not contain.
 */
function parseXmlPart(owner: XMLDocument, name: string, text: string): Element {
  const source = text
    .replace(/<[?].*[?]>/, '')
    .replace(/^\uFEFF/, '')
    .replace(/^\s*<\?xml[\s\S]*?\?>/, '');
  const container = owner.createElementNS(null, 'part');
  const damaged = (cause?: unknown) => new Error(`the document is damaged (${name} is not well-formed XML)`, { cause });
  try {
    container.innerHTML = source;
  } catch (error) {
    throw damaged(error);
  }
  // A fragment may hold several elements or loose text; a document has exactly one root element.
  const looseText = [...container.childNodes].some((node) => node.nodeType === 3 && node.textContent?.trim());
  if (container.childElementCount !== 1 || looseText) {
    throw damaged();
  }
  return container;
}

/** Like docx-preview's XmlParser.attr: the attribute is matched by local name. */
function attribute(element: Element, localName: string): string | null {
  return [...element.attributes].find((candidate) => candidate.localName === localName)?.value ?? null;
}

/** The relationships part of `path` ('' for the package), as docx-preview names it. */
function relationshipsPath(path: string): string {
  if (!path) {
    return '_rels/.rels';
  }
  const slash = path.lastIndexOf('/') + 1;
  return `${path.slice(0, slash)}_rels/${path.slice(slash)}.rels`;
}

/** docx-preview's resolvePath: relative targets resolve against the folder of the part that refers to them. */
function resolvePath(path: string, base: string): string {
  const prefix = 'http://docx/';
  try {
    return new URL(path, prefix + base).toString().substring(prefix.length);
  } catch {
    return `${base}${path}`;
  }
}
