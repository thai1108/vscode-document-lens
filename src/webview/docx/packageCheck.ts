// Checks a .docx package in the webview before docx-preview unpacks it:
// - the archive reads the same way here and in JSZip (zipDirectory.ts, plus JSZip's own list of entries);
// - every part docx-preview parses into an XML DOM is at most MAX_XML_PART_SIZE, the parts stay within a total size
//   and element budget, and each is well-formed. A malformed part would otherwise reach DOMParser, whose error report
//   carries inline styles the CSP refuses (console errors);
// - no entry inflates to anything but its declared size, so the host's ZIP checks (which rely on those sizes) hold.
import JSZip from 'jszip/dist/jszip.min.js';
import { formatMegabytes, MAX_XML_PART_SIZE } from '../../limits.ts';
import { DOCX_LIMITS } from './limits.ts';
import { ForgedSizeError, inflateEntry, readZipDirectory, readZipText, ZipFormatError, type ZipEntry } from './zipDirectory.ts';

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

export interface PackageCheckOptions {
  maxXmlPartSize?: number;
  maxXmlTotalSize?: number;
  maxXmlElements?: number;
  /** Parses one XML part, throwing when it is not well-formed; XML fragment parsing (parseXmlPart) by default. */
  parseXml?: (name: string, text: string) => ParentNode;
}

export interface CheckedPackage {
  /** Every entry's verified size, by the name docx-preview looks it up by. */
  readonly sizes: ReadonlyMap<string, number>;
}

export async function checkPackage(data: Uint8Array, options: PackageCheckOptions = {}): Promise<CheckedPackage> {
  const { maxXmlPartSize = MAX_XML_PART_SIZE, maxXmlTotalSize = DOCX_LIMITS.maxXmlTotalSize, maxXmlElements = DOCX_LIMITS.maxXmlElements } = options;
  const parseXml = options.parseXml ?? parseXmlPart;
  const entries = await zipped(() => readZipDirectory(data));
  await matchJsZip(data, entries);

  const byName = new Map(entries.map((entry) => [entry.name, entry]));
  // OpenXmlPackage.get drops a leading slash (it also tries backslash separators, which readZipDirectory refuses).
  const lookup = (path: string) => byName.get(path.startsWith('/') ? path.slice(1) : path);
  const visited = new Set<ZipEntry>();
  let xmlBytes = 0;
  let xmlElements = 0;
  /** Reads and checks one XML part; undefined when it is missing (docx-preview skips it too) or already read. */
  const readPart = async (path: string): Promise<ParentNode | undefined> => {
    const entry = lookup(path);
    if (!entry || visited.has(entry)) {
      return undefined;
    }
    visited.add(entry);
    if (entry.size > maxXmlPartSize) {
      throw new Error(`its part ${entry.name} is larger than ${formatMegabytes(maxXmlPartSize)}, the most Document Lens reads from one part`);
    }
    xmlBytes += entry.size;
    if (xmlBytes > maxXmlTotalSize) {
      throw new Error(`its XML parts add up to more than ${formatMegabytes(maxXmlTotalSize)}, the most Document Lens reads`);
    }
    const text = await zipped(() => readZipText(data, entry));
    xmlElements += countStartTags(text);
    if (xmlElements > maxXmlElements) {
      throw new Error(`it has more than ${maxXmlElements.toLocaleString('en-US')} XML elements, the most Document Lens reads`);
    }
    return parseXml(entry.name, text);
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

  // Everything else docx-preview may inflate (pictures, fonts) must have its declared size too.
  for (const entry of entries) {
    if (!visited.has(entry)) {
      await zipped(() => inflateEntry(data, entry));
    }
  }
  return { sizes: new Map(entries.map((entry) => [entry.name, entry.size])) };
}

/** Runs a ZIP reading step, turning its format errors into messages for the error banner. */
async function zipped<T>(step: () => T | Promise<T>): Promise<T> {
  try {
    return await step();
  } catch (error) {
    if (error instanceof ForgedSizeError) {
      throw new Error(`the file is damaged: ${error.message}`, { cause: error });
    }
    if (error instanceof ZipFormatError) {
      throw new Error(`the file is damaged or uses ZIP features Document Lens does not read (${error.message})`, { cause: error });
    }
    throw error;
  }
}

/** JSZip (which docx-preview reads the file with) must list exactly the entries this reader checks. */
async function matchJsZip(data: Uint8Array, entries: readonly ZipEntry[]): Promise<void> {
  let names: string[];
  try {
    names = Object.keys((await JSZip.loadAsync(data)).files);
  } catch (error) {
    throw new Error(`the file is damaged (${(error as Error).message})`, { cause: error });
  }
  const ours = new Set(entries.map((entry) => entry.name));
  const unexpected = names.find((name) => !ours.has(name));
  if (unexpected !== undefined || names.length !== ours.size) {
    throw new Error(`the file is damaged or uses ZIP features Document Lens does not read (its entries read differently: ${JSON.stringify(unexpected ?? '')})`);
  }
}

/** Start tags in an XML text (an upper bound: markup inside comments and CDATA counts too). */
export function countStartTags(text: string): number {
  let count = 0;
  for (let index = text.indexOf('<'); index >= 0; index = text.indexOf('<', index + 1)) {
    const next = text.charCodeAt(index + 1);
    // Not </, <! or <?
    if (next !== 47 && next !== 33 && next !== 63) {
      count += 1;
    }
  }
  return count;
}

/**
 * Parses a part as an XML fragment, which (unlike DOMParser) fails with an exception instead of building an error
 * document styled with inline styles. The text is first trimmed the way docx-preview trims it (trimXmlDeclaration),
 * plus a leading XML declaration that spans lines, which DOMParser accepts but a fragment may not contain.
 */
export function parseXmlPart(name: string, text: string): Element {
  const source = text
    .replace(/<[?].*[?]>/, '')
    .replace(/^\uFEFF/, '')
    .replace(/^\s*<\?xml[\s\S]*?\?>/, '');
  const container = document.implementation.createDocument(null, null, null).createElementNS(null, 'part');
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
