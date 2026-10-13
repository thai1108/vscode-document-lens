// The OPC package: parts by name, their relationships, and parsed XML. Shared parts (masters, layouts, themes) are
// parsed once; slides are parsed each time they are drawn so a long deck does not keep every slide's DOM alive.
import { MAX_XML_PART_SIZE, formatMegabytes } from '../../limits.ts';
import { attr, children, decodeXml, parseXml } from './xml.ts';
import { ZipReader } from './zip.ts';

export interface Relationship {
  readonly id: string;
  /** Last segment of the relationship type URI, e.g. "slide", "image", "hyperlink". */
  readonly type: string;
  /** Part name inside the package, or the raw URL for external targets. */
  readonly target: string;
  readonly external: boolean;
}

export type Relationships = ReadonlyMap<string, Relationship>;

const NO_RELATIONSHIPS: Relationships = new Map();

export class PptxPackage {
  private readonly zip: ZipReader;
  private readonly parsed = new Map<string, Promise<Element>>();
  private readonly relationships = new Map<string, Promise<Relationships>>();

  private constructor(zip: ZipReader) {
    this.zip = zip;
  }

  static open(bytes: Uint8Array): PptxPackage {
    return new PptxPackage(ZipReader.open(bytes));
  }

  has(name: string): boolean {
    return this.zip.has(name);
  }

  size(name: string): number | undefined {
    return this.zip.size(name);
  }

  bytes(name: string): Promise<Uint8Array> {
    return this.zip.read(name);
  }

  async text(name: string): Promise<string> {
    const size = this.zip.size(name);
    if (size !== undefined && size > MAX_XML_PART_SIZE) {
      throw new Error(`${name} is too large to parse (limit ${formatMegabytes(MAX_XML_PART_SIZE)})`);
    }
    return decodeXml(await this.zip.read(name));
  }

  /** The part's root element; `shared` parts are parsed once and kept. */
  xml(name: string, shared = false): Promise<Element> {
    const cached = this.parsed.get(name);
    if (cached) {
      return cached;
    }
    const parsing = this.text(name).then((text) => parseXml(text, name));
    if (shared) {
      this.parsed.set(name, parsing);
    }
    return parsing;
  }

  rels(partName: string): Promise<Relationships> {
    let loading = this.relationships.get(partName);
    if (!loading) {
      loading = this.loadRels(partName);
      this.relationships.set(partName, loading);
    }
    return loading;
  }

  private async loadRels(partName: string): Promise<Relationships> {
    const name = relsPartName(partName);
    if (!this.zip.has(name)) {
      return NO_RELATIONSHIPS;
    }
    const root = parseXml(await this.text(name), name);
    const result = new Map<string, Relationship>();
    for (const element of children(root, 'Relationship')) {
      const id = attr(element, 'Id');
      const type = attr(element, 'Type') ?? '';
      const target = attr(element, 'Target');
      if (!id || target === undefined) {
        continue;
      }
      const external = attr(element, 'TargetMode') === 'External';
      result.set(id, {
        id,
        type: type.slice(type.lastIndexOf('/') + 1),
        target: external ? target : resolvePartName(partName, target),
        external,
      });
    }
    return result;
  }
}

/** ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels */
export function relsPartName(partName: string): string {
  const slash = partName.lastIndexOf('/');
  return `${partName.slice(0, slash + 1)}_rels/${partName.slice(slash + 1)}.rels`;
}

/** Resolves a relationship target against the part that holds the relationship ("../media/a.png" → "ppt/media/a.png"). */
export function resolvePartName(source: string, target: string): string {
  const path = target.split(/[?#]/, 1)[0] ?? '';
  const segments = path.startsWith('/') ? [] : source.split('/').slice(0, -1);
  for (const segment of path.split('/')) {
    if (segment === '..') {
      segments.pop();
    } else if (segment !== '.' && segment !== '') {
      segments.push(safeDecode(segment));
    }
  }
  return segments.join('/');
}

function safeDecode(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

export function firstOfType(rels: Relationships, type: string): Relationship | undefined {
  for (const rel of rels.values()) {
    if (rel.type === type && !rel.external) {
      return rel;
    }
  }
  return undefined;
}
