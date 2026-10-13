// XML parsing and DOM navigation for the package parts. Elements are matched by local name: PresentationML and
// DrawingML never give one name two meanings, and producers do not always use the usual prefixes.
import { formatMegabytes, MAX_XML_PART_SIZE } from '../../limits.ts';

export const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function decodeXml(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(bytes);
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(bytes);
  }
  return new TextDecoder().decode(bytes);
}

/**
 * Parses one part into its root element. Parts with a DOCTYPE are refused outright (OOXML never needs one, and entity
 * declarations are how XML bombs expand), as are malformed and oversized parts.
 */
export function parseXml(text: string, partName: string): Element {
  if (text.length > MAX_XML_PART_SIZE) {
    throw new Error(`${partName} is too large to parse (limit ${formatMegabytes(MAX_XML_PART_SIZE)})`);
  }
  if (/<!DOCTYPE/i.test(text)) {
    throw new Error(`${partName} declares a DOCTYPE, which presentations never need`);
  }
  const document = new DOMParser().parseFromString(text, 'application/xml');
  const root = document.documentElement;
  if (!root || document.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`${partName} is not well-formed XML`);
  }
  return root;
}

/**
 * Element children, with Markup Compatibility blocks resolved: the fallback branch is what producers write for
 * readers that do not know the extension (a picture of an equation, ink or a newer chart type).
 */
export function kids(parent: Element | undefined): Element[] {
  if (!parent) {
    return [];
  }
  const result: Element[] = [];
  for (const node of Array.from(parent.children)) {
    if (node.localName === 'AlternateContent') {
      const branches = Array.from(node.children);
      const branch = branches.find((candidate) => candidate.localName === 'Fallback') ?? branches.find((candidate) => candidate.localName === 'Choice');
      if (branch) {
        result.push(...Array.from(branch.children));
      }
    } else {
      result.push(node);
    }
  }
  return result;
}

export function child(parent: Element | undefined, name: string): Element | undefined {
  return kids(parent).find((node) => node.localName === name);
}

export function children(parent: Element | undefined, name: string): Element[] {
  return kids(parent).filter((node) => node.localName === name);
}

/** Follows a path of local names, e.g. find(sp, 'nvSpPr', 'nvPr', 'ph'). */
export function find(parent: Element | undefined, ...names: string[]): Element | undefined {
  let node = parent;
  for (const name of names) {
    node = child(node, name);
    if (!node) {
      return undefined;
    }
  }
  return node;
}

/** First child whose local name is one of `names`. */
export function oneOf(parent: Element | undefined, names: readonly string[]): Element | undefined {
  return kids(parent).find((node) => names.includes(node.localName));
}

export function attr(element: Element | undefined, name: string): string | undefined {
  return element?.getAttribute(name) ?? undefined;
}

/** A numeric attribute; percentages written as "50%" (strict OOXML) become 50000. */
export function num(element: Element | undefined, name: string): number | undefined {
  const value = attr(element, name)?.trim();
  if (value === undefined || value === '') {
    return undefined;
  }
  const parsed = value.endsWith('%') ? Number.parseFloat(value) * 1000 : Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

export function bool(element: Element | undefined, name: string): boolean | undefined {
  const value = attr(element, name);
  if (value === '1' || value === 'true' || value === 'on') {
    return true;
  }
  if (value === '0' || value === 'false' || value === 'off') {
    return false;
  }
  return undefined;
}

/** A relationship id attribute such as r:embed or r:id. */
export function relAttr(element: Element | undefined, name: string): string | undefined {
  if (!element) {
    return undefined;
  }
  const value = element.getAttributeNS(RELATIONSHIPS_NS, name) || element.getAttribute(`r:${name}`);
  return value || undefined;
}
