// OPC relationship parts (`_rels/*.rels`): how a part names the parts it uses.
import { scanXml } from './xml.ts';

export interface Relationship {
  /** Package part name (resolved against the source part) or, if `external`, the URL as written. */
  target: string;
  /** Last segment of the relationship type URI, e.g. `worksheet` (Transitional and Strict URIs end alike). */
  type: string;
  external: boolean;
}

/** The relationships part of a part: `xl/workbook.xml` → `xl/_rels/workbook.xml.rels`. */
export function relationshipsPartOf(part: string): string {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

export function parseRelationships(xml: string, sourcePart: string): Map<string, Relationship> {
  const relationships = new Map<string, Relationship>();
  const base = sourcePart.slice(0, sourcePart.lastIndexOf('/') + 1);
  scanXml(xml, {
    open(tag) {
      if (tag.name !== 'Relationship') {
        return;
      }
      const id = tag.attr('Id');
      const target = tag.attr('Target');
      const type = tag.attr('Type') ?? '';
      if (!id || target === undefined || relationships.has(id)) {
        return;
      }
      const external = tag.attr('TargetMode') === 'External';
      relationships.set(id, {
        target: external ? target : resolvePartName(base, target),
        type: type.slice(type.lastIndexOf('/') + 1),
        external,
      });
    },
  });
  return relationships;
}

/** Resolves a relative target against a folder (`xl/`), normalizing `.` and `..`; a leading `/` is the package root. */
export function resolvePartName(base: string, target: string): string {
  const path = target.startsWith('/') ? target.slice(1) : base + target;
  const parts: string[] = [];
  for (const segment of path.split('/')) {
    if (segment === '..') {
      parts.pop();
    } else if (segment !== '.' && segment !== '') {
      parts.push(percentDecode(segment));
    }
  }
  return parts.join('/');
}

/** Targets are URIs, so a part named `sheet 1.xml` is written `sheet%201.xml`. */
function percentDecode(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}
