// The document outline (bookmarks) as an indented list beside the pages. The outline comes from the file, so its
// size and depth are capped and titles are only ever set as text.
import { element } from '../shell/dom.ts';

/** One node of PDFDocumentProxy.getOutline(), reduced to what the panel uses. */
export interface OutlineNode {
  title: string;
  /** Internal destination: a named destination or an explicit destination array. */
  dest: string | unknown[] | null;
  /** External link, already resolved to an absolute URL by pdf.js. */
  url: string | null;
  items: OutlineNode[];
}

export interface OutlineEntry {
  node: OutlineNode;
  depth: number;
}

export const OUTLINE_LIMITS = { maxEntries: 5000, maxDepth: 12 };

/** The outline in reading order, depth first, stopping at OUTLINE_LIMITS. */
export function flattenOutline(nodes: readonly OutlineNode[], limits = OUTLINE_LIMITS): OutlineEntry[] {
  const entries: OutlineEntry[] = [];
  const stack: OutlineEntry[] = [...nodes].reverse().map((node) => ({ node, depth: 0 }));
  while (stack.length > 0 && entries.length < limits.maxEntries) {
    const entry = stack.pop() as OutlineEntry;
    entries.push(entry);
    const children = Array.isArray(entry.node.items) ? entry.node.items : [];
    if (entry.depth + 1 < limits.maxDepth) {
      for (let i = children.length - 1; i >= 0; i--) {
        stack.push({ node: children[i] as OutlineNode, depth: entry.depth + 1 });
      }
    }
  }
  return entries;
}

/** Fills `panel` with one button per entry; `onSelect` gets the entry the reader picked. */
export function renderOutline(panel: HTMLElement, entries: readonly OutlineEntry[], onSelect: (node: OutlineNode) => void): void {
  const list = element('ul', { className: 'pdf-outline-list', attributes: { role: 'tree', 'aria-label': 'Outline' } });
  for (const { node, depth } of entries) {
    const title = typeof node.title === 'string' ? node.title.trim() : '';
    const button = element('button', {
      className: 'pdf-outline-item',
      text: title || 'Untitled',
      attributes: { type: 'button', role: 'treeitem', 'aria-level': String(depth + 1), title: title || 'Untitled' },
    });
    button.style.setProperty('--depth', String(depth));
    button.addEventListener('click', () => onSelect(node));
    const item = element('li', { attributes: { role: 'none' } });
    item.append(button);
    list.append(item);
  }
  panel.replaceChildren(list);
}
