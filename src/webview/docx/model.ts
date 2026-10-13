// Estimates how much DOM docx-preview would build from a parsed document (WordDocument) and refuses documents that
// would be too large or too deep, before anything is rendered. The estimate follows HtmlRenderer: the body once, a
// header and a footer on every page, and a footnote or endnote once per reference to it.
import { DOCX_LIMITS } from './limits.ts';

interface ModelNode {
  type?: unknown;
  id?: unknown;
  break?: unknown;
  styleName?: unknown;
  pageBreakBefore?: unknown;
  sectionProps?: unknown;
  children?: unknown;
}

interface ParsedDocument {
  documentPart?: { body?: unknown };
  parts?: { rootElement?: unknown }[];
  footnotesPart?: { notes?: unknown };
  endnotesPart?: { notes?: unknown };
  stylesPart?: { styles?: unknown };
}

interface Style {
  id?: unknown;
  basedOn?: unknown;
  paragraphProps?: { pageBreakBefore?: unknown };
}

interface Measure {
  nodes: number;
  depth: number;
  /** Note references by note id, with how often each occurs. */
  footnotes: Map<string, number>;
  endnotes: Map<string, number>;
}

export function checkModel(parsed: unknown): void {
  const doc = parsed as ParsedDocument;
  const root = doc.documentPart?.body;
  if (typeof root !== 'object' || root === null) {
    throw new Error('the document is damaged (its main part is missing)');
  }
  const footnotes = noteSizes(doc.footnotesPart?.notes);
  const endnotes = noteSizes(doc.endnotesPart?.notes);
  /** Nodes of a measured tree plus the notes it makes docx-preview render. */
  const cost = (size: Measure) => size.nodes + referencedNodes(size.footnotes, footnotes) + referencedNodes(size.endnotes, endnotes);

  const body = measure(root);
  let depth = Math.max(body.depth, footnotes.depth, endnotes.depth);
  let header = 0;
  let footer = 0;
  for (const part of doc.parts ?? []) {
    const partRoot = part.rootElement as ModelNode | undefined;
    if (partRoot?.type === 'header' || partRoot?.type === 'footer') {
      const size = measure(partRoot);
      depth = Math.max(depth, size.depth);
      if (partRoot.type === 'header') {
        header = Math.max(header, cost(size));
      } else {
        footer = Math.max(footer, cost(size));
      }
    }
  }
  if (depth > DOCX_LIMITS.maxDepth) {
    throw new Error(`the document is nested too deeply to preview (more than ${DOCX_LIMITS.maxDepth} levels)`);
  }
  const pages = countPages(root as ModelNode, pageBreakStyles(doc.stylesPart?.styles));
  const estimate = cost(body) + pages * (header + footer);
  if (estimate > DOCX_LIMITS.maxRenderedNodes) {
    throw new Error(
      `the document is too large to preview (about ${estimate.toLocaleString('en-US')} elements; limit ${DOCX_LIMITS.maxRenderedNodes.toLocaleString('en-US')})`,
    );
  }
}

/** Counts the model under `root` through `children` arrays, iteratively (the model can be arbitrarily deep). */
function measure(root: unknown): Measure {
  const size: Measure = { nodes: 0, depth: 0, footnotes: new Map(), endnotes: new Map() };
  const stack: [unknown, number][] = [[root, 1]];
  for (let item = stack.pop(); item; item = stack.pop()) {
    const [node, depth] = item;
    if (typeof node !== 'object' || node === null) {
      continue;
    }
    const model = node as ModelNode;
    size.nodes += 1;
    size.depth = Math.max(size.depth, depth);
    if (size.depth > DOCX_LIMITS.maxDepth) {
      break;
    }
    const references = model.type === 'footnoteReference' ? size.footnotes : model.type === 'endnoteReference' ? size.endnotes : undefined;
    if (references) {
      const id = String(model.id);
      references.set(id, (references.get(id) ?? 0) + 1);
    }
    if (Array.isArray(model.children)) {
      for (const child of model.children) {
        stack.push([child, depth + 1]);
      }
    }
  }
  return size;
}

/** Each note's size by id, and the deepest note. */
function noteSizes(notes: unknown): Map<string, number> & { depth: number } {
  const sizes = Object.assign(new Map<string, number>(), { depth: 0 });
  for (const note of Array.isArray(notes) ? notes : []) {
    const size = measure(note);
    sizes.set(String((note as ModelNode).id), size.nodes);
    sizes.depth = Math.max(sizes.depth, size.depth);
  }
  return sizes;
}

function referencedNodes(references: ReadonlyMap<string, number>, sizes: ReadonlyMap<string, number>): number {
  let nodes = 0;
  for (const [id, count] of references) {
    nodes += count * (sizes.get(id) ?? 0);
  }
  return nodes;
}

/**
 * Pages the body is split into, counted as HtmlRenderer.splitBySection may split it (an upper bound): a paragraph
 * starts a page when it or its style breaks the page before it, when it ends a section, and when one of its runs holds
 * a page break or Word's last rendered page break.
 */
function countPages(body: ModelNode, breakingStyles: ReadonlySet<string>): number {
  let pages = 1;
  for (const element of children(body)) {
    if (element.type !== 'paragraph') {
      continue;
    }
    if (element.pageBreakBefore === true || (typeof element.styleName === 'string' && breakingStyles.has(element.styleName))) {
      pages++;
    }
    if (element.sectionProps) {
      pages++;
    }
    if (children(element).some((run) => children(run).some((child) => child.type === 'break' && (child.break === 'page' || child.break === 'lastRenderedPageBreak')))) {
      pages++;
    }
  }
  return pages;
}

/** Paragraph styles that break the page before, directly or through the styles they are based on. */
function pageBreakStyles(styles: unknown): Set<string> {
  const byId = new Map<string, Style>();
  for (const style of Array.isArray(styles) ? (styles as Style[]) : []) {
    if (typeof style?.id === 'string') {
      byId.set(style.id, style);
    }
  }
  const breaking = new Set<string>();
  for (const id of byId.keys()) {
    const seen = new Set<string>();
    for (let style = byId.get(id); style && typeof style.id === 'string' && !seen.has(style.id); style = byId.get(String(style.basedOn))) {
      seen.add(style.id);
      if (style.paragraphProps?.pageBreakBefore === true) {
        breaking.add(id);
        break;
      }
    }
  }
  return breaking;
}

function children(node: ModelNode): ModelNode[] {
  return Array.isArray(node.children) ? (node.children.filter((child) => typeof child === 'object' && child !== null) as ModelNode[]) : [];
}
