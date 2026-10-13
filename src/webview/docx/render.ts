// Turns .docx bytes into sanitized DOM with docx-preview. The library builds every element through the `h` option;
// ours creates them in an inert document, where pictures do not load and nothing is connected to the page, so the
// result can be sanitized before anything in it reaches the live page.
import { parseAsync, renderDocument, type HElement, type Options } from 'docx-preview';
import { sanitizeStyleSheet } from './css.ts';
import { DOCX_LIMITS } from './limits.ts';
import { checkPackage } from './packageCheck.ts';
import { sanitizeTree } from './sanitize.ts';

/** Class of the element the document is rendered into; its style sheets are scoped to it. */
export const DOCUMENT_CLASS = 'dl-docx';
/**
 * Class of the element around it, which zooms and contains the document. The scope root itself can be styled by the
 * document (`:scope { ... }`), so what keeps the document inside its area must sit outside the scope.
 */
export const FRAME_CLASS = 'dl-docx-frame';
/** Prefix of every id in the document (bookmarks). */
const ID_PREFIX = 'dl-docx-';

export interface RenderedDocx {
  /** Sanitized style sheets, in docx-preview's order, for <style nonce> elements in the live page. */
  styles: string[];
  /** Container (in the inert document) whose children are the sanitized pages. */
  content: Element;
  /** blob: URLs created for this document's pictures and fonts; revoke them once the render is replaced. */
  objectUrls: string[];
}

const OPTIONS: Partial<Options> = {
  className: 'docx',
  inWrapper: true,
  breakPages: true,
  ignoreLastRenderedPageBreak: false,
  experimental: false,
  useBase64URL: false,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
  renderComments: false,
  renderChanges: false,
  // An altChunk is a whole HTML/RTF document embedded in the file; docx-preview would show it in an iframe.
  renderAltChunks: false,
  debug: false,
};

export async function renderDocx(data: Uint8Array): Promise<RenderedDocx> {
  await checkPackage(data);
  const inert = document.implementation.createHTMLDocument('');
  const options: Partial<Options> = { ...OPTIONS, h: elementFactory(inert) };

  let parsed: unknown;
  try {
    parsed = await parseAsync(data, options);
  } catch (error) {
    throw new Error(`the document is damaged (${describe(error)})`, { cause: error });
  }
  checkModel(parsed);

  let nodes: Node[];
  try {
    nodes = await renderDocument(parsed, options);
  } catch (error) {
    throw new Error(`the document could not be laid out (${describe(error)})`, { cause: error });
  }
  const rawStyles: string[] = [];
  const content = inert.createElement('div');
  for (const node of nodes) {
    if (node.nodeName === 'STYLE') {
      rawStyles.push(node.textContent ?? '');
    } else if (node.nodeType === 1) {
      content.append(node);
    }
  }
  const objectUrls = collectObjectUrls(content, rawStyles);
  sanitizeTree(content, { idPrefix: ID_PREFIX });
  return {
    styles: rawStyles.map((text) => sanitizeStyleSheet(text, `.${DOCUMENT_CLASS}`)).filter(Boolean),
    content,
    objectUrls,
  };
}

/**
 * docx-preview's `h`, building elements in `owner`. Styles are applied through the CSSOM: a style attribute would be
 * refused by the webview's CSP once the element reaches the page, CSSOM declarations are not.
 */
export function elementFactory(owner: Document): (element: HElement | Node | string) => Node {
  const h = (element: HElement | Node | string): Node => {
    if (typeof element === 'string') {
      return owner.createTextNode(element);
    }
    if (isNode(element)) {
      return element;
    }
    const { ns, tagName, className, style, children, ...properties } = element;
    if (tagName === '#fragment') {
      return appendChildren(owner.createDocumentFragment(), children);
    }
    if (tagName === '#comment') {
      return owner.createComment('');
    }
    const node = ns ? owner.createElementNS(ns, tagName) : owner.createElement(tagName);
    if (className) {
      node.setAttribute('class', className);
    }
    if (style) {
      applyStyle(node, style);
    }
    for (const [key, value] of Object.entries(properties)) {
      if (value !== undefined) {
        Reflect.set(node, key, value);
      }
    }
    return appendChildren(node, children);
  };
  // One appendChild per child: a spread could exceed the argument limit on a huge table. docx-preview passes null for
  // absent MathML operands.
  const appendChildren = <T extends Node>(parent: T, children: HElement['children']): T => {
    for (const child of children ?? []) {
      if (child !== null && child !== undefined) {
        parent.appendChild(h(child));
      }
    }
    return parent;
  };
  return h;
}

function isNode(value: object): value is Node {
  return typeof (value as Partial<Node>).nodeType === 'number';
}

function applyStyle(node: Element, style: Record<string, string> | string): void {
  const declarations = (node as Partial<ElementCSSInlineStyle>).style;
  if (!declarations) {
    return;
  }
  if (typeof style === 'string') {
    declarations.cssText = style;
    return;
  }
  for (const [name, value] of Object.entries(style)) {
    if (value !== undefined && value !== null && !name.startsWith('$')) {
      const property = name.includes('-') ? name : name.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);
      declarations.setProperty(property, String(value));
    }
  }
}

interface ModelNode {
  type?: unknown;
  break?: unknown;
  sectionProps?: unknown;
  children?: unknown;
  notes?: unknown;
  rootElement?: unknown;
}

interface ModelSize {
  nodes: number;
  depth: number;
  pageBreaks: number;
}

/**
 * Refuses documents whose render would be too large or too deep, estimating the DOM docx-preview would build: the
 * body and notes once, a header and a footer on every page.
 */
export function checkModel(parsed: unknown): void {
  const parts = ((parsed as { parts?: unknown }).parts ?? []) as ModelNode[];
  const root = (parsed as { documentPart?: { body?: unknown } }).documentPart?.body;
  if (typeof root !== 'object' || root === null) {
    throw new Error('the document is damaged (its main part is missing)');
  }
  const body = measure([root]);
  const notes = measure(parts.flatMap((part) => (Array.isArray(part.notes) ? part.notes : [])));
  let header = 0;
  let footer = 0;
  let depth = Math.max(body.depth, notes.depth);
  for (const part of parts) {
    const partRoot = part.rootElement as ModelNode | undefined;
    if (partRoot?.type === 'header' || partRoot?.type === 'footer') {
      const size = measure([partRoot]);
      depth = Math.max(depth, size.depth);
      if (partRoot.type === 'header') {
        header = Math.max(header, size.nodes);
      } else {
        footer = Math.max(footer, size.nodes);
      }
    }
  }
  if (depth > DOCX_LIMITS.maxDepth) {
    throw new Error(`the document is nested too deeply to preview (more than ${DOCX_LIMITS.maxDepth} levels)`);
  }
  const estimate = body.nodes + notes.nodes + (body.pageBreaks + 1) * (header + footer);
  if (estimate > DOCX_LIMITS.maxRenderedNodes) {
    throw new Error(`the document is too large to preview (about ${estimate.toLocaleString('en-US')} elements; limit ${DOCX_LIMITS.maxRenderedNodes.toLocaleString('en-US')})`);
  }
}

/** Counts the model under `roots` through `children` arrays, iteratively (the model can be arbitrarily deep). */
function measure(roots: unknown[]): ModelSize {
  const size: ModelSize = { nodes: 0, depth: 0, pageBreaks: 0 };
  const stack: [unknown, number][] = roots.map((root) => [root, 1]);
  for (let item = stack.pop(); item; item = stack.pop()) {
    const [node, depth] = item;
    if (typeof node !== 'object' || node === null) {
      continue;
    }
    const model = node as ModelNode;
    size.nodes += 1;
    size.depth = Math.max(size.depth, depth);
    if ((model.type === 'break' && model.break !== 'textWrapping') || model.sectionProps) {
      size.pageBreaks += 1;
    }
    if (Array.isArray(model.children)) {
      for (const child of model.children) {
        stack.push([child, depth + 1]);
      }
    }
    if (size.depth > DOCX_LIMITS.maxDepth) {
      break;
    }
  }
  return size;
}

function collectObjectUrls(content: Element, styles: readonly string[]): string[] {
  const urls = new Set<string>();
  for (const element of content.querySelectorAll('img[src^="blob:"], image[href^="blob:"]')) {
    urls.add(element.getAttribute('src') ?? element.getAttribute('href') ?? '');
  }
  for (const text of styles) {
    for (const [url] of text.matchAll(/blob:[^\s"')]+/g)) {
      urls.add(url);
    }
  }
  urls.delete('');
  return [...urls];
}

/** A short reason for an error from docx-preview or the XML parser, for the error banner. */
function describe(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  const xml = /error on line \d+ at column \d+: [^\n]*/.exec(message)?.[0];
  const reason = (xml ?? message.split('\n')[0] ?? '').trim();
  return reason.length > 200 ? `${reason.slice(0, 200)}…` : reason || 'unknown error';
}
