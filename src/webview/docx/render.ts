// Turns .docx bytes into sanitized DOM with docx-preview. The library builds every element through the `h` option;
// ours creates them in an inert document, where pictures do not load and nothing is connected to the page, so the
// result can be sanitized before anything in it reaches the live page.
import { parseAsync, renderDocument, type HElement, type Options } from 'docx-preview';
import { formatMegabytes } from '../../limits.ts';
import { embeddedFontNames, sanitizeStyleSheet } from './css.ts';
import { DOCX_LIMITS } from './limits.ts';
import { shareMedia } from './media.ts';
import { checkModel } from './model.ts';
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
  /** For the notice banner: what was left out, if anything. */
  notice: string | null;
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
  const { sizes } = await checkPackage(data);
  const inert = document.implementation.createHTMLDocument('');
  const options: Partial<Options> = { ...OPTIONS, h: elementFactory(inert) };

  let parsed: unknown;
  try {
    parsed = await parseAsync(data, options);
  } catch (error) {
    throw new Error(`the document is damaged (${describe(error)})`, { cause: error });
  }
  checkModel(parsed);
  const media = shareMedia(parsed, sizes);
  try {
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
    const fonts = embeddedFontNames(rawStyles);
    sanitizeTree(content, { idPrefix: ID_PREFIX, fonts });
    return {
      styles: rawStyles.map((text) => sanitizeStyleSheet(text, `.${DOCUMENT_CLASS}`, fonts)).filter(Boolean),
      content,
      objectUrls: media.objectUrls,
      notice: media.overBudget
        ? `Some pictures or fonts are not shown: together they exceed ${formatMegabytes(DOCX_LIMITS.maxMediaSize)}, the most Document Lens loads from one document.`
        : null,
    };
  } catch (error) {
    for (const url of media.objectUrls) {
      URL.revokeObjectURL(url);
    }
    throw error;
  }
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

/** A short reason for an error from docx-preview or the XML parser, for the error banner. */
function describe(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  const xml = /error on line \d+ at column \d+: [^\n]*/.exec(message)?.[0];
  const reason = (xml ?? message.split('\n')[0] ?? '').trim();
  return reason.length > 200 ? `${reason.slice(0, 200)}…` : reason || 'unknown error';
}
