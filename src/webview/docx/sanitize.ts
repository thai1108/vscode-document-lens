// Sanitizes the DOM docx-preview built from the document, before it is moved into the live page. Only the elements
// and attributes docx-preview creates survive; links keep only http(s) and in-document targets; pictures must be
// blob:/data: images (anything else, e.g. a picture linked from the web, becomes a labelled placeholder); inline
// styles go through the same rules as the style sheets (css.ts).
import { externalUrl } from '../../protocol.ts';
import { clampLengths, isSafeCssValue, sanitizeDeclarations } from './css.ts';

const HTML = 'http://www.w3.org/1999/xhtml';
const SVG = 'http://www.w3.org/2000/svg';
const MATHML = 'http://www.w3.org/1998/Math/MathML';

/** Elements docx-preview creates (HtmlRenderer), by namespace. */
const ALLOWED_ELEMENTS: Readonly<Record<string, ReadonlySet<string>>> = {
  [HTML]: new Set([
    'a', 'article', 'br', 'col', 'colgroup', 'del', 'div', 'footer', 'header', 'img', 'ins', 'li', 'ol', 'p', 'section',
    'span', 'sub', 'sup', 'table', 'td', 'tr', 'wbr',
  ]),
  [SVG]: new Set(['svg', 'g', 'rect', 'ellipse', 'line', 'image', 'foreignObject']),
  [MATHML]: new Set([
    'math', 'mfrac', 'mn', 'mo', 'mover', 'mroot', 'mrow', 'ms', 'msqrt', 'msub', 'msubsup', 'msup', 'mtable', 'mtd',
    'mtr', 'munder', 'munderover',
  ]),
};

const GLOBAL_ATTRIBUTES = new Set(['class', 'id', 'lang', 'style', 'title']);
const ELEMENT_ATTRIBUTES: Readonly<Record<string, ReadonlySet<string>>> = {
  a: new Set(['href']),
  img: new Set(['src']),
  td: new Set(['colspan', 'rowspan']),
  ins: new Set(['datetime']),
  del: new Set(['datetime']),
};
const SVG_LENGTH_ATTRIBUTES = new Set(['width', 'height', 'x1', 'y1', 'x2', 'y2', 'cx', 'cy', 'rx', 'ry', 'stroke-width']);
const SVG_PAINT_ATTRIBUTES = new Set(['fill', 'stroke']);
const SVG_LENGTH = /^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?\s*$/i;

const SHOW_ELEMENT = 1;

export interface SanitizeOptions {
  /** Prefix for every id in the document (bookmarks), so none can collide with the viewer's own ids. */
  idPrefix: string;
}

export const IMAGE_PLACEHOLDER_CLASS = 'dl-docx-image-placeholder';

/** Sanitizes the subtree under `root` in place (`root` itself is the caller's container and is left alone). */
export function sanitizeTree(root: Element, options: SanitizeOptions): void {
  const walker = root.ownerDocument.createTreeWalker(root, SHOW_ELEMENT);
  let node = walker.nextNode();
  while (node) {
    const element = node as Element;
    if (!ALLOWED_ELEMENTS[element.namespaceURI ?? '']?.has(element.localName)) {
      node = skipSubtree(walker);
      element.remove();
      continue;
    }
    if (element.namespaceURI === HTML && element.localName === 'img' && !isImageUrl(element.getAttribute('src'))) {
      node = skipSubtree(walker);
      element.replaceWith(imagePlaceholder(element as HTMLImageElement));
      continue;
    }
    sanitizeElement(element, options);
    node = walker.nextNode();
  }
}

/** Moves the walker past the current node's subtree and returns the next node, or null at the end. */
function skipSubtree(walker: TreeWalker): Node | null {
  for (;;) {
    const sibling = walker.nextSibling();
    if (sibling) {
      return sibling;
    }
    if (!walker.parentNode()) {
      return null;
    }
  }
}

function sanitizeElement(element: Element, options: SanitizeOptions): void {
  const svg = element.namespaceURI === SVG;
  const allowed = ELEMENT_ATTRIBUTES[element.localName];
  for (const { name, value } of Array.from(element.attributes)) {
    const keep =
      GLOBAL_ATTRIBUTES.has(name) ||
      (!svg && allowed?.has(name)) ||
      (svg && (SVG_LENGTH_ATTRIBUTES.has(name) || SVG_PAINT_ATTRIBUTES.has(name) || (name === 'href' && element.localName === 'image')));
    if (!keep) {
      element.removeAttribute(name);
    } else if (svg && SVG_PAINT_ATTRIBUTES.has(name) && !isSafeCssValue(value)) {
      element.removeAttribute(name);
    } else if (svg && SVG_LENGTH_ATTRIBUTES.has(name)) {
      element.setAttribute(name, SVG_LENGTH.test(value) ? clampLengths(`${value.trim()}px`).slice(0, -2) : clampLengths(value));
    }
  }
  if (element.id) {
    element.id = options.idPrefix + element.id;
  }
  if (svg && element.localName === 'image' && !isImageUrl(element.getAttribute('href'))) {
    element.removeAttribute('href');
  }
  if (element.localName === 'a' && !svg) {
    sanitizeLink(element as HTMLAnchorElement, options);
  }
  if (element.hasAttribute('style')) {
    const style = (element as Partial<ElementCSSInlineStyle>).style;
    if (style) {
      sanitizeDeclarations(style);
    } else {
      element.removeAttribute('style');
    }
  }
}

function sanitizeLink(link: HTMLAnchorElement, options: SanitizeOptions): void {
  const href = link.getAttribute('href') ?? '';
  if (href.length > 1 && href.startsWith('#')) {
    link.setAttribute('href', `#${options.idPrefix}${href.slice(1)}`);
    return;
  }
  const url = externalUrl(href);
  if (url) {
    link.setAttribute('href', url);
    if (!link.title) {
      link.title = url;
    }
  } else {
    link.removeAttribute('href');
  }
}

function isImageUrl(url: string | null): boolean {
  return url !== null && /^(?:blob:|data:image\/)/i.test(url.trim());
}

function imagePlaceholder(image: HTMLImageElement): HTMLElement {
  const placeholder = image.ownerDocument.createElement('span');
  placeholder.className = IMAGE_PLACEHOLDER_CLASS;
  placeholder.textContent = 'Picture not shown';
  placeholder.title = 'Only pictures stored inside the file are shown';
  for (const property of ['width', 'height']) {
    const value = image.style.getPropertyValue(property);
    if (value && isSafeCssValue(value)) {
      placeholder.style.setProperty(property, clampLengths(value));
    }
  }
  return placeholder;
}
