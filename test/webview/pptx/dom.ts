// happy-dom globals for the presentation renderer's unit tests (DOMParser for the package parts, document for the
// DOM it builds). node:test runs every test file in its own process, so the globals do not leak between files.
import { Window } from 'happy-dom';
import { Deck } from '../../../src/webview/pptx/deck.ts';
import { MediaStore } from '../../../src/webview/pptx/media.ts';
import { renderSlide, type RenderedSlide } from '../../../src/webview/pptx/slide.ts';

const GLOBALS = ['document', 'HTMLElement', 'Element', 'Node'] as const;

export function installDom(): Window {
  const window = new Window({ url: 'https://webview.test/' });
  const scope = globalThis as unknown as Record<string, unknown>;
  const source = window as unknown as Record<string, unknown>;
  for (const name of GLOBALS) {
    scope[name] = source[name];
  }
  scope['DOMParser'] = conformingParser(window);
  return window;
}

/**
 * happy-dom keeps only one of `id` and `r:id` when an element has both (p:sldId, p:sldMasterId, p:sldLayoutId), where
 * browsers keep both. The renderer only reads `r:id` there, so the plain `id` is renamed before parsing.
 */
export function conformingParser(window: Window): typeof DOMParser {
  const Parser = window.DOMParser;
  return class extends Parser {
    override parseFromString(text: string, type: Parameters<InstanceType<typeof Parser>['parseFromString']>[1]) {
      return super.parseFromString(text.replace(/(<(?:\w+:)?sld(?:Master|Layout)?Id\b[^>]*?)\sid=/g, '$1 data-id='), type);
    }
  } as unknown as typeof DOMParser;
}

export function xml(text: string): Element {
  const root = new DOMParser().parseFromString(text, 'application/xml').documentElement;
  if (!root) {
    throw new Error('test XML did not parse');
  }
  return root;
}

export const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** Renders every slide of a package, as the viewer does. */
export async function renderAll(bytes: Uint8Array): Promise<{ deck: Deck; slides: RenderedSlide[]; media: MediaStore }> {
  const deck = await Deck.open(new Uint8Array(bytes));
  const media = new MediaStore(deck.pkg);
  const slides: RenderedSlide[] = [];
  for (let index = 0; index < deck.slideParts.length; index++) {
    slides.push(await renderSlide(deck, index, media));
  }
  return { deck, slides, media };
}

export function byName(root: Element, name: string): HTMLElement {
  const found = [...root.querySelectorAll<HTMLElement>('[data-name]')].find((element) => element.dataset['name'] === name);
  if (!found) {
    throw new Error(`no shape named ${name}`);
  }
  return found;
}
