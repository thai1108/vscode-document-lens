// The .docx viewer: pages rendered by docx-preview (render.ts), its own zoom so the scroll position can follow, fit
// width, and the scroll position kept across reloads of the file and of the webview.
import type { LoadedDocument, Viewer, ViewerContext } from '../shell/app.ts';
import { DOCUMENT_CLASS, FRAME_CLASS, renderDocx } from './render.ts';

/** Horizontal padding around the pages, in CSS px at 100%; docx.css uses the same value. */
const PAGE_GAP = 24;
const SAVE_DELAY_MS = 150;

/** Scroll position in CSS px at 100% zoom, so it survives zoom changes. */
interface DocxViewerState {
  top: number;
  left: number;
}

export function docxViewer(): Viewer {
  /** The content area, the frame (zoomed) and the element the pages are rendered into, from the first render on. */
  let view: { context: ViewerContext; frame: HTMLElement; host: HTMLElement } | undefined;
  let zoom = 1;
  let styles: HTMLStyleElement[] = [];
  let objectUrls: string[] = [];
  let saveTimer: ReturnType<typeof setTimeout> | undefined;

  const position = (content: HTMLElement): DocxViewerState => ({ top: content.scrollTop / zoom, left: content.scrollLeft / zoom });

  const connect = (context: ViewerContext) => {
    const frame = document.createElement('div');
    frame.className = FRAME_CLASS;
    const host = frame.appendChild(document.createElement('div'));
    host.className = DOCUMENT_CLASS;
    const { content } = context;
    content.addEventListener('scroll', () => {
      clearTimeout(saveTimer);
      saveTimer = setTimeout(() => context.saveViewerState(position(content)), SAVE_DELAY_MS);
    }, { passive: true });
    // In-document links (bookmarks, a table of contents) scroll to their target inside the pages.
    content.addEventListener('click', (event) => {
      const link = (event.target as Element | null)?.closest?.('a[href^="#"]');
      if (!link || !host.contains(link)) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      const target = document.getElementById(link.getAttribute('href')?.slice(1) ?? '');
      if (target && host.contains(target)) {
        target.scrollIntoView({ block: 'start' });
      }
    });
    return { context, frame, host };
  };

  return {
    async render(file: LoadedDocument, context: ViewerContext, reload: boolean) {
      view ??= connect(context);
      const { frame, host } = view;
      const { content } = context;
      const saved = reload && frame.isConnected ? position(content) : savedPosition(context.viewerState());
      // Nothing on the page changes until the new document has rendered completely.
      const rendered = await renderDocx(file.data);

      const nextStyles = rendered.styles.map((text) => {
        const style = document.createElement('style');
        style.setAttribute('nonce', context.styleNonce);
        style.textContent = text;
        return style;
      });
      document.head.append(...nextStyles);
      host.replaceChildren(...Array.from(rendered.content.childNodes));
      if (frame.parentNode !== content) {
        content.replaceChildren(frame);
      }
      for (const style of styles) {
        style.remove();
      }
      for (const url of objectUrls) {
        URL.revokeObjectURL(url);
      }
      styles = nextStyles;
      objectUrls = rendered.objectUrls;
      context.showNotice(rendered.notice);

      zoom = context.zoom();
      frame.style.setProperty('zoom', String(zoom));
      sizeVectorShapes(host);
      if (saved) {
        content.scrollTop = saved.top * zoom;
        content.scrollLeft = saved.left * zoom;
      }
    },

    applyZoom(next: number) {
      if (!view?.frame.isConnected || next === zoom) {
        zoom = next;
        return;
      }
      // Keep the point at the centre of the view where it is.
      const { content } = view.context;
      const centerX = (content.scrollLeft + content.clientWidth / 2) / zoom;
      const centerY = (content.scrollTop + content.clientHeight / 2) / zoom;
      zoom = next;
      view.frame.style.setProperty('zoom', String(next));
      content.scrollLeft = centerX * next - content.clientWidth / 2;
      content.scrollTop = centerY * next - content.clientHeight / 2;
    },

    fitWidthZoom() {
      const pages = view?.host.querySelectorAll('.docx-wrapper > section') ?? [];
      if (!view || pages.length === 0) {
        return undefined;
      }
      // getBoundingClientRect includes the zoom; the widest page at 100% plus the gap on both sides must fit.
      let widest = 0;
      for (const page of pages) {
        widest = Math.max(widest, page.getBoundingClientRect().width / zoom);
      }
      const available = view.context.content.clientWidth;
      return widest > 0 && available > 0 ? available / (widest + 2 * PAGE_GAP) : undefined;
    },
  };
}

/** The position saved in the webview state, if it is one (the state outlives versions of the viewer). */
function savedPosition(state: unknown): DocxViewerState | undefined {
  const { top, left } = (state ?? {}) as Partial<Record<keyof DocxViewerState, unknown>>;
  return typeof top === 'number' && typeof left === 'number' && Number.isFinite(top) && Number.isFinite(left) ? { top, left } : undefined;
}

/**
 * docx-preview sizes each VML shape's <svg> from its content's bounding box one frame after creating it; in the
 * inert document that box is empty, so the shapes are measured again once they are on the page.
 */
function sizeVectorShapes(root: Element): void {
  for (const svg of root.querySelectorAll('svg')) {
    const shape = svg.firstElementChild as SVGGraphicsElement | null;
    if (!shape || typeof shape.getBBox !== 'function') {
      continue;
    }
    try {
      const box = shape.getBBox();
      if (box.width > 0 || box.height > 0) {
        svg.setAttribute('width', String(Math.ceil(box.x + box.width)));
        svg.setAttribute('height', String(Math.ceil(box.y + box.height)));
      }
    } catch {
      // Not rendered (e.g. inside a hidden element): keep docx-preview's size.
    }
  }
}
