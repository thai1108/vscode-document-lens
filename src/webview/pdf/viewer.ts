// The PDF viewer. pdf.js's PDFViewer component lays out every page as a placeholder sized from its viewport and
// renders only the pages near the viewport (canvas, text layer, annotation layer), dropping those scrolled far away.
// This module wires it to the shell (zoom, saved position, links, errors) and to the toolbar (pages, find, outline).
import type { LoadedDocument, Viewer, ViewerContext } from '../shell/app.ts';
import { element } from '../shell/dom.ts';
import { embeddedDataFactory, type EmbeddedFiles } from './assets.ts';
import { loadErrorMessage, pageCountError } from './errors.ts';
import { fitWidthZoom, parseViewState, type PdfViewState } from './navigation.ts';
import { flattenOutline, renderOutline, type OutlineNode } from './outline.ts';
import {
  AnnotationEditorType,
  AnnotationMode,
  EventBus,
  FindState,
  getDocument,
  PDFFindController,
  PDFLinkService,
  PDFViewer,
  PDFWorker,
  ScrollMode,
  VerbosityLevel,
  type PDFDocumentLoadingTask,
  type PDFDocumentProxy,
} from './pdfjs.ts';
import { PdfToolbar } from './toolbar.ts';

export interface PdfViewerOptions {
  /** Starts pdf.js's worker; called once, for the first document. */
  startWorker(): Worker;
  /** CMaps, standard fonts and wasm decoders pdf.js may ask for. */
  files: EmbeddedFiles;
}

/** Space around and between pages in CSS pixels (--pdf-gutter in pdf.css). */
const PAGE_GUTTER = 12;
/** Images above this many pixels are left out instead of decoded: a hostile file can declare any size. */
const MAX_IMAGE_PIXELS = 2 ** 26;
const SAVE_DELAY_MS = 250;
/** Wheel events closer together than this are one gesture, which turns at most one page in single-page mode. */
const WHEEL_GESTURE_GAP_MS = 200;
/**
 * pdf.js shows a note's icon by appending "annotation-<name>.svg" (a name taken from the file) to this prefix. After
 * the `#` that is only a URL fragment, so every note gets this icon and nothing is ever requested.
 */
const NOTE_ICON_PREFIX = `data:image/svg+xml,${encodeURIComponent(
  "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 22 22' fill='none' stroke='#7a5c00'>" +
    "<path d='M3 2.5h16l.5.5v11.5l-5 5H3l-.5-.5V3z' fill='#ffd84d'/><path d='M14.5 19.5v-5h5' fill='#f0bd1a'/>" +
    "<path d='M6 7h10M6 10h10M6 13h5'/></svg>",
)}#`;

export function createPdfViewer(options: PdfViewerOptions): Viewer {
  const BinaryDataFactory = embeddedDataFactory(options.files);
  let worker: PDFWorker | undefined;
  let view: PdfView | undefined;
  return {
    async render(document: LoadedDocument, context: ViewerContext, reload: boolean) {
      worker ??= PDFWorker.create({ port: options.startWorker(), verbosity: VerbosityLevel.ERRORS });
      const task = getDocument({
        data: document.data,
        worker,
        BinaryDataFactory,
        useWorkerFetch: false,
        maxImageSize: MAX_IMAGE_PIXELS,
        verbosity: VerbosityLevel.ERRORS,
      });
      const pdf = await openDocument(task);
      // Created on the first document that opens, so a file that fails shows only the error banner.
      view ??= new PdfView(context);
      await view.show(task, pdf, reload);
    },
    applyZoom() {
      view?.applyZoom();
    },
    fitWidthZoom() {
      return view?.fitWidthZoom();
    },
  };
}

async function openDocument(task: PDFDocumentLoadingTask): Promise<PDFDocumentProxy> {
  try {
    const pdf = await task.promise;
    const pageCountProblem = pageCountError(pdf.numPages);
    if (pageCountProblem) {
      throw new Error(pageCountProblem);
    }
    // PDFViewer only logs a first page that cannot be read; checking it here turns that into the error banner.
    await pdf.getPage(1);
    return pdf;
  } catch (error) {
    void task.destroy();
    throw new Error(loadErrorMessage(error), { cause: error });
  }
}

/** Where pdf.js says the view is (PDFViewer's `updateviewarea` event). */
interface ViewArea {
  pageNumber: number;
  left: number;
  top: number;
}

class PdfView {
  readonly #context: ViewerContext;
  readonly #eventBus = new EventBus();
  readonly #linkService: PDFLinkService;
  readonly #viewer: PDFViewer;
  readonly #toolbar: PdfToolbar;
  readonly #layout: HTMLElement;
  readonly #scroller: HTMLDivElement;
  readonly #outlinePanel: HTMLElement;
  #task: PDFDocumentLoadingTask | undefined;
  /** Listeners that belong to the shown document. */
  #documentEvents = new AbortController();
  /** Where the reader is, kept up to date while they scroll. */
  #location: PdfViewState | undefined;
  /** Zoom (and a position to restore) waiting until the pages are laid out and visible. */
  #pendingView: { location?: PdfViewState } | undefined;
  /** The reader scrolled or navigated since the document was shown, so a late position restore must not jump. */
  #moved = false;
  /** A file's first opening starts at fit width when its page is wider than the view at 100%. */
  #fitFirstView = false;
  #outlineOpen: boolean;
  #hasOutline = false;
  #findNotFound = false;
  #saveTimer: ReturnType<typeof setTimeout> | undefined;
  #lastWheel = Number.NEGATIVE_INFINITY;
  #turnedInGesture = false;

  constructor(context: ViewerContext) {
    this.#context = context;
    this.#outlineOpen = parseViewState(context.viewerState())?.outline ?? false;
    const pages = element('div', { className: 'pdfViewer', attributes: { id: 'pdf-pages' } });
    this.#scroller = element('div', { className: 'pdf-scroller', attributes: { id: 'pdf-scroller', tabindex: '0' } });
    this.#scroller.append(pages);
    this.#outlinePanel = element('nav', { className: 'pdf-outline', attributes: { id: 'pdf-outline', 'aria-label': 'Outline' } });
    this.#outlinePanel.hidden = true;
    this.#layout = element('div', { className: 'pdf-layout' });
    this.#layout.append(this.#outlinePanel, this.#scroller);
    context.content.replaceChildren(this.#layout);

    const eventBus = this.#eventBus;
    this.#linkService = new PDFLinkService({ eventBus });
    this.#viewer = new PDFViewer({
      container: this.#scroller,
      viewer: pages,
      eventBus,
      linkService: this.#linkService,
      findController: new PDFFindController({ eventBus, linkService: this.#linkService }),
      // Links and notes in the annotation layer; form fields stay as drawn on the canvas.
      annotationMode: AnnotationMode.ENABLE,
      annotationEditorMode: AnnotationEditorType.DISABLE,
      imageResourcesPath: NOTE_ICON_PREFIX,
      removePageBorders: true,
      enableSelectionRendering: false,
    });
    this.#linkService.setViewer(this.#viewer);

    this.#toolbar = new PdfToolbar(context.tools, {
      goToPage: (page) => {
        this.#moved = true;
        this.#viewer.currentPageNumber = page;
      },
      find: (query, { again, previous }) => this.#find(query, again, previous),
      closeFind: () => eventBus.dispatch('findbarclose', { source: this }),
      toggleOutline: () => {
        this.#outlineOpen = !this.#outlineOpen;
        this.#updateOutline();
        this.#scheduleSave();
      },
      focusDocument: () => this.#scroller.focus({ preventScroll: true }),
    });

    eventBus.on('pagechanging', ({ pageNumber }: { pageNumber: number }) => this.#toolbar.setPage(pageNumber));
    eventBus.on('updateviewarea', ({ location }: { location?: ViewArea }) => this.#track(location));
    eventBus.on('updatefindmatchescount', ({ matchesCount }: { matchesCount: { current: number; total: number } }) =>
      this.#toolbar.setMatches({ ...matchesCount, notFound: this.#findNotFound }),
    );
    eventBus.on('updatefindcontrolstate', ({ state, matchesCount }: { state: number; matchesCount: { current: number; total: number } }) => {
      this.#findNotFound = state === FindState.NOT_FOUND;
      this.#toolbar.setMatches({ ...matchesCount, notFound: this.#findNotFound });
    });
    for (const type of ['wheel', 'pointerdown', 'keydown', 'touchstart']) {
      this.#scroller.addEventListener(type, () => (this.#moved = true), { passive: true });
    }
    // Above 10,000 pages pdf.js shows one page at a time (ScrollMode.PAGE); scrolling on at its edge turns the page.
    this.#scroller.addEventListener('wheel', (event) => this.#onWheel(event), { passive: false });
    this.#scroller.addEventListener('keydown', (event) => this.#onKey(event));
    // A hidden webview (a background tab) has no layout; zoom and scroll wait until it is shown again.
    new ResizeObserver(() => this.#applyPendingView()).observe(this.#scroller);
    window.addEventListener('focus', () => this.#focusIfIdle());
  }

  async show(task: PDFDocumentLoadingTask, pdf: PDFDocumentProxy, reload: boolean): Promise<void> {
    const restore = reload ? this.#location : parseViewState(this.#context.viewerState());
    const previous = this.#task;
    this.#task = task;
    this.#moved = false;
    this.#documentEvents.abort();
    const { signal } = (this.#documentEvents = new AbortController());
    const pagesInit = new Promise<void>((resolve) => this.#eventBus.on('pagesinit', () => resolve(), { once: true, signal }));
    this.#eventBus.on(
      'pagesloaded',
      () => {
        // Pages are first laid out at the first page's size; once all sizes are known a restored position is
        // applied again so that pages of other sizes before it do not shift it.
        if (restore && !this.#moved) {
          this.#queueView(restore);
        }
      },
      { once: true, signal },
    );

    this.#linkService.setDocument(pdf);
    this.#viewer.setDocument(pdf);
    // The viewer has let go of the previous document; free it in the worker too.
    void previous?.destroy();
    // When PDFViewer cannot lay out the pages it never sends pagesinit, but it rejects pagesPromise.
    await Promise.race([pagesInit, this.#viewer.pagesPromise as Promise<void>]);
    this.#toolbar.setPageCount(pdf.numPages);
    // Matches of the previous version are gone. Searching again here would scroll to a match, so the query stays in
    // the find bar until the reader asks for the next match.
    this.#findNotFound = false;
    this.#toolbar.setMatches({ current: 0, total: 0, notFound: false });
    this.#fitFirstView = !reload && restore === undefined && this.#context.zoom() === 1;
    this.#queueView(restore);
    void this.#loadOutline(pdf, task);
    this.#focusIfIdle();
  }

  applyZoom(): void {
    this.#queueView();
  }

  fitWidthZoom(): number | undefined {
    const page = this.#viewer.getPageView(this.#viewer.currentPageNumber - 1) as { width: number; scale: number } | undefined;
    return page ? fitWidthZoom(this.#scroller.clientWidth, page.width / page.scale, PAGE_GUTTER) : undefined;
  }

  #queueView(location?: PdfViewState): void {
    this.#pendingView = { location: location ?? this.#pendingView?.location };
    this.#applyPendingView();
  }

  #applyPendingView(): void {
    const pending = this.#pendingView;
    if (!pending || !this.#viewer.pdfDocument || this.#viewer.pagesCount === 0 || !this.#scroller.offsetParent) {
      return;
    }
    this.#pendingView = undefined;
    // PDFViewer keeps the position on the current page while it rescales.
    this.#viewer.currentScale = this.#context.zoom();
    if (pending.location) {
      const { page, left, top } = pending.location;
      const pageNumber = Math.min(page, this.#viewer.pagesCount);
      const exact = pageNumber === page && left !== undefined && top !== undefined;
      this.#viewer.scrollPageIntoView({
        pageNumber,
        destArray: exact ? [null, { name: 'XYZ' }, left, top, null] : undefined,
        allowNegativeOffset: true,
      });
    }
    if (this.#fitFirstView) {
      this.#fitFirstView = false;
      const fit = this.fitWidthZoom();
      if (fit !== undefined && fit < 1) {
        this.#context.setZoom(fit);
      }
    }
  }

  #onWheel(event: WheelEvent): void {
    if (this.#viewer.scrollMode !== ScrollMode.PAGE || event.ctrlKey || event.metaKey || event.deltaY === 0) {
      return;
    }
    if (event.timeStamp - this.#lastWheel > WHEEL_GESTURE_GAP_MS) {
      this.#turnedInGesture = false;
    }
    this.#lastWheel = event.timeStamp;
    // One flick of a trackpad sends dozens of events; it may turn one page, not run through the document.
    if (!this.#turnedInGesture && this.#turnPageAtEdge(event.deltaY > 0)) {
      this.#turnedInGesture = true;
      event.preventDefault();
    }
  }

  #onKey(event: KeyboardEvent): void {
    if (this.#viewer.scrollMode !== ScrollMode.PAGE || event.altKey || event.ctrlKey || event.metaKey) {
      return;
    }
    const forward = event.key === 'PageDown' || event.key === 'ArrowDown';
    if ((forward || event.key === 'PageUp' || event.key === 'ArrowUp') && this.#turnPageAtEdge(forward)) {
      event.preventDefault();
    }
  }

  /** In single-page mode, turns to the next (or previous) page when the view is already at the page's end (or start). */
  #turnPageAtEdge(forward: boolean): boolean {
    const { scrollTop, scrollHeight, clientHeight } = this.#scroller;
    // The page's own edge is in view once only the gutter around it is left to scroll.
    const atEdge = forward ? scrollTop + clientHeight >= scrollHeight - PAGE_GUTTER - 1 : scrollTop <= PAGE_GUTTER + 1;
    if (!atEdge || !(forward ? this.#viewer.nextPage() : this.#viewer.previousPage())) {
      return false;
    }
    if (!forward) {
      // Going back lands at the end of the previous page, as continuous scrolling would.
      this.#scroller.scrollTop = this.#scroller.scrollHeight;
    }
    return true;
  }

  #track(location: ViewArea | undefined): void {
    if (location && Number.isFinite(location.pageNumber)) {
      this.#location = { page: location.pageNumber, left: location.left, top: location.top };
      this.#scheduleSave();
    }
  }

  #scheduleSave(): void {
    clearTimeout(this.#saveTimer);
    this.#saveTimer = setTimeout(() => {
      const state: PdfViewState = { ...(this.#location ?? { page: 1 }), outline: this.#outlineOpen };
      this.#context.saveViewerState(state);
    }, SAVE_DELAY_MS);
  }

  #find(query: string, again: boolean, previous: boolean): void {
    this.#moved = true;
    this.#eventBus.dispatch('find', {
      source: this,
      type: again ? 'again' : '',
      query,
      caseSensitive: false,
      entireWord: false,
      highlightAll: true,
      findPrevious: previous,
      matchDiacritics: false,
    });
  }

  async #loadOutline(pdf: PDFDocumentProxy, task: PDFDocumentLoadingTask): Promise<void> {
    let outline: OutlineNode[] = [];
    try {
      outline = (await pdf.getOutline()) ?? [];
    } catch {
      // An unreadable outline only means there is no outline panel.
    }
    if (task !== this.#task) {
      return;
    }
    const entries = flattenOutline(outline);
    this.#hasOutline = entries.length > 0;
    renderOutline(this.#outlinePanel, entries, (node) => {
      this.#moved = true;
      if (node.dest) {
        void this.#linkService.goToDestination(node.dest);
      } else if (node.url) {
        this.#context.openLink(node.url);
      }
    });
    this.#updateOutline();
  }

  #updateOutline(): void {
    const open = this.#hasOutline && this.#outlineOpen;
    this.#outlinePanel.hidden = !open;
    this.#layout.classList.toggle('with-outline', open);
    this.#toolbar.setOutline(this.#hasOutline, open);
  }

  /** Puts the keyboard focus on the pages (so arrow keys scroll them) unless the reader is using another control. */
  #focusIfIdle(): void {
    const active = document.activeElement;
    if (document.hasFocus() && (!active || active === document.body || active === this.#context.content)) {
      this.#scroller.focus({ preventScroll: true });
    }
  }
}
