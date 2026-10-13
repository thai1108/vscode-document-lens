// Page and position bookkeeping that does not need pdf.js: what is saved in the webview state, typed page numbers and
// the fit-width zoom.

/** Where the reader is, saved in the webview state so reopening or reloading the file returns there. */
export interface PdfViewState {
  /** 1-based page at the top of the view. */
  page: number;
  /** PDF coordinates (points) of the view's top-left corner on that page, so they hold at any zoom. */
  left?: number;
  top?: number;
  /** Whether the outline panel is open. */
  outline?: boolean;
}

/** The saved view state if `raw` (from a previous session, so possibly stale or foreign) is well-formed. */
export function parseViewState(raw: unknown): PdfViewState | undefined {
  if (typeof raw !== 'object' || raw === null) {
    return undefined;
  }
  const { page, left, top, outline } = raw as Record<string, unknown>;
  if (typeof page !== 'number' || !Number.isSafeInteger(page) || page < 1) {
    return undefined;
  }
  const state: PdfViewState = { page };
  if (isFiniteNumber(left) && isFiniteNumber(top)) {
    state.left = left;
    state.top = top;
  }
  if (typeof outline === 'boolean') {
    state.outline = outline;
  }
  return state;
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

/** The page a typed page number points at, clamped to the document; null unless it is a plain number. */
export function parsePageNumber(text: string, pageCount: number): number | null {
  const trimmed = text.trim();
  if (!/^\d{1,9}$/.test(trimmed) || pageCount < 1) {
    return null;
  }
  return Math.min(pageCount, Math.max(1, Number(trimmed)));
}

/**
 * Zoom at which a page `pageWidth` CSS pixels wide at 100% fills a view `viewWidth` wide, leaving `gutter` pixels on
 * each side; undefined when there is no room or no page.
 */
export function fitWidthZoom(viewWidth: number, pageWidth: number, gutter: number): number | undefined {
  const room = viewWidth - 2 * gutter;
  return pageWidth > 0 && room > 0 ? room / pageWidth : undefined;
}
