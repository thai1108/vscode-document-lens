import type { Viewer } from '../shell/app.ts';
import { LAYOUT_LIMITS, layoutDocument, type DocumentLayout } from './layout.ts';
import { PAPER_GUTTER, PAPER_WIDTH, renderDocument } from './render.ts';
import { DOC_LIMITS, readWordDocument } from './wordDocument.ts';

const NOTICE = 'only the text of this document is shown; formatting, pictures and page layout are left out.';
const TRUNCATED = ' It is very long, so only its beginning is shown.';
/** Room for the content area's vertical scrollbar when fitting the paper to the width. */
const SCROLLBAR_ALLOWANCE = 16;

/** Text-level preview of Word 97–2003 .doc files, parsed in the webview. */
export function docViewer(limits = { ...DOC_LIMITS, ...LAYOUT_LIMITS }): Viewer {
  let content: HTMLElement | undefined;
  return {
    render(document, context, reload) {
      content = context.content;
      const scrollTop = content.scrollTop;
      let layout: DocumentLayout;
      try {
        layout = layoutDocument(readWordDocument(document.data, limits), limits);
      } catch (error) {
        // Leave no stale text or notice under the error banner.
        content.replaceChildren();
        context.showNotice(null);
        throw error;
      }
      content.replaceChildren(renderDocument(layout));
      if (reload) {
        content.scrollTop = scrollTop;
      }
      context.showSimplifiedNotice(NOTICE + (layout.truncated ? TRUNCATED : ''));
    },
    fitWidthZoom() {
      const width = content?.getBoundingClientRect().width;
      return width ? (width - SCROLLBAR_ALLOWANCE) / (PAPER_WIDTH + 2 * PAPER_GUTTER) : undefined;
    },
  };
}
