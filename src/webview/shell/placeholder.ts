import type { Viewer } from './app.ts';
import { element, formatBytes } from './dom.ts';

/** Stands in for a format whose viewer is not built yet: shows what was received, proving the pipeline end to end. */
export function placeholderViewer(kind: string): Viewer {
  return {
    render(document, context) {
      const card = element('section', { className: 'placeholder', attributes: { id: 'placeholder' } });
      card.append(
        element('h1', { text: document.fileName }),
        element('p', { className: 'dim', attributes: { id: 'placeholder-size' }, text: `${kind} · ${formatBytes(document.data.byteLength)}` }),
        element('p', { text: 'Document Lens opens this file type, but its viewer is not built yet.' }),
      );
      context.content.replaceChildren(card);
    },
  };
}
