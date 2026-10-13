// Builds the DOM for a laid-out document: one white paper column with the main text, then labelled sections for notes,
// comments, headers/footers and text boxes. All document text goes in as text nodes, never as HTML.
import { element } from '../shell/dom.ts';
import type { Block, DocumentLayout, Paragraph } from './layout.ts';
import type { StoryKind } from './wordDocument.ts';

const STORY_LABELS: Record<Exclude<StoryKind, 'main'>, string> = {
  footnotes: 'Footnotes',
  endnotes: 'Endnotes',
  comments: 'Comments',
  headers: 'Headers and footers',
  textBoxes: 'Text boxes',
  headerTextBoxes: 'Text boxes in headers and footers',
};

/** Paper width (A4 at 96 dpi) and the gap around it, as in doc.css. */
export const PAPER_WIDTH = 794;
export const PAPER_GUTTER = 24;

export function renderDocument(layout: DocumentLayout): HTMLElement {
  const paper = element('article', { className: 'doc-paper', attributes: { id: 'doc-paper' } });
  for (const story of layout.stories) {
    const target = story.kind === 'main' ? paper : element('section', { className: 'doc-story', attributes: { 'data-story': story.kind } });
    if (story.kind !== 'main') {
      target.append(element('h2', { className: 'doc-story-label', text: STORY_LABELS[story.kind] }));
      paper.append(target);
    }
    target.append(...story.blocks.map(renderBlock));
  }
  if (!paper.textContent?.trim()) {
    paper.replaceChildren(element('p', { className: 'doc-empty', text: 'This document has no text.' }));
  }
  const canvas = element('div', { className: 'doc-canvas' });
  canvas.append(paper);
  return canvas;
}

function renderBlock(block: Block): HTMLElement {
  switch (block.type) {
    case 'paragraph':
      return renderParagraph(block);
    case 'pageBreak':
      return element('div', { className: 'doc-page-break', attributes: { role: 'separator', 'aria-label': 'Page break' } });
    case 'table': {
      const body = element('tbody');
      for (const row of block.rows) {
        const tr = element('tr');
        for (const cell of row) {
          const td = element('td');
          td.append(...cell.map(renderParagraph));
          tr.append(td);
        }
        body.append(tr);
      }
      const table = element('table', { className: 'doc-table' });
      table.append(body);
      const wrapper = element('div', { className: 'doc-table-wrap' });
      wrapper.append(table);
      return wrapper;
    }
  }
}

function renderParagraph(paragraph: Paragraph): HTMLElement {
  const p = element('p', { className: 'doc-p' });
  for (const run of paragraph.runs) {
    if (run.noteRef) {
      p.append(element('sup', { className: 'doc-note-ref', text: run.text }));
    } else if (run.href) {
      // The shell opens http(s) links through the host; layout only keeps http(s) targets.
      p.append(element('a', { className: 'doc-link', text: run.text, attributes: { href: run.href, title: run.href } }));
    } else {
      p.append(run.text);
    }
  }
  return p;
}
