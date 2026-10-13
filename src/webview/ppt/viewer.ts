// Shows a .ppt presentation as a scrolling list of slide cards: each slide's title, its other text as paragraphs or
// bullets, and its pictures (as blob: URLs). All text goes into the DOM as textContent.
import type { Viewer } from '../shell/app.ts';
import { element } from '../shell/dom.ts';
import type { Picture } from './pictures.ts';
import { parsePresentation, type Presentation, type Slide } from './presentation.ts';
import type { Paragraph, TextBlock } from './text.ts';

const SIMPLIFIED_DETAIL = 'only the text and pictures of each slide are shown, without layout, shapes or formatting.';

export function pptViewer(): Viewer {
  let pictureUrls: string[] = [];
  return {
    render(document, context, reload) {
      // Parsing first: a file that fails keeps the previous slides on screen under the error banner.
      const presentation = parsePresentation(document.data);
      // One URL per picture, however many slides show it.
      const urls = new Map<Uint8Array, string>();
      const urlOf: UrlOf = (data, mime) => {
        let url = urls.get(data);
        if (!url) {
          // Picture bytes live in an ordinary ArrayBuffer (the file's, or a converted copy), never in shared memory.
          url = URL.createObjectURL(new Blob([data as Uint8Array<ArrayBuffer>], { type: mime }));
          urls.set(data, url);
        }
        return url;
      };
      const scrollTop = reload ? context.content.scrollTop : 0;
      context.content.replaceChildren(renderDeck(presentation, urlOf));
      context.content.scrollTop = scrollTop;
      context.tools.replaceChildren(element('span', { className: 'dim', attributes: { id: 'ppt-slide-count' }, text: describeCount(presentation) }));
      context.showSimplifiedNotice([SIMPLIFIED_DETAIL, ...presentation.warnings].join(' '));
      pictureUrls.forEach((url) => URL.revokeObjectURL(url));
      pictureUrls = [...urls.values()];
    },
  };
}

function describeCount(presentation: Presentation): string {
  const count = presentation.slides.length;
  const hidden = presentation.slides.filter((slide) => slide.hidden).length;
  return `${count} ${count === 1 ? 'slide' : 'slides'}${hidden ? ` · ${hidden} hidden` : ''}`;
}

type UrlOf = (data: Uint8Array, mime: string) => string;

function renderDeck(presentation: Presentation, urlOf: UrlOf): HTMLElement {
  const deck = element('div', { className: 'ppt-deck', attributes: { id: 'ppt-deck' } });
  deck.style.setProperty('--ppt-aspect', `${presentation.width} / ${presentation.height}`);
  if (!presentation.slides.length) {
    deck.append(element('p', { className: 'ppt-no-slides', text: 'This presentation has no slides.' }));
  }
  deck.append(...presentation.slides.map((slide) => renderSlide(slide, urlOf)));
  return deck;
}

function renderSlide(slide: Slide, urlOf: UrlOf): HTMLElement {
  const name = `Slide ${slide.number}`;
  const section = element('section', {
    className: `ppt-slide${slide.hidden ? ' ppt-slide-hidden' : ''}`,
    attributes: { id: `slide-${slide.number}`, 'aria-label': slide.hidden ? `${name} (hidden)` : name },
  });
  const label = element('header', { className: 'ppt-slide-label' });
  label.append(element('span', { className: 'ppt-slide-number', text: name }));
  if (slide.hidden) {
    label.append(element('span', { className: 'ppt-badge', text: 'Hidden' }));
  }
  const card = element('div', { className: 'ppt-slide-card' });
  card.append(...slide.texts.map(renderText));
  if (slide.pictures.length) {
    const gallery = element('div', { className: 'ppt-pictures' });
    gallery.append(...slide.pictures.map((picture, i) => renderPicture(picture, `Picture ${i + 1} on ${name}`, urlOf)));
    card.append(gallery);
  }
  if (slide.damaged) {
    card.append(element('p', { className: 'ppt-note', text: 'Part of this slide could not be read.' }));
  } else if (!slide.texts.length && !slide.pictures.length) {
    card.append(element('p', { className: 'ppt-note', text: 'No text or pictures on this slide.' }));
  }
  section.append(label, card);
  return section;
}

function renderText(block: TextBlock): HTMLElement {
  if (block.kind === 'title') {
    return element('h2', { className: 'ppt-title', text: block.paragraphs.map((paragraph) => paragraph.text).join('\n') });
  }
  const container = element('div', { className: `ppt-text ppt-text-${block.kind}` });
  let list: HTMLUListElement | undefined;
  for (const paragraph of block.paragraphs) {
    if (paragraph.bullet) {
      list ??= container.appendChild(element('ul'));
      list.append(paragraphElement('li', paragraph));
    } else {
      list = undefined;
      container.append(paragraphElement('p', paragraph));
    }
  }
  return container;
}

function paragraphElement(tag: 'li' | 'p', paragraph: Paragraph): HTMLElement {
  return element(tag, { className: `ppt-level-${paragraph.level}`, text: paragraph.text });
}

function renderPicture(picture: Picture, alt: string, urlOf: UrlOf): HTMLElement {
  const figure = element('figure', { className: 'ppt-picture' });
  const showPlaceholder = (text: string) => {
    figure.classList.add('ppt-picture-placeholder');
    figure.replaceChildren(element('span', { text }));
  };
  if (!picture.data || !picture.mime) {
    showPlaceholder(picture.note ?? 'Picture not shown');
    return figure;
  }
  const image = element('img', {
    attributes: { src: urlOf(picture.data, picture.mime), alt, loading: 'lazy', decoding: 'async' },
  });
  if (picture.width && picture.height) {
    image.width = picture.width;
    image.height = picture.height;
  }
  image.addEventListener('error', () => showPlaceholder('Picture could not be decoded'), { once: true });
  figure.append(image);
  return figure;
}
