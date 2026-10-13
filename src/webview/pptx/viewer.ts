// The presentation viewer: a slide list with lazily drawn thumbnails, the current slide fitted to the window, slide
// navigation (toolbar, ←/→, PageUp/PageDown, Home/End) and zoom. The current slide survives reloads of the file.
import type { LoadedDocument, Viewer, ViewerContext } from '../shell/app.ts';
import { element } from '../shell/dom.ts';
import { clampZoom } from '../shell/zoom.ts';
import { Deck } from './deck.ts';
import { MediaStore } from './media.ts';
import { renderSlide } from './slide.ts';
import { EMU_PER_PX } from './units.ts';

interface SavedState {
  slide?: number;
  /** The zoom follows the window size (set when the slide was last fitted). */
  fit?: boolean;
}

const STAGE_PADDING = 24;
const THUMB_WIDTH = 168;
const THUMB_MARGIN = '300px';

export function createPptxViewer(): Viewer {
  let ui: PresentationView | undefined;
  return {
    async render(document: LoadedDocument, context: ViewerContext) {
      const deck = await Deck.open(document.data);
      ui ??= new PresentationView(context);
      await ui.show(deck);
    },
    applyZoom(zoom) {
      ui?.applyZoom(zoom);
    },
    fitWidthZoom() {
      return ui?.fitZoom();
    },
  };
}

class PresentationView {
  private readonly context: ViewerContext;
  private readonly list: HTMLElement;
  private readonly stage: HTMLElement;
  private readonly frame: HTMLElement;
  private readonly counter: HTMLElement;
  private readonly hiddenTag: HTMLElement;
  private readonly previous: HTMLButtonElement;
  private readonly next: HTMLButtonElement;
  private deck: Deck | undefined;
  private media: MediaStore | undefined;
  private index = 0;
  private zoom: number;
  /** The zoom follows the stage size until the user picks another zoom. */
  private fitting = false;
  /** Set while the viewer applies a fitted zoom itself, so applyZoom does not take it for the user's choice. */
  private applyingFit = false;
  private lastFit: number | undefined;
  private renderToken = 0;
  private hidden: boolean[] = [];
  private truncated = false;
  private observer: IntersectionObserver | undefined;
  private thumbnails: HTMLElement[] = [];

  constructor(context: ViewerContext) {
    this.context = context;
    this.zoom = context.zoom();
    this.previous = element('button', { className: 'icon-button', text: '‹', attributes: { id: 'pptx-previous', title: 'Previous slide (←)', 'aria-label': 'Previous slide' } });
    this.counter = element('span', { className: 'pptx-counter', attributes: { id: 'pptx-counter', 'aria-live': 'polite' } });
    this.next = element('button', { className: 'icon-button', text: '›', attributes: { id: 'pptx-next', title: 'Next slide (→)', 'aria-label': 'Next slide' } });
    this.hiddenTag = element('span', { className: 'pptx-hidden-tag', text: 'Hidden', attributes: { id: 'pptx-hidden-tag', title: 'This slide is hidden in slide shows' } });
    this.hiddenTag.hidden = true;
    context.tools.replaceChildren(this.previous, this.counter, this.next, this.hiddenTag);

    this.list = element('nav', { className: 'pptx-slide-list', attributes: { id: 'pptx-slide-list', 'aria-label': 'Slides' } });
    this.frame = element('div', { className: 'pptx-frame', attributes: { id: 'pptx-frame' } });
    const inner = element('div', { className: 'pptx-stage-inner' });
    inner.append(this.frame);
    this.stage = element('div', { className: 'pptx-stage', attributes: { id: 'pptx-stage' } });
    this.stage.append(inner);
    const root = element('div', { className: 'pptx-viewer', attributes: { id: 'pptx-viewer' } });
    root.append(this.list, this.stage);
    context.content.replaceChildren(root);

    this.previous.addEventListener('click', () => this.go(this.index - 1));
    this.next.addEventListener('click', () => this.go(this.index + 1));
    this.list.addEventListener('click', (event) => {
      const thumb = (event.target as Element | null)?.closest?.('[data-index]');
      if (thumb) {
        this.go(Number(thumb.getAttribute('data-index')));
      }
    });
    this.frame.addEventListener('click', (event) => this.followSlideLink(event));
    window.addEventListener('keydown', (event) => this.onKey(event));
    if (typeof ResizeObserver !== 'undefined') {
      new ResizeObserver(() => this.refit()).observe(this.stage);
    }
    const saved = context.viewerState<SavedState>();
    this.index = Math.max(0, Math.trunc(saved?.slide ?? 0));
    this.fitting = saved?.fit ?? saved === undefined;
  }

  async show(deck: Deck): Promise<void> {
    const previousMedia = this.media;
    this.deck = deck;
    this.media = new MediaStore(deck.pkg);
    this.index = Math.min(this.index, Math.max(0, deck.slideParts.length - 1));
    this.hidden = deck.slideParts.map(() => false);
    this.buildList(deck);
    await this.renderCurrent();
    previousMedia?.dispose();
    if (this.fitting) {
      this.fit();
    }
    void this.scanHidden(deck);
  }

  /** Zoom at which the whole slide fits the stage. */
  fitZoom(): number | undefined {
    if (!this.deck) {
      return undefined;
    }
    const width = this.stage.clientWidth - STAGE_PADDING * 2;
    const height = this.stage.clientHeight - STAGE_PADDING * 2;
    if (width <= 0 || height <= 0) {
      return undefined;
    }
    const fit = Math.min(width / (this.deck.width / EMU_PER_PX), height / (this.deck.height / EMU_PER_PX));
    this.lastFit = fit;
    return fit;
  }

  applyZoom(zoom: number): void {
    if (Math.abs(zoom - this.zoom) > 0.0001 && !this.applyingFit) {
      // The user zoomed; the Fit button lands here too, right after fitZoom() measured the stage.
      this.fitting = this.lastFit !== undefined && Math.abs(zoom - clampZoom(this.lastFit)) < 0.001;
    }
    this.zoom = zoom;
    this.layoutFrame();
    this.save();
  }

  private fit(): void {
    const fit = this.fitZoom();
    if (fit === undefined || Math.abs(clampZoom(fit) - this.zoom) < 0.001) {
      return;
    }
    this.applyingFit = true;
    try {
      this.context.setZoom(fit);
    } finally {
      this.applyingFit = false;
    }
  }

  private refit(): void {
    if (this.fitting) {
      requestAnimationFrame(() => {
        if (this.fitting) {
          this.fit();
        }
      });
    }
  }

  private layoutFrame(): void {
    if (!this.deck) {
      return;
    }
    const width = (this.deck.width / EMU_PER_PX) * this.zoom;
    const height = (this.deck.height / EMU_PER_PX) * this.zoom;
    this.frame.style.setProperty('width', `${width}px`);
    this.frame.style.setProperty('height', `${height}px`);
    const slide = this.frame.firstElementChild as HTMLElement | null;
    slide?.style.setProperty('transform', `scale(${this.zoom})`);
  }

  private save(): void {
    this.context.saveViewerState({ slide: this.index, fit: this.fitting } satisfies SavedState);
  }

  private go(index: number): void {
    const count = this.deck?.slideParts.length ?? 0;
    const target = Math.max(0, Math.min(count - 1, index));
    if (count === 0 || !Number.isInteger(target) || target === this.index) {
      return;
    }
    this.index = target;
    this.save();
    void this.renderCurrent();
  }

  private onKey(event: KeyboardEvent): void {
    const target = event.target as Element | null;
    if (event.metaKey || event.ctrlKey || event.altKey || target?.closest?.('input, textarea, select, [contenteditable="true"]')) {
      return;
    }
    const count = this.deck?.slideParts.length ?? 0;
    const inList = Boolean(target?.closest?.('.pptx-slide-list'));
    const moves: Record<string, number> = {
      ArrowLeft: this.index - 1,
      ArrowRight: this.index + 1,
      PageUp: this.index - 1,
      PageDown: this.index + 1,
      Home: 0,
      End: count - 1,
      ...(inList ? { ArrowUp: this.index - 1, ArrowDown: this.index + 1 } : {}),
    };
    const destination = moves[event.key];
    if (destination === undefined || count === 0) {
      return;
    }
    event.preventDefault();
    this.go(destination);
    if (inList) {
      this.thumbnails[this.index]?.focus();
    }
  }

  /** Links inside a slide that jump to another slide (the shell handles web links). */
  private followSlideLink(event: MouseEvent): void {
    const link = (event.target as Element | null)?.closest?.('[data-slide-index], [data-slide-jump]');
    if (!link) {
      return;
    }
    event.preventDefault();
    const count = this.deck?.slideParts.length ?? 0;
    const jump = link.getAttribute('data-slide-jump');
    const destinations: Record<string, number> = { nextslide: this.index + 1, previousslide: this.index - 1, firstslide: 0, lastslide: count - 1 };
    this.go(jump ? (destinations[jump] ?? this.index) : Number(link.getAttribute('data-slide-index')));
  }

  private buildList(deck: Deck): void {
    this.observer?.disconnect();
    const ratio = deck.height / deck.width;
    this.thumbnails = deck.slideParts.map((_part, index) => {
      const thumb = element('button', { className: 'pptx-thumb', attributes: { 'data-index': String(index), 'aria-label': `Slide ${index + 1}` } });
      const number = element('span', { className: 'pptx-thumb-number', text: String(index + 1) });
      const frame = element('span', { className: 'pptx-thumb-frame' });
      frame.setAttribute('inert', '');
      frame.style.setProperty('width', `${THUMB_WIDTH}px`);
      frame.style.setProperty('height', `${Math.round(THUMB_WIDTH * ratio)}px`);
      thumb.append(number, frame);
      return thumb;
    });
    this.list.replaceChildren(...this.thumbnails);
    if (typeof IntersectionObserver !== 'undefined') {
      this.observer = new IntersectionObserver((entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) {
            this.observer?.unobserve(entry.target);
            void this.renderThumbnail(deck, Number(entry.target.getAttribute('data-index')));
          }
        }
      }, { root: this.list, rootMargin: THUMB_MARGIN });
      for (const thumb of this.thumbnails) {
        this.observer.observe(thumb);
      }
    }
    if (deck.slideParts.length === 0) {
      this.list.append(element('p', { className: 'pptx-empty dim', text: 'No slides' }));
    }
  }

  private async renderThumbnail(deck: Deck, index: number): Promise<void> {
    const thumb = this.thumbnails[index];
    const media = this.media;
    if (!thumb || !media || deck !== this.deck) {
      return;
    }
    const frame = thumb.querySelector<HTMLElement>('.pptx-thumb-frame');
    try {
      const rendered = await renderSlide(deck, index, media);
      if (deck !== this.deck || !frame) {
        return;
      }
      rendered.element.style.setProperty('transform', `scale(${THUMB_WIDTH / (deck.width / EMU_PER_PX)})`);
      frame.replaceChildren(rendered.element);
      this.markHidden(index, rendered.hidden);
    } catch {
      frame?.replaceChildren(element('span', { className: 'pptx-thumb-error', text: '!' }));
    }
  }

  private async scanHidden(deck: Deck): Promise<void> {
    for (let index = 0; index < deck.slideParts.length && deck === this.deck; index++) {
      try {
        this.markHidden(index, await deck.isHidden(index));
      } catch {
        // The slide itself will report what is wrong with it when it is shown.
      }
    }
  }

  private markHidden(index: number, hidden: boolean): void {
    this.hidden[index] = hidden;
    this.thumbnails[index]?.classList.toggle('pptx-thumb-hidden', hidden);
    this.thumbnails[index]?.setAttribute('aria-label', `Slide ${index + 1}${hidden ? ' (hidden)' : ''}`);
    if (index === this.index) {
      this.hiddenTag.hidden = !hidden;
    }
  }

  private updateTools(): void {
    const count = this.deck?.slideParts.length ?? 0;
    this.counter.textContent = count === 0 ? '0 / 0' : `${this.index + 1} / ${count}`;
    this.previous.disabled = this.index <= 0;
    this.next.disabled = this.index >= count - 1;
    this.hiddenTag.hidden = !this.hidden[this.index];
    this.thumbnails.forEach((thumb, index) => {
      if (index === this.index) {
        thumb.setAttribute('aria-current', 'true');
        thumb.scrollIntoView?.({ block: 'nearest' });
      } else {
        thumb.removeAttribute('aria-current');
      }
    });
  }

  private async renderCurrent(): Promise<void> {
    const deck = this.deck;
    const media = this.media;
    if (!deck || !media) {
      return;
    }
    const token = ++this.renderToken;
    this.updateTools();
    if (deck.slideParts.length === 0) {
      this.frame.replaceChildren(element('p', { className: 'pptx-empty', text: 'This presentation has no slides.' }));
      this.updateNotice();
      return;
    }
    // Another slide may be chosen while this one renders; only the latest render is shown.
    const index = this.index;
    let content: HTMLElement;
    let truncated = false;
    try {
      const rendered = await renderSlide(deck, index, media);
      content = rendered.element;
      truncated = rendered.truncated;
      this.markHidden(index, rendered.hidden);
    } catch (error) {
      content = element('div', { className: 'pptx-slide pptx-slide-error', text: `Slide ${index + 1} could not be displayed: ${(error as Error).message}` });
      content.style.setProperty('width', `${deck.width / EMU_PER_PX}px`);
      content.style.setProperty('height', `${deck.height / EMU_PER_PX}px`);
    }
    if (token !== this.renderToken) {
      return;
    }
    this.truncated = truncated;
    content.id = 'pptx-current-slide';
    content.setAttribute('aria-label', `Slide ${index + 1}`);
    this.frame.replaceChildren(content);
    this.layoutFrame();
    this.updateNotice();
  }

  private updateNotice(): void {
    const notes: string[] = [];
    if (this.deck && this.deck.omittedSlides > 0) {
      notes.push(`Only the first ${this.deck.slideParts.length} slides are shown.`);
    }
    if (this.truncated) {
      notes.push('Some content on this slide could not be shown.');
    }
    this.context.showNotice(notes.length > 0 ? notes.join(' ') : null);
  }
}
