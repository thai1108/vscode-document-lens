// The PDF viewer's controls in the shell toolbar: outline toggle, page navigation and the find bar. It knows nothing
// about pdf.js; the viewer reacts to the actions and feeds back the page, page count and match count.
import { element } from '../shell/dom.ts';
import { parsePageNumber } from './navigation.ts';

export interface ToolbarActions {
  goToPage(page: number): void;
  /** `again` steps to the next (or `previous`) match of an unchanged query instead of starting a new search. */
  find(query: string, options: { again: boolean; previous: boolean }): void;
  closeFind(): void;
  toggleOutline(): void;
  /** Moves the keyboard focus back to the pages, e.g. after a jump or closing the find bar. */
  focusDocument(): void;
}

export interface MatchCount {
  /** 1-based index of the selected match; 0 while none is selected. */
  current: number;
  total: number;
  /** The whole document has been searched without a match. */
  notFound: boolean;
}

export function matchesLabel(query: string, { current, total, notFound }: MatchCount): string {
  if (!query) {
    return '';
  }
  if (total > 0 && current > 0) {
    return `${current} of ${total}`;
  }
  return notFound ? 'No results' : '';
}

const MOD = /Mac|iPhone|iPad/.test(globalThis.navigator?.userAgent ?? '') ? '⌘' : 'Ctrl+';

export class PdfToolbar {
  readonly #actions: ToolbarActions;
  readonly #outline: HTMLButtonElement;
  readonly #previous: HTMLButtonElement;
  readonly #next: HTMLButtonElement;
  readonly #pageInput: HTMLInputElement;
  readonly #pageCount: HTMLElement;
  readonly #findToggle: HTMLButtonElement;
  readonly #findBar: HTMLElement;
  readonly #findInput: HTMLInputElement;
  readonly #findCount: HTMLElement;
  #page = 1;
  #pages = 0;

  constructor(tools: HTMLElement, actions: ToolbarActions) {
    this.#actions = actions;
    this.#outline = iconButton('pdf-outline-toggle', 'Outline', 'outline');
    this.#outline.setAttribute('aria-pressed', 'false');
    this.#outline.hidden = true;
    this.#previous = iconButton('pdf-previous-page', 'Previous page', 'up');
    this.#next = iconButton('pdf-next-page', 'Next page', 'down');
    this.#pageInput = element('input', {
      className: 'pdf-input pdf-page-input',
      attributes: { id: 'pdf-page-number', type: 'text', inputmode: 'numeric', autocomplete: 'off', 'aria-label': 'Page number', title: 'Page number (Enter to go)' },
    });
    this.#pageCount = element('span', { className: 'pdf-page-count dim', attributes: { id: 'pdf-page-count' } });
    const navigation = element('div', { className: 'pdf-tool-group', attributes: { id: 'pdf-navigation' } });
    navigation.append(this.#outline, this.#previous, this.#next, this.#pageInput, this.#pageCount);

    this.#findToggle = iconButton('pdf-find-toggle', `Find (${MOD}F)`, 'search');
    this.#findInput = element('input', {
      className: 'pdf-input pdf-find-input',
      attributes: { id: 'pdf-find-input', type: 'text', placeholder: 'Find', autocomplete: 'off', spellcheck: 'false', 'aria-label': 'Find in document' },
    });
    this.#findCount = element('span', { className: 'pdf-find-count', attributes: { id: 'pdf-find-count', 'aria-live': 'polite' } });
    const findPrevious = iconButton('pdf-find-previous', 'Previous match (Shift+Enter)', 'up');
    const findNext = iconButton('pdf-find-next', 'Next match (Enter)', 'down');
    const findClose = iconButton('pdf-find-close', 'Close (Escape)', 'close');
    this.#findBar = element('div', { className: 'pdf-find-bar', attributes: { id: 'pdf-find-bar', role: 'search' } });
    this.#findBar.hidden = true;
    this.#findBar.append(this.#findInput, this.#findCount, findPrevious, findNext, findClose);
    const find = element('div', { className: 'pdf-tool-group' });
    find.append(this.#findToggle, this.#findBar);

    tools.replaceChildren(navigation, element('span', { className: 'pdf-separator' }), find);
    this.#render();

    this.#outline.addEventListener('click', () => actions.toggleOutline());
    this.#previous.addEventListener('click', () => this.#goTo(this.#page - 1));
    this.#next.addEventListener('click', () => this.#goTo(this.#page + 1));
    this.#pageInput.addEventListener('keydown', this.#onPageKey);
    this.#pageInput.addEventListener('blur', () => this.#render());
    this.#findToggle.addEventListener('click', () => this.openFind());
    this.#findInput.addEventListener('input', (event) => {
      // An IME sends input events for every keystroke of a composition; search once it is committed.
      if (!(event as InputEvent).isComposing) {
        this.#search();
      }
    });
    this.#findInput.addEventListener('compositionend', () => this.#search());
    this.#findInput.addEventListener('keydown', this.#onFindKey);
    findPrevious.addEventListener('click', () => this.#step(true));
    findNext.addEventListener('click', () => this.#step(false));
    findClose.addEventListener('click', () => this.closeFind());
    window.addEventListener('keydown', this.#onKey);
  }

  setPageCount(count: number): void {
    this.#pages = count;
    this.#page = Math.min(Math.max(this.#page, 1), Math.max(count, 1));
    this.#render();
  }

  setPage(page: number): void {
    this.#page = page;
    this.#render();
  }

  setMatches(matches: MatchCount): void {
    this.#findCount.textContent = matchesLabel(this.#findInput.value, matches);
    this.#findCount.classList.toggle('no-results', matches.notFound && matches.total === 0 && this.#findInput.value !== '');
  }

  setOutline(available: boolean, open: boolean): void {
    this.#outline.hidden = !available;
    this.#outline.setAttribute('aria-pressed', String(available && open));
  }

  /** Shows the find bar and focuses its field, re-running the last search so its matches are highlighted again. */
  openFind(): void {
    const wasHidden = this.#findBar.hidden;
    this.#findBar.hidden = false;
    this.#findToggle.hidden = true;
    this.#findInput.focus();
    this.#findInput.select();
    if (wasHidden && this.#findInput.value) {
      this.#search();
    }
  }

  closeFind(): void {
    if (this.#findBar.hidden) {
      return;
    }
    this.#findBar.hidden = true;
    this.#findToggle.hidden = false;
    this.setMatches({ current: 0, total: 0, notFound: false });
    this.#actions.closeFind();
    this.#actions.focusDocument();
  }

  dispose(): void {
    window.removeEventListener('keydown', this.#onKey);
  }

  /** `resetField` also overwrites what the reader is typing in the page number field. */
  #render(resetField = false): void {
    const known = this.#pages > 0;
    this.#pageCount.textContent = `of ${this.#pages}`;
    if (resetField || this.#pageInput.ownerDocument.activeElement !== this.#pageInput) {
      this.#pageInput.value = known ? String(this.#page) : '';
    }
    this.#pageInput.size = Math.max(2, String(this.#pages).length);
    this.#pageInput.disabled = !known;
    this.#previous.disabled = !known || this.#page <= 1;
    this.#next.disabled = !known || this.#page >= this.#pages;
  }

  #goTo(page: number): void {
    if (this.#pages > 0) {
      this.#actions.goToPage(Math.min(Math.max(page, 1), this.#pages));
    }
  }

  #search(): void {
    this.#actions.find(this.#findInput.value, { again: false, previous: false });
  }

  #step(previous: boolean): void {
    if (this.#findInput.value) {
      this.#actions.find(this.#findInput.value, { again: true, previous });
    }
  }

  readonly #onPageKey = (event: KeyboardEvent) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      const page = parsePageNumber(this.#pageInput.value, this.#pages);
      this.#actions.focusDocument();
      if (page !== null) {
        this.#goTo(page);
      }
      this.#render(true);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      this.#actions.focusDocument();
      this.#render(true);
    }
  };

  readonly #onFindKey = (event: KeyboardEvent) => {
    if (event.key === 'Enter' && !event.isComposing) {
      event.preventDefault();
      this.#step(event.shiftKey);
    }
  };

  readonly #onKey = (event: KeyboardEvent) => {
    const modifier = (event.metaKey || event.ctrlKey) && !event.altKey;
    const key = event.key.toLowerCase();
    if (modifier && key === 'f') {
      event.preventDefault();
      this.openFind();
    } else if ((modifier && key === 'g') || event.key === 'F3') {
      event.preventDefault();
      if (this.#findInput.value) {
        if (this.#findBar.hidden) {
          this.openFind();
        }
        this.#step(event.shiftKey);
      }
    } else if (event.key === 'Escape' && !event.isComposing && !this.#findBar.hidden) {
      event.preventDefault();
      this.closeFind();
    }
  };
}

type IconName = 'outline' | 'up' | 'down' | 'search' | 'close';

const ICON_PATHS: Record<IconName, string> = {
  outline: 'M2.5 4h1M6 4h7.5M2.5 8h1M6 8h7.5M2.5 12h1M6 12h7.5',
  up: 'M3.5 10 8 5.5l4.5 4.5',
  down: 'M3.5 6 8 10.5 12.5 6',
  search: 'M10.2 10.2 14 14M11.5 6.75a4.75 4.75 0 1 1-9.5 0 4.75 4.75 0 0 1 9.5 0Z',
  close: 'm4 4 8 8m0-8-8 8',
};

function iconButton(id: string, label: string, icon: IconName): HTMLButtonElement {
  const button = element('button', { className: 'icon-button pdf-icon-button', attributes: { id, type: 'button', title: label, 'aria-label': label } });
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('aria-hidden', 'true');
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', ICON_PATHS[icon]);
  svg.append(path);
  button.append(svg);
  return button;
}
