// Controls around the grid: the sheet tabs under it and the cell bar in the toolbar.
import type { SheetEntry } from '../model/types.ts';

export class SheetTabs {
  readonly element: HTMLElement;
  private readonly buttons = new Map<number, HTMLButtonElement>();

  /** Hidden sheets get no tab, as in Excel. */
  constructor(document: Document, sheets: readonly SheetEntry[], onSelect: (index: number) => void) {
    this.element = document.createElement('nav');
    this.element.className = 'dl-tabs';
    this.element.setAttribute('role', 'tablist');
    this.element.setAttribute('aria-label', 'Sheets');
    sheets.forEach((sheet, index) => {
      if (sheet.hidden) {
        return;
      }
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'dl-tab';
      button.setAttribute('role', 'tab');
      button.setAttribute('aria-selected', 'false');
      button.textContent = sheet.name;
      button.title = sheet.name;
      button.addEventListener('click', () => onSelect(index));
      this.buttons.set(index, button);
      this.element.append(button);
    });
  }

  setActive(index: number): void {
    for (const [i, button] of this.buttons) {
      button.setAttribute('aria-selected', String(i === index));
      button.tabIndex = i === index ? 0 : -1;
    }
    this.buttons.get(index)?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }
}

export class CellBar {
  readonly element: HTMLElement;
  private readonly address: HTMLElement;
  private readonly value: HTMLElement;

  constructor(document: Document) {
    this.element = document.createElement('div');
    this.element.className = 'dl-cellbar';
    this.element.setAttribute('aria-live', 'polite');
    this.address = document.createElement('span');
    this.address.className = 'dl-cellbar-address';
    this.address.setAttribute('aria-label', 'Selected cell');
    this.value = document.createElement('span');
    this.value.className = 'dl-cellbar-value';
    this.value.setAttribute('aria-label', 'Cell content');
    this.element.append(this.address, this.value);
    this.show(undefined);
  }

  show(info: { address: string; text: string } | undefined): void {
    this.element.hidden = !info;
    this.address.textContent = info?.address ?? '';
    // One line in the toolbar; the full text is in the tooltip.
    this.value.textContent = info?.text.replace(/[\r\n]+/g, ' ') ?? '';
    this.value.title = info?.text ?? '';
  }
}
