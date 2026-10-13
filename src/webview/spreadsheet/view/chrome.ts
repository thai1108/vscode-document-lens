// Controls around the grid: the sheet tabs under it, the cell bar and, for .csv/.tsv files, the CSV bar in the toolbar.
import { delimiterLabel, lineEndingLabel } from '../../../csv/dialect.ts';
import { encodingLabel } from '../../../csv/encoding.ts';
import type { CsvDetails } from '../csv/workbook.ts';
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

/** For .csv and .tsv files: how the file is written, the header row switch and the way back to the text editor. */
export class CsvBar {
  readonly element: HTMLElement;

  constructor(document: Document, details: CsvDetails, actions: { toggleHeaderRow(): void; openAsText(): void }) {
    this.element = document.createElement('div');
    this.element.className = 'dl-csvbar';
    const info = document.createElement('span');
    info.className = 'dl-csv-info';
    info.id = 'dl-csv-info';
    const lineBreaks = `${lineEndingLabel(details.lineEnding)}${details.mixedLineEndings ? ' (mixed)' : ''}`;
    info.textContent = [delimiterLabel(details.delimiter), encodingLabel(details.encoding), lineBreaks].join(' · ');
    info.title = `Delimiter: ${delimiterLabel(details.delimiter)}\nEncoding: ${encodingLabel(details.encoding)}\nLine breaks: ${lineBreaks}`;
    const header = button(document, 'dl-csv-header', 'Header row', 'Show the first row as column names, kept at the top');
    header.setAttribute('aria-pressed', String(details.headerRow));
    header.addEventListener('click', () => actions.toggleHeaderRow());
    const text = button(document, 'dl-open-as-text', 'Open as text', 'Reopen this file in the text editor');
    text.addEventListener('click', () => actions.openAsText());
    this.element.append(info, header, text);
  }
}

function button(document: Document, id: string, label: string, title: string): HTMLButtonElement {
  const element = document.createElement('button');
  element.type = 'button';
  element.id = id;
  element.className = 'text-button';
  element.textContent = label;
  element.title = title;
  return element;
}
