// The spreadsheet viewer (.xlsx and .xls): reads the workbook, shows one sheet at a time with tabs to switch, and
// remembers the sheet, scroll position and selected cell in the webview state.
import type { LoadedDocument, Viewer, ViewerContext } from '../shell/app.ts';
import { parseCellRef } from './model/address.ts';
import type { Sheet, Workbook } from './model/types.ts';
import { CellBar, SheetTabs } from './view/chrome.ts';
import { Grid, type GridHost, type GridPosition } from './view/grid.ts';
import { StyleRules } from './view/styleRules.ts';
import { openXls } from './xls/workbook.ts';
import { openXlsx } from './xlsx/workbook.ts';
import { isZip } from './xlsx/zip.ts';

/** Sheets up to this many cells stay parsed after switching away from them. */
const CACHED_SHEET_CELLS = 200_000;
const SAVE_DELAY_MS = 250;
const MAX_SAVED_POSITIONS = 1000;
const SCOPE = 'dl-sheet';

export interface SavedState {
  sheet?: string;
  positions?: Record<string, GridPosition>;
}

export function spreadsheetViewer(): Viewer {
  let session: Session | undefined;
  return {
    async render(document: LoadedDocument, context: ViewerContext) {
      const previous = session;
      const state = previous?.state() ?? readSavedState(context.viewerState());
      if (!previous) {
        context.content.replaceChildren(message(context.content.ownerDocument, 'Loading…'));
      }
      let next: Session | undefined;
      try {
        next = new Session(await openWorkbook(document.data), context, state, context.zoom());
        await next.start();
      } catch (error) {
        next?.dispose();
        if (!previous) {
          context.content.replaceChildren();
        }
        throw error;
      }
      // A reloaded file replaces the previous one only once it is ready.
      previous?.dispose();
      session = next;
    },
    applyZoom(zoom: number) {
      session?.setZoom(zoom);
    },
  };
}

export async function openWorkbook(bytes: Uint8Array): Promise<Workbook> {
  if (isZip(bytes)) {
    return openXlsx(bytes);
  }
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
    return openXls(bytes);
  }
  throw new Error('Not a spreadsheet file');
}

class Session {
  private readonly workbook: Workbook;
  private readonly context: ViewerContext;
  private readonly root: HTMLElement;
  private readonly gridHost: HTMLElement;
  private readonly tabs: SheetTabs;
  private readonly cellBar: CellBar;
  private readonly rules: StyleRules;
  private readonly cache = new Map<number, Sheet>();
  private readonly positions = new Map<string, GridPosition>();
  private readonly initialSheet: string | undefined;
  private zoom: number;
  private grid: Grid | undefined;
  private sheet: Sheet | undefined;
  private current = -1;
  private loading: AbortController | undefined;
  private saveTimer: ReturnType<typeof setTimeout> | undefined;
  private disposed = false;

  constructor(workbook: Workbook, context: ViewerContext, state: SavedState, zoom: number) {
    this.workbook = workbook;
    this.context = context;
    this.zoom = zoom;
    this.initialSheet = state.sheet;
    for (const [name, position] of Object.entries(state.positions ?? {})) {
      this.positions.set(name, position);
    }
    const document = context.content.ownerDocument;
    this.rules = new StyleRules(document, context.styleNonce, workbook.formats, `.${SCOPE}`);
    this.root = document.createElement('div');
    this.root.className = SCOPE;
    this.gridHost = document.createElement('div');
    this.gridHost.className = 'dl-grid-host';
    this.tabs = new SheetTabs(document, workbook.sheets, (index) => void this.show(index));
    this.root.append(this.gridHost, this.tabs.element);
    this.cellBar = new CellBar(document);
    document.addEventListener('copy', this.onCopy);
  }

  /** Reads the first sheet and puts the viewer on the page; a sheet that cannot be read fails the whole render. */
  async start(): Promise<void> {
    const saved = this.workbook.sheets.findIndex((sheet) => sheet.name === this.initialSheet && !sheet.hidden);
    const index = saved >= 0 ? saved : this.workbook.activeSheet;
    const sheet = await this.workbook.loadSheet(index);
    if (this.disposed) {
      return;
    }
    this.keep(index, sheet);
    this.mount(index, sheet);
    this.context.content.replaceChildren(this.root);
    this.context.tools.replaceChildren(this.cellBar.element);
    this.activate();
  }

  /** Switches to sheet `index`; a sheet that cannot be read shows its error in place of the grid. */
  async show(index: number): Promise<void> {
    if (index === this.current && this.grid) {
      return;
    }
    this.loading?.abort();
    const loading = new AbortController();
    this.loading = loading;
    this.rememberPosition();
    let sheet = this.cache.get(index);
    if (!sheet) {
      this.gridHost.append(message(this.root.ownerDocument, `Loading “${this.workbook.sheets[index]?.name ?? ''}”…`, 'dl-message dl-loading'));
      try {
        sheet = await this.workbook.loadSheet(index, loading.signal);
      } catch (error) {
        if (!loading.signal.aborted && !this.disposed) {
          this.showSheetError(index, error as Error);
        }
        return;
      }
      if (loading.signal.aborted || this.disposed) {
        return;
      }
      this.keep(index, sheet);
    }
    this.mount(index, sheet);
    this.activate();
  }

  setZoom(zoom: number): void {
    this.zoom = zoom;
    this.grid?.setZoom(zoom);
    this.updateNotices();
  }

  /** What to keep across reloads of the file and of the webview. */
  state(): SavedState {
    this.rememberPosition();
    const state: SavedState = { positions: Object.fromEntries([...this.positions].slice(-MAX_SAVED_POSITIONS)) };
    if (this.sheet) {
      state.sheet = this.sheet.name;
    }
    return state;
  }

  dispose(): void {
    this.disposed = true;
    this.loading?.abort();
    clearTimeout(this.saveTimer);
    this.grid?.dispose();
    this.rules.dispose();
    this.root.ownerDocument.removeEventListener('copy', this.onCopy);
    this.cellBar.element.remove();
    this.root.remove();
  }

  private keep(index: number, sheet: Sheet): void {
    if (sheet.cells.size <= CACHED_SHEET_CELLS) {
      this.cache.set(index, sheet);
    }
  }

  private mount(index: number, sheet: Sheet): void {
    this.grid?.dispose();
    this.current = index;
    this.sheet = sheet;
    this.grid = new Grid(this.root.ownerDocument, this.gridCallbacks(), sheet, this.zoom);
    this.gridHost.replaceChildren(this.grid.element);
  }

  /** Lays out the mounted grid (it must be on the page by now) and updates everything around it. */
  private activate(): void {
    if (!this.grid || !this.sheet) {
      return;
    }
    this.grid.start(this.positions.get(this.sheet.name));
    this.grid.focus();
    this.tabs.setActive(this.current);
    this.cellBar.show(this.grid.selectedInfo());
    this.updateNotices();
    this.saveSoon();
  }

  private gridCallbacks(): GridHost {
    return {
      workbook: this.workbook,
      rules: this.rules,
      goTo: (location) => this.goTo(location),
      selectionChanged: () => {
        this.cellBar.show(this.grid?.selectedInfo());
        this.saveSoon();
      },
      viewChanged: () => this.saveSoon(),
    };
  }

  private showSheetError(index: number, error: Error): void {
    this.grid?.dispose();
    this.grid = undefined;
    this.sheet = undefined;
    this.current = index;
    this.tabs.setActive(index);
    this.cellBar.show(undefined);
    this.updateNotices();
    const name = this.workbook.sheets[index]?.name ?? '';
    this.gridHost.replaceChildren(message(this.root.ownerDocument, `Could not display “${name}”: ${error.message}`, 'dl-message dl-sheet-error'));
  }

  private updateNotices(): void {
    const notices = [...this.workbook.notices, ...(this.sheet?.notices ?? [])];
    const cut = this.grid?.cutOff;
    if (cut?.rows || cut?.cols) {
      const what = cut.rows && cut.cols ? 'rows and columns' : cut.rows ? 'rows' : 'columns';
      notices.push(`Zoom out to see the last ${what} of this sheet: at this zoom the grid would be larger than the view can show.`);
    }
    if (this.sheet?.kind === 'macrosheet' || this.sheet?.kind === 'dialogsheet') {
      notices.push(`“${this.sheet.name}” is a ${this.sheet.kind === 'macrosheet' ? 'macro' : 'dialog'} sheet; only its cells are shown.`);
    }
    this.context.showNotice(notices.length ? notices.join(' ') : null);
  }

  private rememberPosition(): void {
    if (this.grid && this.sheet) {
      this.positions.delete(this.sheet.name);
      this.positions.set(this.sheet.name, this.grid.savedPosition());
    }
  }

  private saveSoon(): void {
    clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => {
      if (!this.disposed) {
        this.context.saveViewerState(this.state());
      }
    }, SAVE_DELAY_MS);
  }

  /** Follows a link inside the workbook: `Sheet2!B3`, `'My sheet'!B3` or `B3` on this sheet. */
  private goTo(location: string): void {
    const bang = location.lastIndexOf('!');
    const sheetName = bang < 0 ? undefined : location.slice(0, bang).replace(/^'(.*)'$/s, '$1').replaceAll("''", "'");
    const cell = parseCellRef((location.slice(bang + 1).split(':')[0] ?? '').trim());
    const index =
      sheetName === undefined ? this.current : this.workbook.sheets.findIndex((sheet) => !sheet.hidden && sheet.name.toLowerCase() === sheetName.toLowerCase());
    if (index < 0 || !cell) {
      return;
    }
    void this.show(index).then(() => {
      if (this.current === index) {
        this.grid?.goToCell(cell.row, cell.col);
      }
    });
  }

  /** Cmd/Ctrl+C with a cell selected copies the cell's text as shown. */
  private readonly onCopy = (event: ClipboardEvent) => {
    const document = this.root.ownerDocument;
    const selected = document.getSelection?.();
    if (!this.grid || !this.root.contains(document.activeElement) || (selected && !selected.isCollapsed)) {
      return;
    }
    const text = this.grid.selectedText();
    if (text !== undefined && event.clipboardData) {
      event.clipboardData.setData('text/plain', text);
      event.preventDefault();
    }
  };
}

function message(document: Document, text: string, className = 'dl-message'): HTMLElement {
  const element = document.createElement('div');
  element.className = className;
  element.textContent = text;
  return element;
}

/** The saved state as far as it is well-formed; the webview state is not trusted to have any shape. */
export function readSavedState(raw: unknown): SavedState {
  if (typeof raw !== 'object' || raw === null) {
    return {};
  }
  const value = raw as Record<string, unknown>;
  const state: SavedState = {};
  if (typeof value['sheet'] === 'string') {
    state.sheet = value['sheet'];
  }
  const positions = value['positions'];
  if (typeof positions === 'object' && positions !== null) {
    state.positions = {};
    for (const [name, position] of Object.entries(positions).slice(0, MAX_SAVED_POSITIONS)) {
      const p = (typeof position === 'object' && position !== null ? position : {}) as Record<string, unknown>;
      const top = p['top'];
      const left = p['left'];
      if (typeof top !== 'number' || typeof left !== 'number' || !Number.isFinite(top) || !Number.isFinite(left)) {
        continue;
      }
      const saved: GridPosition = { top: Math.max(0, top), left: Math.max(0, left) };
      const row = p['row'];
      const col = p['col'];
      if (typeof row === 'number' && typeof col === 'number' && Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0) {
        saved.row = row;
        saved.col = col;
      }
      state.positions[name] = saved;
    }
  }
  return state;
}
