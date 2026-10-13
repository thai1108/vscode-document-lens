// The spreadsheet viewer in the shell, in happy-dom: tabs, cell bar, selection, links, state and the virtual window.
// happy-dom does no layout, so the grid sizes itself from the window (1200 × 800); geometry and looks are checked
// in Chrome by test/e2e/spreadsheet.spec.ts.
import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import type { HostMessage } from '../../../src/protocol.ts';
import { spreadsheetViewer } from '../../../src/webview/spreadsheet/viewer.ts';
import type { ShellState } from '../../../src/webview/shell/vscodeApi.ts';
import { GITHUB_URL, sampleXlsx } from '../../fixtures/xlsx.ts';
import { sampleXls } from '../../fixtures/xls.ts';
import { cell, largeSheetValue, largeSheetXlsx, xlsxPackage } from '../../support/xlsxWriter.ts';
import { bootShell, type FakeHost } from '../dom/env.ts';

let host: FakeHost | undefined;
afterEach(() => host?.close());

/** The bytes go out as a plain Uint8Array, as the host sends them (src/payload.ts), not as a Node Buffer. */
function documentMessage(fileName: string, data: Uint8Array, format = 'xlsx'): HostMessage {
  return { type: 'document', format, fileName, data: new Uint8Array(data) };
}

async function open(fileName: string, data: Uint8Array, options: { state?: ShellState; format?: string } = {}): Promise<FakeHost> {
  const shell = await bootShell(spreadsheetViewer(), options.state ? { state: options.state } : {});
  host = shell;
  shell.deliver(documentMessage(fileName, data, options.format));
  await shell.until(() => shell.window.document.querySelector('.dl-grid .dl-cell') !== null || !shell.$('#error').hidden, 'the grid', 5000);
  return shell;
}

const texts = (shell: FakeHost, selector: string) => [...shell.window.document.querySelectorAll(selector)].map((node) => node.textContent ?? '');
const cellBar = (shell: FakeHost) => [shell.$('.dl-cellbar-address').textContent, shell.$('.dl-cellbar-value').textContent];
const grid = (shell: FakeHost) => shell.$<HTMLElement>('.dl-grid');

function mouseDown(shell: FakeHost, clientX: number, clientY: number): void {
  grid(shell).dispatchEvent(new shell.window.MouseEvent('mousedown', { bubbles: true, button: 0, clientX, clientY }) as unknown as Event);
}

function scroll(shell: FakeHost, left: number, top: number): void {
  grid(shell).scrollLeft = left;
  grid(shell).scrollTop = top;
  grid(shell).dispatchEvent(new shell.window.Event('scroll') as unknown as Event);
}

const oneSheet = (body: string, styles?: string) => new Uint8Array(xlsxPackage({ sheets: [{ name: 'S', body }], ...(styles ? { styles } : {}) }));

async function tab(shell: FakeHost, name: string): Promise<void> {
  const button = [...shell.window.document.querySelectorAll('.dl-tab')].find((node) => node.textContent === name);
  assert.ok(button, `a tab named ${name}`);
  shell.click(button as unknown as Element);
  await shell.until(() => shell.$('.dl-grid').getAttribute('aria-label') === name, `sheet ${name}`);
}

describe('spreadsheet viewer', () => {
  it('shows the visible sheets as tabs, the active one first, and its cells', async () => {
    const shell = await open('sample.xlsx', sampleXlsx());
    assert.deepEqual(texts(shell, '.dl-tab'), ['Summary', 'Data', 'メモ']);
    assert.equal(shell.$('.dl-tab[aria-selected="true"]').textContent, 'Summary');
    const cells = texts(shell, '.dl-cell');
    for (const text of ['Quarterly report 四半期報告', '1,234.50', '2024-01-15', '令和6年1月15日', 'Plain, bold red, italic', '東京本社']) {
      assert.ok(cells.includes(text), `shows ${text}`);
    }
    assert.ok(!cells.includes('secret'), 'cells of hidden columns are not drawn');
    assert.ok(!cells.includes('hidden row'), 'cells of hidden rows are not drawn');
    assert.deepEqual(texts(shell, '.dl-col-header').slice(0, 6), ['A', 'B', 'C', 'D', 'E', 'G']);
    assert.equal(shell.$('#notice').hidden, true);
  });

  it('starts with the cell the file selected in the cell bar, and moves the selection with the keyboard', async () => {
    const shell = await open('sample.xlsx', sampleXlsx());
    assert.deepEqual(cellBar(shell), ['B3', '1234.5']);
    shell.key('ArrowDown', {}, grid(shell));
    assert.deepEqual(cellBar(shell), ['B4', '980.25']);
    shell.key('ArrowDown', { metaKey: true }, grid(shell));
    shell.key('ArrowUp', {}, grid(shell));
    shell.key('ArrowUp', {}, grid(shell));
    shell.key('ArrowUp', {}, grid(shell));
    shell.key('ArrowUp', {}, grid(shell));
    assert.deepEqual(cellBar(shell), ['B10', ''], 'the hidden row 11 is stepped over');
    shell.key('ArrowLeft', {}, grid(shell));
    shell.key('ArrowUp', {}, grid(shell));
    shell.key('ArrowUp', {}, grid(shell));
    assert.deepEqual(cellBar(shell), ['A8', 'Total']);
    shell.key('Home', { ctrlKey: true }, grid(shell));
    assert.deepEqual(cellBar(shell), ['A1:D1', 'Quarterly report 四半期報告'], 'a merged cell is selected whole');
    shell.key('ArrowRight', {}, grid(shell));
    assert.deepEqual(cellBar(shell), ['E1', ''], 'moving right leaves the merged cell');
    assert.equal(shell.$('.dl-col-header.dl-selected').textContent, 'E');
    assert.ok(shell.window.document.querySelector('.dl-selection'));
  });

  it('switches sheets, selects cells with the mouse and remembers both in the webview state', async () => {
    const shell = await open('sample.xlsx', sampleXlsx());
    await tab(shell, 'Data');
    assert.ok(texts(shell, '.dl-cell').includes('Product 1'));
    // Row headers are 33px wide (3 digits), the header row 20px and the frozen row 1 another 20px.
    mouseDown(shell, 33 + 64 + 10, 20 + 20 + 25);
    assert.deepEqual(cellBar(shell), ['B3', 'Product 2']);
    await shell.until(() => (shell.state?.viewer as { sheet?: string } | undefined)?.sheet === 'Data', 'saved state');
    const saved = shell.state?.viewer as { positions: Record<string, { row?: number; col?: number }> };
    assert.deepEqual([saved.positions['Data']?.row, saved.positions['Data']?.col], [2, 1]);

    // Reopening the webview with that state comes back to the same sheet and cell.
    const state = structuredClone(shell.state) as ShellState;
    await shell.close();
    host = undefined;
    const again = await open('sample.xlsx', sampleXlsx(), { state });
    assert.equal(again.$('.dl-tab[aria-selected="true"]').textContent, 'Data');
    assert.deepEqual(cellBar(again), ['B3', 'Product 2']);
  });

  it('keeps the sheet and selection when the file is reloaded', async () => {
    const shell = await open('sample.xlsx', sampleXlsx());
    await tab(shell, 'メモ');
    shell.key('ArrowRight', {}, grid(shell));
    const before = shell.$('.dl-sheet');
    shell.deliver(documentMessage('sample.xlsx', sampleXlsx()));
    await shell.until(() => !before.isConnected && shell.window.document.querySelectorAll('.dl-sheet').length === 1, 'reloaded');
    assert.equal(shell.$('.dl-tab[aria-selected="true"]').textContent, 'メモ');
    assert.deepEqual(cellBar(shell), ['B1', '42']);
    assert.ok(texts(shell, '.dl-cell').includes('日本語のメモ'));
    assert.equal(shell.window.document.querySelectorAll('style[nonce]').length, 1, 'the previous file’s styles are gone');
  });

  it('opens web links through the host and follows links inside the workbook', async () => {
    const shell = await open('sample.xlsx', sampleXlsx());
    const web = shell.$('a.dl-link');
    assert.equal(web.getAttribute('href'), GITHUB_URL);
    assert.equal(shell.click(web).defaultPrevented, true);
    assert.deepEqual(shell.posted.at(-1), { type: 'openLink', url: GITHUB_URL });
    shell.click(shell.$('[data-location]'));
    await shell.until(() => shell.$('.dl-tab[aria-selected="true"]').textContent === 'Data', 'the linked sheet');
    await shell.until(() => shell.$('.dl-cellbar-address').textContent === 'A1', 'the linked cell');
  });

  it('copies the selected cell as shown', async () => {
    const shell = await open('sample.xlsx', sampleXlsx());
    const copied: string[] = [];
    const event = new shell.window.Event('copy', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: { setData: (_type: string, text: string) => copied.push(text) } });
    grid(shell).dispatchEvent(event as unknown as Event);
    assert.deepEqual(copied, ['1,234.50']);
    assert.equal(event.defaultPrevented, true);
  });

  it('applies the shared zoom to the grid itself', async () => {
    const shell = await open('sample.xlsx', sampleXlsx());
    shell.key('+');
    assert.equal(grid(shell).style.getPropertyValue('--dl-zoom'), '1.1');
    assert.equal(shell.$('#content').style.getPropertyValue('zoom'), '', 'no CSS zoom on the content');
  });

  it('draws only the window around the viewport of a large sheet, and follows scrolling', async () => {
    const shell = await open('large.xlsx', largeSheetXlsx(5000, 30));
    const count = shell.window.document.querySelectorAll('.dl-cell').length;
    assert.ok(count > 100 && count < 5000, `${count} cells drawn of 150000`);
    assert.ok(!texts(shell, '.dl-cell').includes(String(largeSheetValue(3000, 5))));
    grid(shell).scrollTop = 3000 * 20;
    grid(shell).dispatchEvent(new shell.window.Event('scroll') as unknown as Event);
    assert.ok(texts(shell, '.dl-cell').includes(String(largeSheetValue(3000, 5))), 'row 3001 is drawn after scrolling to it');
    assert.ok(texts(shell, '.dl-row-header').includes('1'), 'the frozen first row stays');
    assert.ok(shell.window.document.querySelectorAll('.dl-cell').length < 5000);
  });

  it('reads .xls workbooks the same way', async () => {
    const shell = await open('sample.xls', sampleXls(), { format: 'xls' });
    assert.deepEqual(texts(shell, '.dl-tab'), ['Sales', 'Second']);
    assert.ok(texts(shell, '.dl-cell').includes('2024/01/15'));
    assert.ok(texts(shell, '.dl-cell').includes('Sales 売上'));
  });

  it('shows an error banner for a file it cannot read', async () => {
    const shell = await open('broken.xlsx', new TextEncoder().encode('PK\u0003\u0004 not really a zip'));
    assert.match(shell.$('#error').textContent ?? '', /^Could not display broken\.xlsx: Not a ZIP archive/);
    assert.equal(shell.window.document.querySelector('.dl-sheet'), null);
  });
});

describe('spreadsheet viewer with crafted sheets', () => {
  it('draws one of many repeated merged cells', async () => {
    const merges = '<mergeCell ref="A1:J40"/>'.repeat(20_000);
    const shell = await open('merges.xlsx', oneSheet(`<sheetData><row r="1">${cell('A1', 'merged')}</row></sheetData><mergeCells>${merges}</mergeCells>`));
    assert.equal(shell.window.document.querySelectorAll('.dl-merge').length, 1);
  });

  it('reaches the last row of a full-height merge, and asks to zoom out only when the grid gets too tall', async () => {
    const shell = await open('tall.xlsx', oneSheet(`<sheetData><row r="1">${cell('A1', 'tall')}</row></sheetData><mergeCells><mergeCell ref="A1:B1048576"/></mergeCells>`));
    assert.equal(shell.$('#notice').hidden, true, 'every row fits at 100%');
    scroll(shell, 0, 1_048_576 * 20);
    assert.ok(texts(shell, '.dl-row-header').includes('1048576'), 'the last row is drawn');
    for (let i = 0; i < 5; i++) {
      shell.key('+');
    }
    assert.equal(grid(shell).style.getPropertyValue('--dl-zoom'), '2');
    assert.match(shell.$('#notice').textContent ?? '', /Zoom out to see the last rows of this sheet/);
  });

  it('freezes the rows from the top of the view, not from row 1', async () => {
    const rows = Array.from({ length: 10 }, (_value, i) => `<row r="${i + 1}">${cell(`A${i + 1}`, `row ${i + 1}`)}</row>`).join('');
    const view = '<sheetViews><sheetView workbookViewId="0" topLeftCell="A5"><pane ySplit="2" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>';
    const shell = await open('frozen.xlsx', oneSheet(`${view}<sheetData>${rows}</sheetData>`));
    assert.ok(grid(shell).classList.contains('dl-frozen-rows'));
    const headers = texts(shell, '.dl-row-header');
    assert.deepEqual(headers.slice(0, 4), ['5', '6', '7', '8']);
    assert.ok(!headers.includes('4') && !texts(shell, '.dl-cell').includes('row 1'), 'rows scrolled out above the pane are not shown');
    scroll(shell, 0, 0);
    assert.deepEqual(texts(shell, '.dl-row-header').slice(0, 2), ['5', '6'], 'the frozen rows stay');
  });

  it('keeps text running over formatted blank cells after scrolling its own cell out of the window', async () => {
    const blanks = Array.from({ length: 25 }, (_value, i) => `<c r="${String.fromCharCode(66 + i)}1" s="1"/>`).join('');
    const styles =
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<borders count="2"><border/><border><bottom style="thin"><color rgb="FF000000"/></bottom></border></borders>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/></cellXfs>';
    const text = 'a heading that runs across the whole row';
    const shell = await open('spill.xlsx', oneSheet(`<sheetData><row r="1">${cell('A1', text)}${blanks}</row><row r="2">${cell('CC2', 'far')}</row></sheetData>`, styles));
    assert.ok(texts(shell, '.dl-cell').includes(text));
    scroll(shell, 40 * 64, 0);
    assert.ok(!texts(shell, '.dl-col-header').includes('A'), 'column A is out of the drawn window');
    assert.ok(texts(shell, '.dl-cell').includes(text), 'its text is still drawn where it runs');
  });
});
