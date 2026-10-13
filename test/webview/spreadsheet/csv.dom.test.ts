// .csv and .tsv files in the spreadsheet viewer, in happy-dom: cells as written, the CSV bar (dialect, header row
// switch, open as text), notices for what does not fit, and the state kept across reloads. Looks are checked in
// Chrome by test/e2e/csv.spec.ts.
import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import type { HostMessage } from '../../../src/protocol.ts';
import { openCsv } from '../../../src/webview/spreadsheet/csv/workbook.ts';
import { readCell } from '../../../src/webview/spreadsheet/model/display.ts';
import { spreadsheetViewer } from '../../../src/webview/spreadsheet/viewer.ts';
import type { ShellState } from '../../../src/webview/shell/vscodeApi.ts';
import { SAMPLE_CSV_ROWS, sampleCsv, semicolonCsv, shiftJisCsv } from '../../fixtures/csv.ts';
import { sampleTsv } from '../../fixtures/tsv.ts';
import { bootShell, type FakeHost } from '../dom/env.ts';

let host: FakeHost | undefined;
afterEach(() => host?.close());

function documentMessage(fileName: string, data: Uint8Array): HostMessage {
  return { type: 'document', format: fileName.endsWith('.tsv') ? 'tsv' : 'csv', fileName, data: new Uint8Array(data) };
}

async function open(fileName: string, data: Uint8Array, state?: ShellState): Promise<FakeHost> {
  const shell = await bootShell(spreadsheetViewer(), state ? { state } : {});
  host = shell;
  shell.deliver(documentMessage(fileName, data));
  await shell.until(() => shell.window.document.querySelector('.dl-grid') !== null || !shell.$('#error').hidden, 'the grid', 5000);
  return shell;
}

const texts = (shell: FakeHost, selector: string) => [...shell.window.document.querySelectorAll(selector)].map((node) => node.textContent ?? '');
const cellBar = (shell: FakeHost) => [shell.$('.dl-cellbar-address').textContent, shell.$('.dl-cellbar-value').getAttribute('title')];
const grid = (shell: FakeHost) => shell.$<HTMLElement>('.dl-grid');
const encode = (text: string) => new TextEncoder().encode(text);

describe('openCsv', () => {
  it('keeps every field as the text it is, numbers included', async () => {
    const workbook = openCsv(sampleCsv(), 'sample.csv', { tabSeparated: false });
    const sheet = await workbook.loadSheet(0);
    SAMPLE_CSV_ROWS.forEach((row, r) =>
      row.forEach((value, c) => assert.equal(readCell(workbook, sheet, r, c)?.raw ?? '', value, `row ${r + 1} column ${c + 1}`)),
    );
    assert.equal(readCell(workbook, sheet, 1, 4)?.text, '007');
    assert.equal(readCell(workbook, sheet, 2, 4)?.text, '1.50');
    assert.deepEqual(sheet.used, { rows: 6, cols: 6 });
    assert.deepEqual([workbook.csv.delimiter, workbook.csv.encoding, workbook.csv.lineEnding, workbook.csv.headerRow], [',', { name: 'utf-8', bom: 3 }, '\r\n', true]);
    assert.deepEqual(sheet.frozen, { rows: 1, cols: 0, top: 0, left: 0 });
  });

  it('right-aligns fields that read as numbers and sizes columns to their text', async () => {
    const workbook = openCsv(encode('n,t\n1,abc\n"1,234.5",x\n-3e5,12%\n'), 'n.csv', { tabSeparated: false });
    const sheet = await workbook.loadSheet(0);
    const style = (row: number, col: number) => readCell(workbook, sheet, row, col)?.style;
    assert.deepEqual([style(1, 0), style(2, 0), style(3, 0), style(3, 1), style(1, 1)], [2, 2, 2, 2, 0]);
    const wide = await openCsv(encode(`a,b\n${'w'.repeat(30)},${'x'.repeat(500)}\n`), 'w.csv', { tabSeparated: false }).loadSheet(0);
    assert.ok((wide.cols.sizes.get(0) ?? 0) > 200);
    assert.equal(wide.cols.sizes.get(1), 400, 'capped');
  });

  it('guesses whether the first row is a header, and follows the switch', async () => {
    assert.equal(openCsv(encode('1,2\n3,4\n'), 'a.csv', { tabSeparated: false }).csv.headerRow, false);
    assert.equal(openCsv(encode('a,a\n3,4\n'), 'a.csv', { tabSeparated: false }).csv.headerRow, false, 'repeated names');
    assert.equal(openCsv(encode('a,b\n'), 'a.csv', { tabSeparated: false }).csv.headerRow, false, 'one row');
    const forced = openCsv(sampleCsv(), 'sample.csv', { tabSeparated: false, headerRow: false });
    assert.equal(forced.csv.headerRow, false);
    assert.equal((await forced.loadSheet(0)).frozen.rows, 0);
  });

  it('caps columns, field length and rows with notices, and says when a quote is never closed', async () => {
    const wide = await openCsv(encode(`${','.repeat(20_000)}x\n`), 'w.csv', { tabSeparated: false }).loadSheet(0);
    assert.equal(wide.used.cols, 0, 'the field past the cap is not kept');
    assert.match(wide.notices.join(' '), /Fields after column 16,384 are not shown/);
    const long = await openCsv(encode(`a\n${'y'.repeat(40_000)}\n`), 'l.csv', { tabSeparated: false }).loadSheet(0);
    assert.match(long.notices.join(' '), /Fields longer than 32,767 characters are cut off/);
    const open = await openCsv(encode('a,b\n1,"never closed\n2,3\n'), 'o.csv', { tabSeparated: false }).loadSheet(0);
    assert.match(open.notices.join(' '), /A quote opened in row 2 is never closed/);
    const rows = await openCsv(encode('x\n'.repeat(1_048_577)), 'r.csv', { tabSeparated: false }).loadSheet(0);
    assert.match(rows.notices.join(' '), /Only the first 1,048,576 rows are shown/);
  });

  it('notes bytes that are not valid text', () => {
    assert.match(openCsv(new Uint8Array([0x61, 0x0a, 0xff, 0xfe, 0x0a]), 'bad.csv', { tabSeparated: false }).notices.join(' '), /not valid UTF-8 or Shift_JIS text/);
  });
});

describe('spreadsheet viewer with .csv and .tsv files', () => {
  it('shows the cells, a frozen header row and how the file is written, without sheet tabs', async () => {
    const shell = await open('sample.csv', sampleCsv());
    const cells = texts(shell, '.dl-cell');
    for (const text of ['Coffee beans', 'Green tea 緑茶', '1,234.50', 'Fair trade, organic', 'Says "crunchy"', '007', 'extra field']) {
      assert.ok(cells.includes(text), `shows ${text}`);
    }
    assert.ok(cells.includes('Line one line two'), 'a line break inside a field shows as a space in the grid');
    assert.equal(shell.$('#dl-csv-info').textContent, 'Comma · UTF-8 with BOM · CRLF');
    assert.equal(shell.$('#dl-csv-header').getAttribute('aria-pressed'), 'true');
    assert.ok(grid(shell).classList.contains('dl-frozen-rows'));
    assert.equal(shell.$('.dl-tabs').hidden, true);
    assert.equal(shell.$('#notice').hidden, true);
    assert.equal(grid(shell).getAttribute('aria-label'), 'sample.csv');
  });

  it('shows a field’s full text, line breaks included, in the cell bar', async () => {
    const shell = await open('sample.csv', sampleCsv());
    shell.key('ArrowDown', {}, grid(shell));
    for (let i = 0; i < 3; i++) {
      shell.key('ArrowRight', {}, grid(shell));
    }
    shell.key('ArrowDown', {}, grid(shell));
    assert.deepEqual(cellBar(shell), ['D3', 'Line one\nline two']);
  });

  it('asks the host to reopen the file as text', async () => {
    const shell = await open('sample.csv', sampleCsv());
    shell.click(shell.$('#dl-open-as-text'));
    assert.deepEqual(shell.posted.at(-1), { type: 'openAsText' });
  });

  it('turns the header row off and on, and remembers the choice', async () => {
    const shell = await open('sample.csv', sampleCsv());
    shell.click(shell.$('#dl-csv-header'));
    await shell.until(() => shell.$('#dl-csv-header').getAttribute('aria-pressed') === 'false', 'header row off');
    assert.ok(!grid(shell).classList.contains('dl-frozen-rows'));
    await shell.until(() => (shell.state?.viewer as { headerRow?: boolean } | undefined)?.headerRow === false, 'saved state');
    const state = structuredClone(shell.state) as ShellState;
    await shell.close();
    host = undefined;
    const again = await open('sample.csv', sampleCsv(), state);
    assert.equal(again.$('#dl-csv-header').getAttribute('aria-pressed'), 'false');
    again.click(again.$('#dl-csv-header'));
    await again.until(() => again.$('#dl-csv-header').getAttribute('aria-pressed') === 'true', 'header row on');
    assert.ok(grid(again).classList.contains('dl-frozen-rows'));
  });

  it('reads Shift_JIS, semicolons and UTF-16 tab-separated files', async () => {
    for (const [fileName, bytes, info, text] of [
      ['shift-jis.csv', shiftJisCsv(), 'Comma · Shift_JIS · CRLF', '緑茶（ペットボトル）'],
      ['semicolon.csv', semicolonCsv(), 'Semicolon · UTF-8 · LF', '12,50'],
      ['sample.tsv', sampleTsv(), 'Tab · UTF-16 LE · CRLF', 'Emoji 🎉 and comma, here'],
    ] as const) {
      const shell = await open(fileName, bytes);
      assert.equal(shell.$('#dl-csv-info').textContent, info);
      assert.ok(texts(shell, '.dl-cell').includes(text), `${fileName} shows ${text}`);
      await shell.close();
      host = undefined;
    }
  });

  it('keeps the selection when the file is reloaded', async () => {
    const shell = await open('sample.csv', sampleCsv());
    shell.key('ArrowDown', {}, grid(shell));
    shell.key('ArrowRight', {}, grid(shell));
    const before = shell.$('.dl-sheet');
    shell.deliver(documentMessage('sample.csv', sampleCsv()));
    await shell.until(() => !before.isConnected && shell.window.document.querySelectorAll('.dl-sheet').length === 1, 'reloaded');
    assert.deepEqual(cellBar(shell), ['B2', 'Coffee beans']);
    assert.equal(shell.window.document.querySelectorAll('.dl-csvbar').length, 1);
  });

  it('shows an error for binary data', async () => {
    const shell = await open('image.csv', new Uint8Array(4096).map((_byte, i) => (i % 3 ? 7 : 0)));
    assert.match(shell.$('#error').textContent ?? '', /^Could not display image\.csv: The file contains binary data, not text$/);
  });
});
