// The spreadsheet viewer in Chrome through the browser preview: what the synthetic workbooks look like, frozen panes,
// sheet switching, links, state, zoom, a 100,000-row sheet, and files it cannot show. Set E2E_SCREENSHOTS to a
// folder to keep screenshots of each step.
import { expect, test, type Locator, type Page } from '@playwright/test';
import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import { ZipArchive } from '../../src/zip.ts';
import { XLS_LINK } from '../fixtures/xls.ts';
import { GITHUB_URL } from '../fixtures/xlsx.ts';
import { biffStream, compoundFile } from '../support/biffWriter.ts';
import { writeZip } from '../support/zipWriter.ts';
import { cell, largeSheetValue, largeSheetXlsx, xlsxPackage } from '../support/xlsxWriter.ts';
import { expectShellLoaded, servePreview, watchProblems, writeSample, type PreviewSite } from './preview-server.ts';

const SCREENSHOTS = process.env['E2E_SCREENSHOTS'];

function writeTemp(fileName: string, bytes: Uint8Array): string {
  const file = path.join(mkdtempSync(path.join(tmpdir(), 'document-lens-spreadsheet-')), fileName);
  writeFileSync(file, bytes);
  return file;
}

async function shoot(page: Page, name: string): Promise<void> {
  if (SCREENSHOTS) {
    mkdirSync(SCREENSHOTS, { recursive: true });
    await page.screenshot({ path: path.join(SCREENSHOTS, `${name}.png`) });
  }
}

/** Opens a file in the preview and waits for the grid; returns the problems seen and the site to close. */
async function open(page: Page, file: string): Promise<{ site: PreviewSite; problems: string[] }> {
  const site = await servePreview(file);
  const problems = watchProblems(page);
  await page.goto(site.url);
  await expectShellLoaded(page, path.basename(file));
  await expect(page.locator('.dl-grid .dl-cell').first()).toBeVisible({ timeout: 60_000 });
  return { site, problems };
}

/** The cell (or merged cell) showing exactly `text`. */
function cellWithText(page: Page, text: string): Locator {
  return page.locator('.dl-cell').filter({ hasText: new RegExp(`^${text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`) });
}

async function style(locator: Locator, property: string): Promise<string> {
  return locator.first().evaluate((element, name) => getComputedStyle(element).getPropertyValue(name), property);
}

async function box(locator: Locator): Promise<{ x: number; y: number; width: number; height: number }> {
  const bounds = await locator.first().boundingBox();
  expect(bounds).not.toBeNull();
  return bounds as { x: number; y: number; width: number; height: number };
}

const cellBar = (page: Page) => Promise.all([page.locator('.dl-cellbar-address').textContent(), page.locator('.dl-cellbar-value').textContent()]);

test('sample.xlsx: values, formats, merged cells, sizes and frozen panes as in the file', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.xlsx'));
  try {
    await expect(page.locator('.dl-tab')).toHaveText(['Summary', 'Data', 'メモ']);
    for (const text of ['1,234.50', '2024-01-15', '39.97%', '令和6年1月15日', 'TRUE', '#DIV/0!', '東京本社', 'Plain, bold red, italic']) {
      await expect(cellWithText(page, text).first()).toBeVisible();
    }
    await expect(cellWithText(page, 'secret')).toHaveCount(0);

    // The merged title is drawn in the frozen corner and continued in the top band: both span A1:D1.
    const title = page.locator('.dl-band-top > .dl-merge');
    expect(await style(title, 'background-color')).toBe('rgb(31, 78, 121)');
    expect(await style(title, 'color')).toBe('rgb(255, 255, 255)');
    expect(await style(title, 'font-weight')).toBe('700');
    const header = (name: string) => page.locator('.dl-col-header', { hasText: new RegExp(`^${name}$`) });
    const widths = await Promise.all(['A', 'B', 'C', 'D'].map(async (name) => (await box(header(name))).width));
    expect(widths).toEqual([173, 103, 89, 75]);
    const titleBox = await box(title);
    expect(titleBox.width).toBe(173 + 103 + 89 + 75 + 1);
    expect(titleBox.height).toBe(40 + 1);
    expect((await box(page.locator('.dl-row-header', { hasText: /^4$/ }))).height).toBe(60);
    await expect(page.locator('.dl-col-header', { hasText: /^F$/ })).toHaveCount(0);
    await expect(page.locator('.dl-row-header', { hasText: /^11$/ })).toHaveCount(0);

    expect(await style(cellWithText(page, '-120.00'), 'color')).toBe('rgb(255, 0, 0)');
    expect(await style(cellWithText(page, 'Item'), 'border-bottom')).toBe('2px solid rgb(0, 0, 0)');
    expect(await style(cellWithText(page, 'Accent color'), 'color')).toBe('rgb(192, 80, 77)');
    expect(await style(cellWithText(page, 'Accent color'), 'font-style')).toBe('italic');
    expect(await style(cellWithText(page, 'Document Lens on GitHub'), 'text-decoration-line')).toBe('underline');
    const wrapped = cellWithText(page, 'This note is long enough to wrap onto a second line inside its cell. 長いメモは折り返して表示されます。');
    expect(await style(wrapped, 'white-space')).toBe('pre-wrap');

    // Text in B12 runs over the empty cells to its right.
    const spill = await box(page.locator('.dl-spill-text', { hasText: 'This title runs across' }));
    expect(spill.width).toBeGreaterThan(103 + 89);

    await shoot(page, 'xlsx-summary');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('sample.xlsx: frozen rows stay put while the rest scrolls, on a sheet switched to with its tab', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.xlsx'));
  try {
    await page.locator('.dl-tab', { hasText: 'Data' }).click();
    await expect(page.locator('.dl-tab[aria-selected="true"]')).toHaveText('Data');
    const id = cellWithText(page, 'ID');
    const product = cellWithText(page, 'Product 5');
    const before = { id: await box(id), product: await box(product) };
    await page.locator('.dl-grid').evaluate((grid) => {
      grid.scrollTop = 600;
    });
    await expect(cellWithText(page, 'Product 40')).toBeInViewport();
    expect((await box(id)).y).toBe(before.id.y);
    await expect(product).not.toBeInViewport();
    expect(before.product.y).toBeGreaterThan(before.id.y);
    await shoot(page, 'xlsx-data-scrolled');

    await page.locator('.dl-tab', { hasText: 'メモ' }).click();
    await expect(cellWithText(page, '日本語のメモ')).toBeVisible();
    await expect(page.locator('.dl-gridline').first()).toBeHidden();
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('sample.xlsx: selection, cell bar, keyboard, links, reload and saved state', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.xlsx'));
  try {
    expect(await cellBar(page)).toEqual(['B3', '1234.5']);
    await cellWithText(page, '39.97%').click();
    expect(await cellBar(page)).toEqual(['D4', '=B4/$B$8']);
    await page.keyboard.press('ArrowDown');
    expect(await cellBar(page)).toEqual(['D5', '=B5/$B$8']);
    await expect(page.locator('.dl-row-header.dl-selected')).toHaveText('5');

    await page.locator('a.dl-link').click();
    await expect.poll(() => page.evaluate(() => (window as unknown as { previewLinks: string[] }).previewLinks)).toEqual([GITHUB_URL]);
    await page.locator('[data-location]').click();
    await expect(page.locator('.dl-tab[aria-selected="true"]')).toHaveText('Data');
    await expect(page.locator('.dl-cellbar-address')).toHaveText('A1');
    expect(await cellBar(page)).toEqual(['A1', 'ID']);

    await page.keyboard.press('ArrowDown');
    const before = await page.locator('.dl-sheet').elementHandle();
    await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
    await expect.poll(() => before?.evaluate((element) => element.isConnected)).toBe(false);
    await expect(page.locator('.dl-tab[aria-selected="true"]')).toHaveText('Data');
    expect(await cellBar(page)).toEqual(['A2', '1']);

    await page.waitForTimeout(400);
    await page.reload();
    await expect(page.locator('.dl-tab[aria-selected="true"]')).toHaveText('Data');
    await expect(page.locator('.dl-cellbar-address')).toHaveText('A2');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('sample.xlsx: zoom scales the grid and keeps virtual scrolling in step', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.xlsx'));
  try {
    await page.locator('#zoom-in').click();
    await page.locator('#zoom-in').click();
    await expect(page.locator('#zoom-value')).toHaveText('125%');
    expect((await box(page.locator('.dl-col-header', { hasText: /^A$/ }))).width).toBe(Math.round(173 * 1.25));
    expect(parseFloat(await style(cellWithText(page, '1,234.50'), 'font-size'))).toBeCloseTo(((11 * 4) / 3) * 1.25, 2);

    await page.locator('.dl-tab', { hasText: 'Data' }).click();
    await page.locator('.dl-grid').evaluate((grid) => {
      grid.scrollTop = 4000;
    });
    // At 125% a row is 25px: about row 160 is at the top, under the frozen header row.
    await expect(cellWithText(page, 'Product 165')).toBeVisible();
    const headerRow = await box(page.locator('.dl-row-header', { hasText: /^1$/ }));
    const firstVisible = await box(page.locator('.dl-col-header', { hasText: /^A$/ }));
    expect(headerRow.y).toBe(firstVisible.y + firstVisible.height);
    await shoot(page, 'xlsx-zoomed');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('sample.xls: an Excel 97-2003 workbook with its formatting, 1904 dates and links', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.xls'));
  try {
    await expect(page.locator('.dl-tab')).toHaveText(['Sales', 'Second']);
    const title = page.locator('.dl-merge', { hasText: 'Sales 売上' });
    expect(await style(title, 'background-color')).toBe('rgb(0, 0, 128)');
    // Columns of 24, 14 and 12 characters of Arial 10.
    expect((await box(title)).width).toBe(168 + 98 + 84 + 1);
    for (const text of ['1,500.00', '2024/01/15', '5,230.50', '12.50%', 'TRUE', '#DIV/0!', 'Tokyo 東京']) {
      await expect(cellWithText(page, text)).toBeVisible();
    }
    expect(await style(cellWithText(page, 'Region'), 'background-color')).toBe('rgb(192, 192, 192)');
    expect(await style(cellWithText(page, 'Italic red'), 'color')).toBe('rgb(255, 0, 0)');
    await cellWithText(page, '5,230.50').click();
    expect(await cellBar(page)).toEqual(['B6', '=SUM(B3:B5)']);
    await page.locator('a.dl-link').click();
    await expect.poll(() => page.evaluate(() => (window as unknown as { previewLinks: string[] }).previewLinks)).toEqual([XLS_LINK]);
    await shoot(page, 'xls-sales');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('a 100,000 × 50 sheet: only a window of cells exists, scrolling anywhere shows the right cells', async ({ page }) => {
  test.setTimeout(180_000);
  const file = writeTemp('large.xlsx', largeSheetXlsx(100_000, 50));
  const { site, problems } = await open(page, file);
  try {
    const grid = page.locator('.dl-grid');
    expect(await page.locator('.dl-cell').count()).toBeLessThan(10_000);
    const corner = cellWithText(page, String(largeSheetValue(0, 0)));
    const cornerBefore = await box(corner);

    // Steady scrolling, timed: drawing a new window must stay cheap (it once took ~1 s per window).
    const elapsed = await grid.evaluate(async (element) => {
      const started = performance.now();
      for (let i = 0; i < 60; i++) {
        element.scrollTop += 300;
        element.dispatchEvent(new Event('scroll'));
        await new Promise((resolve) => requestAnimationFrame(resolve));
      }
      return performance.now() - started;
    });
    test.info().annotations.push({ type: 'scrolling', description: `60 steps of 300px in ${Math.round(elapsed)} ms` });
    expect(elapsed).toBeLessThan(6000);

    await grid.evaluate((element) => {
      element.scrollTop = element.scrollHeight;
      element.scrollLeft = element.scrollWidth;
    });
    await expect(cellWithText(page, String(largeSheetValue(99_999, 49)))).toBeInViewport();
    await expect(page.locator('.dl-row-header', { hasText: /^100000$/ })).toBeInViewport();
    await expect(cellWithText(page, String(largeSheetValue(99_999, 0)))).toBeInViewport();
    expect(await page.locator('.dl-cell').count()).toBeLessThan(10_000);
    // The frozen first row and column stay where they were.
    expect(await box(corner)).toEqual(cornerBefore);
    await shoot(page, 'large-bottom-right');

    const middle = await grid.evaluate((element) => {
      element.scrollTop = element.scrollHeight / 2;
      element.scrollLeft = 0;
      // Rows are 20px: a row ten rows below the top of the scrolling part.
      return Math.floor(element.scrollTop / 20) + 10;
    });
    await expect(page.locator('.dl-row-header', { hasText: new RegExp(`^${middle + 1}$`) })).toBeInViewport();
    await expect(cellWithText(page, String(largeSheetValue(middle, 3)))).toBeInViewport();
    await shoot(page, 'large-middle');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('files that cannot be shown get a clear message', async ({ page }) => {
  const sheetlessWorkbook = writeTemp('no-sheets.xlsx', xlsxPackage({ sheets: [] }));
  const encrypted = (() => {
    const stream = biffStream({ fonts: [{ name: 'Arial', size: 10 }], xfs: [{ font: 0 }], sheets: [{ name: 'S', cells: [{ row: 0, col: 0, xf: 0, value: 1 }] }] });
    const filePass = Buffer.from([0x2f, 0x00, 0x06, 0x00, 0x00, 0x00, 0x34, 0x12, 0x78, 0x56]);
    const bofEnd = 4 + stream.readUInt16LE(2);
    return writeTemp('locked.xls', compoundFile(Buffer.concat([stream.subarray(0, bofEnd), filePass, stream.subarray(bofEnd)])));
  })();
  const webPage = writeTemp('export.xls', new TextEncoder().encode('<html><body><table><tr><td>1</td></tr></table></body></html>'));
  const cases: [string, RegExp][] = [
    [sheetlessWorkbook, /^Could not display no-sheets\.xlsx: Not a spreadsheet: the workbook lists no sheets$/],
    [encrypted, /^Could not display locked\.xls: This \.xls file is password-protected; Document Lens cannot open encrypted files$/],
    [webPage, /^This \.xls file is a web page or XML spreadsheet with an \.xls extension/],
  ];
  for (const [file, message] of cases) {
    const site = await servePreview(file);
    try {
      const problems = watchProblems(page);
      await page.goto(site.url);
      await expect(page.locator('#error')).toHaveText(message);
      expect(problems).toEqual([]);
    } finally {
      await site.close();
    }
  }
});

test('a sheet that cannot be read shows its error in place of the grid; the others still open', async ({ page }) => {
  const good = xlsxPackage({ sheets: [{ name: 'Good', body: `<sheetData><row r="1">${cell('A1', 'fine')}</row></sheetData>` }, { name: 'Broken', body: '<sheetData/>' }] });
  // Replace the second sheet's data with bytes that are not deflate data, keeping every header valid.
  const entries: { name: string; data: Buffer; store?: boolean }[] = [];
  const zip = ZipArchive.open(good);
  for (const name of zip.names()) {
    entries.push({ name, data: name === 'xl/worksheets/sheet2.xml' ? Buffer.alloc(32, 0xff) : zip.read(name), store: name === 'xl/worksheets/sheet2.xml' });
  }
  const broken = writeZip(entries);
  const header = broken.indexOf(Buffer.from('xl/worksheets/sheet2.xml')) - 30;
  broken.writeUInt16LE(8, header + 8);
  broken.writeUInt16LE(8, broken.lastIndexOf(Buffer.from('xl/worksheets/sheet2.xml')) - 46 + 10);
  const { site, problems } = await open(page, writeTemp('half-broken.xlsx', broken));
  try {
    await page.locator('.dl-tab', { hasText: 'Broken' }).click();
    await expect(page.locator('.dl-sheet-error')).toHaveText('Could not display “Broken”: Corrupt ZIP entry (bad compressed data): xl/worksheets/sheet2.xml');
    await page.locator('.dl-tab', { hasText: 'Good' }).click();
    await expect(cellWithText(page, 'fine')).toBeVisible();
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

