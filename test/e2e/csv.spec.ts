// .csv and .tsv files in Chrome through the browser preview: the grid with a frozen header row, the CSV bar, the way
// back to the text editor, encodings, a 200,000-row file and files that cannot or can only partly be shown. Set
// E2E_SCREENSHOTS to a folder to keep screenshots of each step.
import { expect, test, type Locator, type Page } from '@playwright/test';
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import { expectShellLoaded, servePreview, watchProblems, writeSample, type PreviewSite } from './preview-server.ts';

const SCREENSHOTS = process.env['E2E_SCREENSHOTS'];

function writeTemp(fileName: string, bytes: Uint8Array | string): string {
  const file = path.join(mkdtempSync(path.join(tmpdir(), 'document-lens-csv-')), fileName);
  writeFileSync(file, bytes);
  return file;
}

async function shoot(page: Page, name: string): Promise<void> {
  if (SCREENSHOTS) {
    mkdirSync(SCREENSHOTS, { recursive: true });
    await page.screenshot({ path: path.join(SCREENSHOTS, `${name}.png`) });
  }
}

async function open(page: Page, file: string): Promise<{ site: PreviewSite; problems: string[] }> {
  const site = await servePreview(file);
  const problems = watchProblems(page);
  await page.goto(site.url);
  await expectShellLoaded(page, path.basename(file));
  await expect(page.locator('.dl-grid .dl-cell').first()).toBeVisible({ timeout: 60_000 });
  return { site, problems };
}

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

const cellBar = (page: Page) => Promise.all([page.locator('.dl-cellbar-address').textContent(), page.locator('.dl-cellbar-value').getAttribute('title')]);

test('sample.csv: fields as written, a frozen header row, the CSV bar and open as text', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.csv'));
  try {
    await expect(page.locator('#dl-csv-info')).toHaveText('Comma · UTF-8 with BOM · CRLF');
    await expect(page.locator('.dl-tabs')).toBeHidden();
    for (const text of ['Coffee beans', 'Green tea 緑茶', '1,234.50', 'Fair trade, organic', 'Says "crunchy"', '007', '1.50', 'extra field']) {
      await expect(cellWithText(page, text).first()).toBeVisible();
    }
    const header = cellWithText(page, 'Item');
    expect(await style(header, 'font-weight')).toBe('700');
    expect(await style(header, 'background-color')).toBe('rgb(242, 242, 242)');
    expect(await style(cellWithText(page, '007'), 'justify-content')).toBe('flex-end');
    expect(await style(cellWithText(page, 'Coffee beans'), 'justify-content')).toBe('flex-start');
    // The Note column is sized to "Fair trade, organic"; the others keep Excel's default 64px.
    const noteHeader = page.locator('.dl-col-header', { hasText: /^D$/ });
    expect((await box(noteHeader)).width).toBeGreaterThan(130);
    expect((await box(page.locator('.dl-col-header', { hasText: /^A$/ }))).width).toBe(64);

    await cellWithText(page, 'Line one line two').click();
    expect(await cellBar(page)).toEqual(['D3', 'Line one\nline two']);
    await shoot(page, 'csv-sample');

    // The header row stays put while the rest scrolls.
    await page.setViewportSize({ width: 900, height: 140 });
    const before = await box(header);
    await page.locator('.dl-grid').evaluate((grid) => {
      grid.scrollTop = 200;
    });
    await expect(cellWithText(page, 'Coffee beans')).not.toBeInViewport();
    expect((await box(header)).y).toBe(before.y);
    await page.setViewportSize({ width: 1400, height: 900 });

    await page.locator('#dl-open-as-text').click();
    await expect.poll(() => page.evaluate(() => (window as unknown as { previewOpenedAsText?: number }).previewOpenedAsText)).toBe(1);
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('sample.csv: the header row switch', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.csv'));
  try {
    await page.locator('#dl-csv-header').click();
    await expect(page.locator('#dl-csv-header')).toHaveAttribute('aria-pressed', 'false');
    expect(await style(cellWithText(page, 'Item'), 'font-weight')).toBe('400');
    await expect(page.locator('.dl-grid')).not.toHaveClass(/dl-frozen-rows/);
    await shoot(page, 'csv-no-header');
    await page.waitForTimeout(400);
    await page.reload();
    await expect(page.locator('#dl-csv-header')).toHaveAttribute('aria-pressed', 'false');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('Shift_JIS, semicolon and UTF-16 files', async ({ page }) => {
  for (const [fileName, info, text] of [
    ['shift-jis.csv', 'Comma · Shift_JIS · CRLF', '緑茶（ペットボトル）'],
    ['semicolon.csv', 'Semicolon · UTF-8 · LF', '12,50'],
    ['sample.tsv', 'Tab · UTF-16 LE · CRLF', 'Emoji 🎉 and comma, here'],
  ] as const) {
    const { site, problems } = await open(page, await writeSample(fileName));
    try {
      await expect(page.locator('#dl-csv-info')).toHaveText(info);
      await expect(cellWithText(page, text)).toBeVisible();
      await shoot(page, `csv-${fileName.replace('.', '-')}`);
      expect(problems).toEqual([]);
    } finally {
      await site.close();
    }
  }
});

test('a 200,000-row file: only a window of cells exists and the last row is reachable', async ({ page }) => {
  test.setTimeout(120_000);
  const rows = ['id,name,value'];
  for (let i = 1; i <= 200_000; i++) {
    rows.push(`${i},row ${i},${(i * 1.5).toFixed(1)}`);
  }
  const { site, problems } = await open(page, writeTemp('large.csv', rows.join('\n')));
  try {
    expect(await page.locator('.dl-cell').count()).toBeLessThan(5_000);
    await page.locator('.dl-grid').evaluate((grid) => {
      grid.scrollTop = grid.scrollHeight;
    });
    await expect(cellWithText(page, 'row 200000')).toBeInViewport();
    expect(await page.locator('.dl-cell').count()).toBeLessThan(5_000);
    await shoot(page, 'csv-large-bottom');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('a quote that is never closed and fields past the caps are shown with a notice', async ({ page }) => {
  const { site, problems } = await open(page, writeTemp('broken-quote.csv', `a,b\n1,"never closed\n2,3\n${'x'.repeat(40_000)}\n`));
  try {
    await expect(page.locator('#notice')).toContainText('A quote opened in row 2 is never closed: the rest of the file is in that cell.');
    await shoot(page, 'csv-open-quote');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('binary files and workbooks under a .csv name get a clear message', async ({ page }) => {
  const cases: [string, RegExp][] = [
    [writeTemp('image.csv', new Uint8Array(4096).map((_byte, i) => (i % 3 ? 7 : 0))), /^This \.csv file contains binary data, not text$/],
    [writeTemp('renamed.csv', readFileSync(await writeSample('sample.xlsx'))), /^This \.csv file is a ZIP package, probably an \.xlsx workbook/],
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
