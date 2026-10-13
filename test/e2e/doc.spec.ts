// The .doc viewer in Chrome through the browser preview, on synthetic documents from test/support/wordWriter.ts.
// Screenshots go to the test output folder (playwright --output <dir> to keep them elsewhere).
import { expect, test, type Page } from '@playwright/test';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import { writeWordDocument, type WordDocumentInput } from '../support/wordWriter.ts';
import { expectShellLoaded, servePreview, watchProblems, writeSample, type PreviewSite } from './preview-server.ts';

const NOTICE = 'Simplified preview: only the text of this document is shown; formatting, pictures and page layout are left out.';

function writeDoc(fileName: string, input: WordDocumentInput): string {
  const file = path.join(mkdtempSync(path.join(tmpdir(), 'document-lens-doc-')), fileName);
  writeFileSync(file, writeWordDocument(input));
  return file;
}

async function open(page: Page, file: string): Promise<{ site: PreviewSite; problems: string[] }> {
  const site = await servePreview(file);
  const problems = watchProblems(page);
  await page.goto(site.url);
  return { site, problems };
}

async function horizontalOverflow(page: Page): Promise<number> {
  return page.locator('#content').evaluate((content) => content.scrollWidth - content.clientWidth);
}

test('sample.doc: text, Japanese, table, link, footnote and page break on white paper, with the notice', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.doc'));
  try {
    await expectShellLoaded(page, 'sample.doc');
    await expect(page.locator('#notice')).toHaveText(NOTICE);
    const paper = page.locator('#doc-paper');
    await expect(paper.getByText('Document Lens sample', { exact: true })).toBeVisible();
    await expect(paper.getByText('日本語のテキストも表示されます。')).toBeVisible();
    await expect(paper.getByText('Prices: 5 € – “quoted” text.')).toBeVisible();
    await expect(page.locator('.doc-table tr')).toHaveCount(3);
    await expect(page.locator('.doc-table tr').nth(2).locator('td')).toHaveText(['Pears', '5', 'Twoparagraphs']);
    await expect(page.locator('.doc-page-break')).toBeVisible();
    await expect(page.locator('[data-story="footnotes"]')).toContainText('A footnote at the end of the text.');
    expect(await paper.evaluate((element) => getComputedStyle(element).backgroundColor)).toBe('rgb(255, 255, 255)');
    expect(await horizontalOverflow(page)).toBeLessThanOrEqual(0);
    await page.screenshot({ path: test.info().outputPath('doc-sample.png') });

    await page.locator('.doc-link').click();
    expect(await page.evaluate(() => (window as unknown as { previewLinks: string[] }).previewLinks)).toEqual(['https://example.com/']);
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('sample.doc: reloads in place, keeping zoom and scroll position', async ({ page }) => {
  await page.setViewportSize({ width: 1000, height: 400 });
  const { site, problems } = await open(page, await writeSample('sample.doc'));
  try {
    await expectShellLoaded(page, 'sample.doc');
    await page.locator('#zoom-in').click();
    await expect(page.locator('#zoom-value')).toHaveText('110%');
    await page.locator('#content').evaluate((content) => {
      content.scrollTop = 200;
    });
    await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
    await expect(page.locator('#doc-paper')).toContainText('Second page, back to cp1252.');
    await expect(page.locator('#zoom-value')).toHaveText('110%');
    expect(await page.locator('#content').evaluate((content) => content.scrollTop)).toBe(200);
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('sample.doc: fit width fills the view, and narrow views reflow without horizontal scrolling', async ({ page }) => {
  const { site, problems } = await open(page, await writeSample('sample.doc'));
  try {
    await expectShellLoaded(page, 'sample.doc');
    await page.locator('#fit-width').click();
    const zoom = Number.parseInt((await page.locator('#zoom-value').textContent()) ?? '', 10);
    expect(zoom).toBeGreaterThan(150);
    const content = await page.locator('#content').boundingBox();
    const paper = await page.locator('#doc-paper').boundingBox();
    expect(paper && content && paper.width / content.width).toBeGreaterThan(0.9);
    expect(await horizontalOverflow(page)).toBeLessThanOrEqual(0);
    await page.screenshot({ path: test.info().outputPath('doc-fit-width.png') });

    await page.setViewportSize({ width: 420, height: 800 });
    await page.keyboard.press('0');
    await expect(page.locator('#zoom-value')).toHaveText('100%');
    expect(await horizontalOverflow(page)).toBeLessThanOrEqual(0);
    await page.screenshot({ path: test.info().outputPath('doc-narrow.png') });
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('a password-protected .doc shows the error banner', async ({ page }) => {
  const { site, problems } = await open(page, writeDoc('locked.doc', { body: ['Secret\r'], encrypted: true }));
  try {
    await expect(page.locator('#file-name')).toHaveText('locked.doc');
    await expect(page.locator('#error')).toHaveText('Could not display locked.doc: This .doc file is password-protected; Document Lens cannot open encrypted files');
    await expect(page.locator('#notice')).toBeHidden();
    await expect(page.locator('#content')).toBeEmpty();
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('a very long .doc is cut at the character limit with a notice', async ({ page }) => {
  const paragraph = `${'Lorem ipsum dolor sit amet. '.repeat(4)}\r`;
  const file = writeDoc('long.doc', { body: [paragraph.repeat(Math.ceil(3_200_000 / paragraph.length))] });
  const { site, problems } = await open(page, file);
  try {
    await expectShellLoaded(page, 'long.doc');
    await expect(page.locator('#notice')).toHaveText(`${NOTICE} It is very long, so only its beginning is shown.`, { timeout: 30_000 });
    expect(await page.locator('.doc-p').count()).toBe(Math.ceil(3_000_000 / paragraph.length));
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});
