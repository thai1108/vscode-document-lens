// The .ppt viewer in Chrome through the browser preview: slide cards from the synthetic sample (text, Japanese,
// picture, hidden slide), reloads, zoom, and the error banner for files it cannot show. .ppt files listed in E2E_FILES
// are checked too (locally only; nothing is saved from them).
import { expect, test, type Page } from '@playwright/test';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import { SAMPLE_PPT } from '../fixtures/ppt.ts';
import { pptFile, pptStreams, writePpt } from '../support/pptWriter.ts';
import { writeZip } from '../support/zipWriter.ts';
import { expectShellLoaded, servePreview, watchProblems, writeSample } from './preview-server.ts';

const SIMPLIFIED_NOTICE = 'Simplified preview: only the text and pictures of each slide are shown, without layout, shapes or formatting.';

function writeFile(name: string, bytes: Uint8Array): string {
  const file = path.join(mkdtempSync(path.join(tmpdir(), 'document-lens-ppt-')), name);
  writeFileSync(file, bytes);
  return file;
}

/** Scrolls the picture into view (they load lazily) and returns its decoded width, or 0 if it failed. */
async function decodedWidth(page: Page, selector: string): Promise<number> {
  const image = page.locator(selector);
  await image.scrollIntoViewIfNeeded();
  await expect.poll(() => image.evaluate((element: HTMLImageElement) => element.complete)).toBe(true);
  return image.evaluate((element: HTMLImageElement) => element.naturalWidth);
}

async function openFile(page: Page, file: string, check: (problems: string[]) => Promise<void>): Promise<void> {
  const site = await servePreview(file);
  try {
    const problems = watchProblems(page);
    await page.goto(site.url);
    await check(problems);
  } finally {
    await site.close();
  }
}

test('sample.ppt: one card per slide with its text, Japanese text, picture and hidden mark', async ({ page }, testInfo) => {
  await openFile(page, await writeSample('sample.ppt'), async (problems) => {
    await expectShellLoaded(page, 'sample.ppt');
    await expect(page.locator('.ppt-slide')).toHaveCount(3);
    await expect(page.locator('#notice')).toHaveText(SIMPLIFIED_NOTICE);
    await expect(page.locator('#ppt-slide-count')).toHaveText('3 slides · 1 hidden');

    await expect(page.locator('#slide-1 .ppt-slide-number')).toHaveText('Slide 1');
    await expect(page.locator('#slide-1 h2.ppt-title')).toHaveText('Quarterly review');
    await expect(page.locator('#slide-1 .ppt-text p')).toHaveText(['四半期レビュー', 'Document Lens sample']);
    await expect(page.locator('#slide-2 h2.ppt-title')).toHaveText('Highlights');
    await expect(page.locator('#slide-2 .ppt-text-body li')).toHaveText(['Revenue grew 12%', 'New customers in Osaka', 'Roadmap\nsecond line', '東京オフィス開設']);
    await expect(page.locator('#slide-2 .ppt-text-other')).toHaveText('注: 速報値です');

    // Japanese text is laid out with real glyphs, as wide as the Latin title at the same size or wider per character.
    const japanese = await page.locator('#slide-1 .ppt-text p').first().evaluate((element) => element.getBoundingClientRect().width);
    expect(japanese).toBeGreaterThan(50);

    // Nested bullets are indented.
    const indents = await page.locator('#slide-2 .ppt-text-body li').evaluateAll((items) => items.map((item) => item.getBoundingClientRect().left));
    expect(indents[1]).toBeGreaterThan(indents[0] ?? Infinity);

    await expect(page.locator('#slide-2 img')).toHaveAttribute('src', /^blob:/);
    expect(await decodedWidth(page, '#slide-2 img')).toBe(120);

    await expect(page.locator('#slide-3')).toHaveClass(/ppt-slide-hidden/);
    await expect(page.locator('#slide-3 .ppt-badge')).toHaveText('Hidden');
    await expect(page.locator('#slide-2 .ppt-badge')).toHaveCount(0);

    // Cards keep the 16:9 slide shape.
    const card = await page.locator('#slide-3 .ppt-slide-card').boundingBox();
    expect((card?.width ?? 0) / (card?.height ?? 1)).toBeCloseTo(16 / 9, 1);

    await page.locator('#content').evaluate((element) => element.scrollTo(0, 0));
    await page.screenshot({ path: testInfo.outputPath('sample-top.png') });
    for (const id of ['slide-1', 'slide-2', 'slide-3']) {
      await page.locator(`#${id}`).screenshot({ path: testInfo.outputPath(`sample-${id}.png`) });
    }

    // The host re-sends the file when it changes on disk.
    await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
    await expect(page.locator('.ppt-slide')).toHaveCount(3);
    expect(await decodedWidth(page, '#slide-2 img')).toBe(120);
    await expect(page.locator('#error')).toBeHidden();
    expect(problems).toEqual([]);
  });
});

test('sample.ppt: zoom scales the slide cards', async ({ page }) => {
  await openFile(page, await writeSample('sample.ppt'), async (problems) => {
    await expectShellLoaded(page, 'sample.ppt');
    const before = await page.locator('#slide-1 .ppt-slide-card').boundingBox();
    await page.locator('#content').focus();
    await page.keyboard.press('-');
    await page.keyboard.press('-');
    await expect(page.locator('#zoom-value')).toHaveText('80%');
    const after = await page.locator('#slide-1 .ppt-slide-card').boundingBox();
    expect(after?.width).toBeCloseTo((before?.width ?? 0) * 0.8, 0);
    expect(problems).toEqual([]);
  });
});

const UNREADABLE: [string, () => Uint8Array, string][] = [
  ['encrypted.ppt', () => writePpt({ ...SAMPLE_PPT, encrypted: { token: true } }), 'This .ppt file is password-protected; Document Lens cannot open encrypted files'],
  ['old.ppt', () => pptFile({ PP40: new Uint8Array(6000), 'Current User': new Uint8Array(12) }), 'This .ppt file was saved by PowerPoint 95 or older, a format too old for Document Lens to read'],
  [
    'damaged.ppt',
    () => {
      const streams = pptStreams(SAMPLE_PPT);
      streams.document.writeUInt32LE(0x7fffffff, (streams.offsets.get(1) ?? 0) + 4);
      return pptFile({ 'PowerPoint Document': streams.document, 'Current User': streams.currentUser, Pictures: streams.pictures });
    },
    'This .ppt file is damaged: its document record could not be found',
  ],
];

for (const [fileName, build, message] of UNREADABLE) {
  test(`${fileName}: shows why it cannot be displayed`, async ({ page }) => {
    await openFile(page, writeFile(fileName, build()), async (problems) => {
      await expect(page.locator('#file-name')).toHaveText(fileName);
      await expect(page.locator('#error')).toHaveText(`Could not display ${fileName}: ${message}`);
      await expect(page.locator('.ppt-slide')).toHaveCount(0);
      expect(problems).toEqual([]);
    });
  });
}

test('a .pptx renamed to .ppt is refused by the host check', async ({ page }) => {
  await openFile(page, writeFile('renamed.ppt', writeZip([{ name: '[Content_Types].xml', data: '<Types/>' }])), async (problems) => {
    await expect(page.locator('#error')).toHaveText('This .ppt file is a ZIP package, probably a .pptx file with the wrong extension');
    expect(problems).toEqual([]);
  });
});

const ownFiles = (process.env['E2E_FILES'] ?? '').split(path.delimiter).filter((file) => file.toLowerCase().endsWith('.ppt'));
for (const file of ownFiles) {
  test(`own file ${path.basename(file)}: slides render and pictures decode`, async ({ page }) => {
    await openFile(page, file, async (problems) => {
      await expectShellLoaded(page, path.basename(file));
      await expect(page.locator('#notice')).toContainText(SIMPLIFIED_NOTICE);
      const slides = await page.locator('.ppt-slide').count();
      expect(slides).toBeGreaterThan(0);
      const images = await page.locator('.ppt-picture img').count();
      for (let i = 0; i < images; i++) {
        expect(await decodedWidth(page, `.ppt-picture img >> nth=${i}`), `picture ${i + 1}`).toBeGreaterThan(0);
      }
      expect(problems).toEqual([]);
      console.log(`${path.basename(file)}: ${slides} slides, ${images} pictures decoded`);
    });
  });
}
