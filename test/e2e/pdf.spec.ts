// The PDF viewer in Chrome through the browser preview, on the synthetic samples of test/fixtures/pdf.ts. Set
// E2E_FILES to also open your own .pdf files (they are only read locally).
import { expect, test, type Page } from '@playwright/test';
import * as path from 'node:path';
import { SAMPLE_PDF } from '../fixtures/pdf.ts';
import { expectShellLoaded, servePreview, watchProblems, writeSample } from './preview-server.ts';

test.use({ deviceScaleFactor: 2 });

/**
 * Opens a sample (or a file by path) in the preview, expects it to load (or to show `error`) and runs `body`; fails
 * on console, CSP or network errors seen meanwhile.
 */
async function withPdf(page: Page, file: string, body: () => Promise<void>, { error }: { error?: string } = {}): Promise<void> {
  const site = await servePreview(file.includes(path.sep) ? file : await writeSample(file));
  try {
    const problems = watchProblems(page);
    await page.goto(site.url);
    if (error === undefined) {
      await expectShellLoaded(page, path.basename(file));
    } else {
      await expect(page.locator('#file-name')).toHaveText(path.basename(file));
      await expect(page.locator('#error')).toHaveText(error);
    }
    await body();
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
}

const pageView = (page: Page, number: number) => page.locator(`.pdfViewer .page[data-page-number="${number}"]`);
const pageNumber = (page: Page) => page.locator('#pdf-page-number');
const findCount = (page: Page) => page.locator('#pdf-find-count');

async function expectRendered(page: Page, number: number): Promise<void> {
  await expect(pageView(page, number)).toHaveAttribute('data-loaded', 'true');
  await expect(pageView(page, number).locator('.canvasWrapper canvas').first()).toBeVisible();
}

async function goToPage(page: Page, number: number): Promise<void> {
  await pageNumber(page).fill(String(number));
  await pageNumber(page).press('Enter');
  await expect(pageNumber(page)).toHaveValue(String(number));
  await expectRendered(page, number);
}

/** The page's top edge relative to the scroll container's, in CSS pixels. */
async function pageOffset(page: Page, number: number): Promise<number> {
  const scroller = await page.locator('#pdf-scroller').boundingBox();
  const box = await pageView(page, number).boundingBox();
  return (box?.y ?? Number.NaN) - (scroller?.y ?? Number.NaN);
}

/** Red, green and blue of the canvas pixel at a point given in PDF coordinates of the page (origin bottom left). */
async function pixelAt(page: Page, number: number, x: number, y: number, size: number[]): Promise<number[]> {
  return pageView(page, number)
    .locator('.canvasWrapper canvas')
    .first()
    .evaluate(
      (canvas: HTMLCanvasElement, [fx, fy]: [number, number]) => {
        const data = canvas.getContext('2d')?.getImageData(Math.floor(fx * canvas.width), Math.floor(fy * canvas.height), 1, 1).data;
        return data ? [data[0] ?? 0, data[1] ?? 0, data[2] ?? 0] : [];
      },
      [x / (size[0] ?? 1), 1 - y / (size[1] ?? 1)] as [number, number],
    );
}

test('sample.pdf: pages of every size rendered crisp, with text and annotation layers', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expect(page.locator('#pdf-page-count')).toHaveText(`of ${SAMPLE_PDF.pageCount}`);
    await expect(pageNumber(page)).toHaveValue('1');
    await expect(page.locator('.pdfViewer .page')).toHaveCount(SAMPLE_PDF.pageCount);
    await expectRendered(page, 1);

    const textLayer = pageView(page, 1).locator('.textLayer');
    await expect(textLayer).toContainText(SAMPLE_PDF.title);
    // Japanese text in a non-embedded font only decodes with the CMaps embedded in the bundle.
    await expect(textLayer).toContainText(SAMPLE_PDF.japanese);
    await expect(pageView(page, 1).locator('.annotationLayer .linkAnnotation')).toHaveCount(2);
    await expect(pageView(page, 1).locator('.annotationLayer .textAnnotation img')).toHaveCount(2);

    // At 100% a point is 4/3 CSS pixels; the canvas has a device pixel per CSS pixel at device scale 2.
    const canvas = await pageView(page, 1).locator('.canvasWrapper canvas').first().evaluate((element: HTMLCanvasElement) => ({
      width: element.width,
      cssWidth: element.getBoundingClientRect().width,
    }));
    expect(canvas.cssWidth).toBeCloseTo((612 * 4) / 3, 0);
    expect(canvas.width).toBeCloseTo(canvas.cssWidth * 2, -1);

    for (const [index, [width = 0, height = 0]] of SAMPLE_PDF.pageSizes.entries()) {
      const rotated = index === 3;
      await expect
        .poll(async () => {
          const box = await pageView(page, index + 1).boundingBox();
          return [Math.round(box?.width ?? 0), Math.round(box?.height ?? 0)];
        })
        .toEqual(rotated ? [Math.floor((height * 4) / 3), Math.floor((width * 4) / 3)] : [Math.floor((width * 4) / 3), Math.floor((height * 4) / 3)]);
    }

    // The CCITT fax image only shows if the wasm decoder embedded in the bundle ran.
    await goToPage(page, SAMPLE_PDF.image.page);
    const { x, y, size } = SAMPLE_PDF.image;
    const pageSize = SAMPLE_PDF.pageSizes[1] ?? [];
    expect(Math.max(...(await pixelAt(page, 2, x + size * 0.75, y + size / 2, pageSize)))).toBeLessThan(60);
    expect(Math.min(...(await pixelAt(page, 2, x + size * 0.25, y + size / 2, pageSize)))).toBeGreaterThan(200);
  });
});

test('text in the text layer can be selected', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    const title = pageView(page, 1).locator('.textLayer span', { hasText: SAMPLE_PDF.title });
    const box = await title.boundingBox();
    expect(box).not.toBeNull();
    if (box) {
      await page.mouse.move(box.x + 2, box.y + box.height / 2);
      await page.mouse.down();
      await page.mouse.move(box.x + box.width - 2, box.y + box.height / 2, { steps: 5 });
      await page.mouse.up();
    }
    await expect.poll(() => page.evaluate(() => window.getSelection()?.toString().trim())).toMatch(/^Document Lens sampl/);
  });
});

test('find: shortcut, match count, next and previous, highlights, Escape', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    await expect(page.locator('#pdf-find-bar')).toBeHidden();
    await page.locator('#pdf-scroller').click({ position: { x: 5, y: 5 } });
    await page.keyboard.press('ControlOrMeta+f');
    await expect(page.locator('#pdf-find-input')).toBeFocused();

    await page.keyboard.type(SAMPLE_PDF.findTerm);
    await expect(findCount(page)).toHaveText(`1 of ${SAMPLE_PDF.findMatches}`);
    await expect(pageView(page, 1).locator('.textLayer .highlight')).toHaveCount(2);
    await expect(pageView(page, 1).locator('.textLayer .highlight.selected')).toHaveCount(1);
    await page.keyboard.press('Enter');
    await expect(findCount(page)).toHaveText(`2 of ${SAMPLE_PDF.findMatches}`);
    await page.keyboard.press('Enter');
    await expect(findCount(page)).toHaveText(`3 of ${SAMPLE_PDF.findMatches}`);
    await expect(pageNumber(page)).toHaveValue('2');
    await page.keyboard.press('Shift+Enter');
    await expect(findCount(page)).toHaveText(`2 of ${SAMPLE_PDF.findMatches}`);
    await page.locator('#pdf-find-next').click();
    await page.locator('#pdf-find-next').click();
    await expect(findCount(page)).toHaveText(`4 of ${SAMPLE_PDF.findMatches}`);
    await expect(pageNumber(page)).toHaveValue('4');

    await page.locator('#pdf-find-input').fill('日本語');
    await expect(findCount(page)).toHaveText('1 of 1');
    await expect(pageNumber(page)).toHaveValue('1');
    await page.locator('#pdf-find-input').fill('no such words');
    await expect(findCount(page)).toHaveText('No results');

    await page.keyboard.press('Escape');
    await expect(page.locator('#pdf-find-bar')).toBeHidden();
    await expect(page.locator('#pdf-find-toggle')).toBeVisible();
    await expect(page.locator('.textLayer .highlight')).toHaveCount(0);
    await expect(page.locator('#pdf-scroller')).toBeFocused();

    // The toolbar button opens it again with the last query.
    await page.locator('#pdf-find-toggle').click();
    await expect(page.locator('#pdf-find-input')).toHaveValue('no such words');
  });
});

test('page navigation with the buttons and the page number field', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    await expect(page.locator('#pdf-previous-page')).toBeDisabled();
    await page.locator('#pdf-next-page').click();
    await expect(pageNumber(page)).toHaveValue('2');
    await expect.poll(() => pageOffset(page, 2)).toBeLessThan(2);
    await goToPage(page, 4);
    await expect(page.locator('#pdf-next-page')).toBeDisabled();
    await page.locator('#pdf-previous-page').click();
    await expect(pageNumber(page)).toHaveValue('3');

    // Numbers out of range are clamped; anything else puts the current page back.
    await pageNumber(page).fill('99');
    await pageNumber(page).press('Enter');
    await expect(pageNumber(page)).toHaveValue('4');
    await pageNumber(page).fill('abc');
    await pageNumber(page).press('Enter');
    await expect(pageNumber(page)).toHaveValue('4');
  });
});

test('zoom re-renders at the new scale on the same page; fit width fills the view', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    await goToPage(page, 3);
    const before = await pageOffset(page, 3);

    await page.locator('#zoom-in').click();
    await expect(page.locator('#zoom-value')).toHaveText('110%');
    await expect.poll(async () => Math.round((await pageView(page, 3).boundingBox())?.width ?? 0)).toBe(Math.floor((842 * 4 * 1.1) / 3));
    await expect(pageNumber(page)).toHaveValue('3');
    expect(Math.abs((await pageOffset(page, 3)) - before)).toBeLessThan(4);
    await expect
      .poll(() =>
        pageView(page, 3)
          .locator('.canvasWrapper canvas')
          .first()
          .evaluate((canvas: HTMLCanvasElement) => Math.abs(canvas.width - canvas.getBoundingClientRect().width * 2)),
      )
      .toBeLessThan(4);

    await page.locator('#fit-width').click();
    const scroller = await page.locator('#pdf-scroller').evaluate((element) => element.clientWidth);
    await expect.poll(async () => Math.round((await pageView(page, 3).boundingBox())?.width ?? 0)).toBeGreaterThanOrEqual(scroller - 24 - 2);
    expect(Math.round((await pageView(page, 3).boundingBox())?.width ?? 0)).toBeLessThanOrEqual(scroller - 24);
    await expect(pageNumber(page)).toHaveValue('3');
  });
});

test('a first look in a narrow editor starts at fit width; a reopened file keeps its zoom', async ({ page }) => {
  await page.setViewportSize({ width: 700, height: 800 });
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    const scroller = await page.locator('#pdf-scroller').evaluate((element) => element.clientWidth);
    await expect.poll(async () => Math.round((await pageView(page, 1).boundingBox())?.width ?? 0)).toBeLessThanOrEqual(scroller - 24);
    const fitted = await page.locator('#zoom-value').textContent();
    expect(Number.parseInt(fitted ?? '', 10)).toBeLessThan(100);

    await page.locator('#zoom-in').click();
    const zoomed = await page.locator('#zoom-value').textContent();
    await expect.poll(() => page.evaluate(() => JSON.parse(sessionStorage.getItem('document-lens-state') ?? '{}').viewer?.page)).toBe(1);
    await page.reload();
    await expectRendered(page, 1);
    await expect(page.locator('#zoom-value')).toHaveText(zoomed ?? '');
  });
});

test('links: http(s) through the host, internal destinations and the outline scroll to their page', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    await pageView(page, 1).locator(`.linkAnnotation a[href="${SAMPLE_PDF.externalUrl}"]`).click();
    await expect.poll(() => page.evaluate(() => (window as unknown as { previewLinks: string[] }).previewLinks)).toEqual([SAMPLE_PDF.externalUrl]);

    // Named destination.
    await pageView(page, 1).locator('.linkAnnotation a[href^="#"]').click();
    await expect(pageNumber(page)).toHaveValue('3');
    await expectRendered(page, 3);
    // Explicit destination.
    await pageView(page, 3).locator('.linkAnnotation a').click();
    await expect(pageNumber(page)).toHaveValue('4');

    await expect(page.locator('#pdf-outline')).toBeHidden();
    await page.locator('#pdf-outline-toggle').click();
    await expect(page.locator('#pdf-outline-toggle')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.locator('#pdf-outline .pdf-outline-item')).toHaveText(SAMPLE_PDF.outlineTitles);
    await expect(page.locator('#pdf-outline .pdf-outline-item', { hasText: 'Rotated page' })).toHaveAttribute('aria-level', '2');
    await page.locator('#pdf-outline .pdf-outline-item', { hasText: 'Images and symbols' }).click();
    await expect(pageNumber(page)).toHaveValue('2');
    await page.locator('#pdf-outline .pdf-outline-item', { hasText: 'Project website' }).click();
    await expect
      .poll(() => page.evaluate(() => (window as unknown as { previewLinks: string[] }).previewLinks))
      .toEqual([SAMPLE_PDF.externalUrl, SAMPLE_PDF.outlineUrl]);
    await page.locator('#pdf-outline-toggle').click();
    await expect(page.locator('#pdf-outline')).toBeHidden();
  });
});

test('a reload of the file keeps the page, zoom, outline and find query; reopening restores them', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    await goToPage(page, 3);
    await page.locator('#zoom-in').click();
    await page.locator('#pdf-outline-toggle').click();
    await page.keyboard.press('ControlOrMeta+f');
    await page.keyboard.type(SAMPLE_PDF.findTerm);
    // The search starts at the current page: the next match after page 3 is the one on page 4.
    await expect(findCount(page)).toHaveText(`4 of ${SAMPLE_PDF.findMatches}`);
    await expect(pageNumber(page)).toHaveValue('4');
    await goToPage(page, 3);

    await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
    await expect(page.locator('.pdfViewer .page')).toHaveCount(SAMPLE_PDF.pageCount);
    await expectRendered(page, 3);
    await expect(pageNumber(page)).toHaveValue('3');
    await expect.poll(() => pageOffset(page, 3)).toBeLessThan(2);
    await expect(page.locator('#zoom-value')).toHaveText('110%');
    await expect(page.locator('#pdf-outline')).toBeVisible();
    // The search is not re-run on its own (it would scroll); its query waits for Enter.
    await expect(page.locator('#pdf-find-input')).toHaveValue(SAMPLE_PDF.findTerm);
    await expect(findCount(page)).toHaveText('');
    await page.waitForTimeout(500);
    await expect(pageNumber(page)).toHaveValue('3');

    await expect
      .poll(() => page.evaluate(() => JSON.parse(sessionStorage.getItem('document-lens-state') ?? '{}').viewer))
      .toMatchObject({ page: 3, outline: true });
    await page.reload();
    await expectShellLoaded(page, 'sample.pdf');
    await expectRendered(page, 3);
    await expect(pageNumber(page)).toHaveValue('3');
    await expect(page.locator('#zoom-value')).toHaveText('110%');
    await expect(page.locator('#pdf-outline')).toBeVisible();
  });
});

test('a reload that cannot be read keeps the last good pages until the next good reload', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    await goToPage(page, 2);
    // Swap the bytes the preview's mock host sends (same length, as its serialized message expects) for a file with
    // a PDF header and nothing behind it.
    const good = await page.evaluate(() => {
      const preview = (window as unknown as { previewMessage: { buffers: string[] } }).previewMessage;
      const bytes = preview.buffers[0] ?? '';
      preview.buffers[0] = btoa('%PDF-1.7\nhalf-written'.padEnd(atob(bytes).length, ' '));
      (window as unknown as { previewReload(): void }).previewReload();
      return bytes;
    });
    await expect(page.locator('#error')).toHaveText('Could not display sample.pdf: Not a valid .pdf file: Invalid PDF structure.');
    await expect(page.locator('.pdfViewer .page')).toHaveCount(SAMPLE_PDF.pageCount);
    await expect(pageNumber(page)).toHaveValue('2');

    await page.evaluate((bytes) => {
      (window as unknown as { previewMessage: { buffers: string[] } }).previewMessage.buffers[0] = bytes;
      (window as unknown as { previewReload(): void }).previewReload();
    }, good);
    await expect(page.locator('#error')).toBeHidden();
    await expectRendered(page, 2);
    await expect(pageNumber(page)).toHaveValue('2');
  });
});

test('a reload while the view is hidden (a background tab) keeps the page once it is shown again', async ({ page }) => {
  await withPdf(page, 'sample.pdf', async () => {
    await expectRendered(page, 1);
    await goToPage(page, 3);
    await page.evaluate(() => {
      (document.getElementById('app') as HTMLElement).hidden = true;
      (window as unknown as { previewReload(): void }).previewReload();
    });
    await page.waitForTimeout(500);
    await page.evaluate(() => {
      (document.getElementById('app') as HTMLElement).hidden = false;
    });
    await expectRendered(page, 3);
    await expect(pageNumber(page)).toHaveValue('3');
    await expect.poll(() => pageOffset(page, 3)).toBeLessThan(2);
  });
});

test('a 600-page document only renders the pages near the view', async ({ page }) => {
  await withPdf(page, 'long.pdf', async () => {
    await expect(page.locator('#pdf-page-count')).toHaveText(`of ${SAMPLE_PDF.longPageCount}`);
    await expect(page.locator('.pdfViewer .page')).toHaveCount(SAMPLE_PDF.longPageCount);
    await expectRendered(page, 1);
    for (const number of [100, 200, 300, 400, 500, 600]) {
      await goToPage(page, number);
      await expect(pageView(page, number).locator('.textLayer')).toContainText(`Long document page ${number}`);
    }
    // pdf.js keeps the canvases of the 10 most recently drawn pages and frees the others.
    expect(await page.locator('.canvasWrapper canvas').count()).toBeLessThanOrEqual(10);
    await expect(pageView(page, 1)).not.toHaveAttribute('data-loaded', 'true');
    await expect(pageView(page, 1).locator('canvas')).toHaveCount(0);
  });
});

const UNREADABLE = [
  ['encrypted.pdf', 'Could not display encrypted.pdf: This .pdf file is password-protected; Document Lens cannot open encrypted files'],
  ['broken.pdf', 'Could not display broken.pdf: Not a valid .pdf file: Invalid PDF structure.'],
] as const;

for (const [fileName, message] of UNREADABLE) {
  test(`${fileName}: the error banner instead of pages`, async ({ page }) => {
    await withPdf(
      page,
      fileName,
      async () => {
        await expect(page.locator('#pdf-pages')).toHaveCount(0);
        await expect(page.locator('#pdf-navigation')).toHaveCount(0);
      },
      { error: message },
    );
  });
}

const ownFiles = (process.env['E2E_FILES'] ?? '').split(path.delimiter).filter((file) => file.toLowerCase().endsWith('.pdf'));
for (const file of ownFiles) {
  test(`own file ${path.basename(file)}: the first page renders with its text layer`, async ({ page }) => {
    await withPdf(page, file, async () => {
      await expectRendered(page, 1);
      await expect(pageView(page, 1).locator('.textLayer')).toBeAttached();
    });
  });
}
