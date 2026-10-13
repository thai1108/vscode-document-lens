// The presentation viewer in Chrome on the synthetic deck (test/fixtures/pptx.ts): what is drawn where, in which
// color, and how the viewer behaves. Set E2E_FILES to also step through every slide of your own .pptx files.
import { expect, test, type Locator, type Page } from '@playwright/test';
import * as path from 'node:path';
import { decodePng } from '../support/pngDecode.ts';
import { expectShellLoaded, servePreview, watchProblems, writeSample, type PreviewSite } from './preview-server.ts';

const SLIDE_WIDTH_PX = 1280;
const emu = (value: number) => value / 9525;

let site: PreviewSite | undefined;
let problems: string[] = [];

test.afterEach(async () => {
  expect(problems).toEqual([]);
  await site?.close();
  site = undefined;
});

async function open(page: Page, file?: string): Promise<void> {
  site = await servePreview(file ?? (await writeSample('sample.pptx')));
  problems = watchProblems(page);
  await page.goto(site.url);
  await expectShellLoaded(page, file ? path.basename(file) : 'sample.pptx');
  await expect(page.locator('#pptx-current-slide')).toBeVisible();
}

async function goTo(page: Page, slide: number, count = 4): Promise<void> {
  await page.locator('.pptx-thumb').nth(slide - 1).click();
  await expect(page.locator('#pptx-counter')).toHaveText(`${slide} / ${count}`);
}

/** Current zoom of the slide, from the frame that holds it. */
async function slideZoom(page: Page): Promise<number> {
  const frame = await page.locator('#pptx-frame').boundingBox();
  return (frame?.width ?? 0) / SLIDE_WIDTH_PX;
}

/** Position of a shape relative to the slide, in slide pixels (unzoomed). */
async function offsetInSlide(page: Page, shape: Locator): Promise<{ x: number; y: number; width: number; height: number }> {
  const [frame, box, zoom] = await Promise.all([page.locator('#pptx-frame').boundingBox(), shape.boundingBox(), slideZoom(page)]);
  if (!frame || !box) {
    throw new Error('not laid out');
  }
  return { x: (box.x - frame.x) / zoom, y: (box.y - frame.y) / zoom, width: box.width / zoom, height: box.height / zoom };
}

async function colorAt(page: Page, element: Locator, fx = 0.5, fy = 0.5): Promise<string> {
  const box = await element.boundingBox();
  if (!box) {
    throw new Error('not laid out');
  }
  const [r, g, b] = decodePng(await page.screenshot()).pixel(box.x + box.width * fx, box.y + box.height * fy);
  return `#${[r, g, b].map((channel) => channel.toString(16).padStart(2, '0')).join('')}`;
}

function expectColor(actual: string, expected: string, tolerance = 3): void {
  const channels = (hex: string) => [1, 3, 5].map((i) => Number.parseInt(hex.slice(i, i + 2), 16));
  const difference = Math.max(...channels(actual).map((value, i) => Math.abs(value - (channels(expected)[i] ?? 0))));
  expect(difference, `${actual} ≈ ${expected}`).toBeLessThanOrEqual(tolerance);
}

const shape = (page: Page, name: string) => page.locator(`#pptx-current-slide [data-name="${name}"]`);

test('sample.pptx: shows the slide list and the first slide fitted to the window', async ({ page }) => {
  await open(page);
  await expect(page.locator('#pptx-counter')).toHaveText('1 / 4');
  await expect(page.locator('.pptx-thumb')).toHaveCount(4);
  await expect(page.locator('.pptx-thumb-frame .pptx-slide')).toHaveCount(4);
  await expect(page.locator('.pptx-thumb').nth(3)).toHaveClass(/pptx-thumb-hidden/);
  await expect(page.locator('#pptx-current-slide')).toContainText('Document Lens サンプル');

  const [stage, frame] = await Promise.all([page.locator('#pptx-stage').boundingBox(), page.locator('#pptx-frame').boundingBox()]);
  expect(frame && stage).toBeTruthy();
  if (frame && stage) {
    // Fitted: the slide fills the stage (less its padding) in one direction and fits in the other.
    const fitsWidth = Math.abs(frame.width - (stage.width - 48)) < 2;
    const fitsHeight = Math.abs(frame.height - (stage.height - 48)) < 2;
    expect(fitsWidth || fitsHeight).toBe(true);
    expect(frame.x).toBeGreaterThanOrEqual(stage.x);
    expect(frame.x + frame.width).toBeLessThanOrEqual(stage.x + stage.width);
  }
  await expect(page.locator('#zoom-value')).toHaveText(`${Math.round((await slideZoom(page)) * 100)}%`);

  // The text is real, selectable text.
  await page.locator('#pptx-current-slide [data-name="Title 1"] .pptx-p').click({ clickCount: 3 });
  expect(await page.evaluate(() => window.getSelection()?.toString().trim())).toBe('Document Lens サンプル');
});

test('sample.pptx: draws shapes at their positions in their theme colors', async ({ page }) => {
  await open(page);
  await goTo(page, 2);
  const rect = await offsetInSlide(page, shape(page, 'Accent rect'));
  expect(rect.x).toBeCloseTo(emu(6000000), 0);
  expect(rect.y).toBeCloseTo(emu(1825625), 0);
  expect(rect.width).toBeCloseTo(emu(1400000), 0);
  expectColor(await colorAt(page, shape(page, 'Accent rect')), '#203864');
  expectColor(await colorAt(page, shape(page, 'Translucent ellipse')), '#80d8a8', 4);
  expectColor(await colorAt(page, shape(page, 'Star')), '#ffc000');
  expectColor(await colorAt(page, shape(page, 'Rounded'), 0.1, 0.5), '#ed7d31');

  // A scaled group: members are twice their size in the group's child space.
  const member = await offsetInSlide(page, shape(page, 'Group B'));
  expect(member.x).toBeCloseTo(emu(9000000 + 1200000), 0);
  expect(member.width).toBeCloseTo(emu(1200000), 0);
  expectColor(await colorAt(page, shape(page, 'Group B')), '#a5a5a5');

  // Rotated by 30°: its bounding box grows beyond the unrotated frame.
  const rotated = await offsetInSlide(page, shape(page, 'Rotated triangle'));
  expect(rotated.height).toBeGreaterThan(emu(1000000) + 10);
  await expect(shape(page, 'Rotated triangle')).toContainText('Rotated');

  // Bullets and numbering from the master's body style.
  const bullets = page.locator('#pptx-current-slide [data-name="Content Placeholder 2"] .pptx-bullet');
  await expect(bullets).toHaveText(['•', '•', '•', '1.', '2.']);
});

test('sample.pptx: decodes and crops pictures, merges table cells and stands in for charts', async ({ page }) => {
  await open(page);
  await goTo(page, 3);
  const picture = shape(page, 'Cropped picture');
  // srcRect l=50%: only the right half (green over yellow) is visible.
  expectColor(await colorAt(page, picture, 0.5, 0.25), '#00ff00');
  expectColor(await colorAt(page, picture, 0.5, 0.75), '#ffff00');
  const natural = await picture.locator('pattern image').evaluate(
    (image) =>
      new Promise<number[]>((resolve, reject) => {
        const probe = new Image();
        probe.onload = () => resolve([probe.naturalWidth, probe.naturalHeight]);
        probe.onerror = () => reject(new Error('picture did not decode'));
        probe.src = image.getAttribute('href') ?? '';
      }),
  );
  expect(natural).toEqual([64, 48]);

  const cells = shape(page, 'Merged table').locator('td');
  await expect(cells).toHaveText(['Merged header', 'Q3', 'Tall', 'A', 'B', 'C', 'D']);
  await expect(cells.nth(0)).toHaveAttribute('colspan', '2');
  await expect(cells.nth(2)).toHaveAttribute('rowspan', '2');
  expectColor(await colorAt(page, cells.nth(1), 0.8, 0.5), '#4472c4');
  expectColor(await colorAt(page, cells.nth(6), 0.8, 0.8), '#c00000');
  const [merged, single] = await Promise.all([cells.nth(0).boundingBox(), cells.nth(1).boundingBox()]);
  expect((merged?.width ?? 0) / (single?.width ?? 1)).toBeCloseTo(2, 1);

  await expect(shape(page, 'Sales chart')).toHaveText('Bar chart');
});

test('sample.pptx: keyboard navigation, hidden slides and slide numbers', async ({ page }) => {
  await open(page);
  await page.locator('#content').focus();
  await expect(shape(page, 'Slide Number')).toHaveText('1');
  await page.keyboard.press('ArrowRight');
  await expect(page.locator('#pptx-counter')).toHaveText('2 / 4');
  await expect(shape(page, 'Slide Number')).toHaveText('2');
  await page.keyboard.press('PageDown');
  await expect(page.locator('#pptx-counter')).toHaveText('3 / 4');
  await expect(page.locator('#pptx-hidden-tag')).toBeHidden();
  await page.keyboard.press('End');
  await expect(page.locator('#pptx-counter')).toHaveText('4 / 4');
  await expect(page.locator('#pptx-hidden-tag')).toBeVisible();
  await expect(page.locator('#pptx-next')).toBeDisabled();
  await expect(page.locator('#pptx-current-slide')).toContainText('Hidden slide');
  expectColor(await colorAt(page, page.locator('#pptx-current-slide'), 0.9, 0.5), '#e7e6e6');
  await page.keyboard.press('PageUp');
  await expect(page.locator('#pptx-counter')).toHaveText('3 / 4');
  await page.keyboard.press('Home');
  await expect(page.locator('#pptx-counter')).toHaveText('1 / 4');
  await page.keyboard.press('ArrowLeft');
  await expect(page.locator('#pptx-counter')).toHaveText('1 / 4');
  await page.locator('#pptx-next').click();
  await expect(page.locator('#pptx-counter')).toHaveText('2 / 4');
});

test('sample.pptx: links open through the host only', async ({ page }) => {
  await open(page);
  const link = page.locator('#pptx-current-slide a.pptx-link');
  await expect(link).toHaveText('example.com/lens');
  await link.click();
  await expect.poll(() => page.evaluate(() => (window as unknown as { previewLinks: string[] }).previewLinks)).toEqual(['https://example.com/lens']);
  expect(page.url()).toBe(site?.url);
});

test('sample.pptx: keeps the slide and zoom across a reload of the file and of the view', async ({ page }) => {
  await open(page);
  await goTo(page, 3);
  await page.locator('#zoom-in').click();
  const zoom = await page.locator('#zoom-value').textContent();
  const before = await shape(page, 'Cropped picture').locator('pattern image').getAttribute('href');
  await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
  await expect(shape(page, 'Cropped picture').locator('pattern image')).not.toHaveAttribute('href', before ?? '');
  await expect(page.locator('#pptx-counter')).toHaveText('3 / 4');
  await expect(page.locator('#zoom-value')).toHaveText(zoom ?? '');
  expectColor(await colorAt(page, shape(page, 'Cropped picture'), 0.5, 0.25), '#00ff00');
  await page.reload();
  await expect(page.locator('#pptx-counter')).toHaveText('3 / 4');
  await expect(page.locator('#zoom-value')).toHaveText(zoom ?? '');
});

test('sample.pptx: follows the window size until the user zooms', async ({ page }) => {
  await open(page);
  const fitted = await slideZoom(page);
  await page.setViewportSize({ width: 1000, height: 700 });
  await expect.poll(() => slideZoom(page)).toBeLessThan(fitted - 0.05);
  await page.locator('#zoom-in').click();
  const zoomed = await page.locator('#zoom-value').textContent();
  await page.setViewportSize({ width: 1400, height: 900 });
  await page.waitForTimeout(200);
  await expect(page.locator('#zoom-value')).toHaveText(zoomed ?? '');
  await page.locator('#fit-width').click();
  await expect.poll(() => slideZoom(page)).toBeCloseTo(fitted, 2);
});

const ownFiles = (process.env['E2E_FILES'] ?? '')
  .split(path.delimiter)
  .filter((file) => file.toLowerCase().endsWith('.pptx'));
for (const file of ownFiles) {
  test(`own file ${path.basename(file)}: every slide renders`, async ({ page }) => {
    test.setTimeout(300_000);
    await open(page, file);
    const count = await page.locator('.pptx-thumb').count();
    await page.locator('#content').focus();
    for (let slide = 1; slide <= count; slide++) {
      await expect(page.locator('#pptx-counter')).toHaveText(`${slide} / ${count}`);
      await expect(page.locator('#pptx-current-slide')).toBeVisible();
      await expect(page.locator('#pptx-current-slide.pptx-slide-error')).toHaveCount(0);
      await page.keyboard.press('ArrowRight');
    }
  });
}
