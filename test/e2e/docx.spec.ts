// The .docx viewer in Chrome through the browser preview: rendering of the synthetic samples (test/fixtures/docx.ts),
// links, zoom and reloads, the hostile sample, and broken files. Set E2E_FILES to also render your own .docx files.
import { expect, test, type Page } from '@playwright/test';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import { MAX_XML_PART_SIZE } from '../../src/limits.ts';
import { FONT_INJECTION, SAMPLE_LINK } from '../fixtures/docx.ts';
import { compoundFileHeaderStub } from '../support/compoundFile.ts';
import { png } from '../support/png.ts';
import { CONTENT_TYPE, inlinePicture, paragraph, RELATIONSHIP, run, WORD_NAMESPACES, wordPackage } from '../support/wordPackage.ts';
import { junkDeflateEntry, markDeflated, prependedDecoy } from '../support/zipTamper.ts';
import { expectShellLoaded, servePreview, watchProblems, writeSample, type PreviewSite } from './preview-server.ts';

/** The page gap in docx.css, in CSS px at 100%. */
const PAGE_GAP = 24;
/** DOCX_LIMITS.maxLengthPt in CSS px. */
const MAX_LENGTH_PX = (5000 * 4) / 3;

const pages = (page: Page) => page.locator('#content .docx-wrapper > section.docx');

function writeFile(name: string, data: Uint8Array): string {
  const file = path.join(mkdtempSync(path.join(tmpdir(), 'document-lens-docx-')), name);
  writeFileSync(file, data);
  return file;
}

/** Opens `file` in the preview and waits until its pages are shown (failing with the error banner's text if not). */
async function openDocument(page: Page, file: string, timeout = 10_000): Promise<{ site: PreviewSite; problems: string[] }> {
  const site = await servePreview(file);
  const problems = watchProblems(page);
  await page.goto(site.url);
  await expectShellLoaded(page, path.basename(file));
  await expect(page.locator('#content .docx-wrapper > section.docx, #error:not([hidden])').first()).toBeVisible({ timeout });
  await expect(page.locator('#error')).toBeHidden();
  return { site, problems };
}

const scrollTop = (page: Page) => page.locator('#content').evaluate((content) => content.scrollTop);

test.describe('.docx viewer', () => {
  let site: PreviewSite | undefined;
  test.afterEach(async () => {
    await site?.close();
    site = undefined;
  });

  test('lays the sample out as pages with headers, footers, styles, lists, a table, a picture, Japanese text and its font', async ({ page }) => {
    const opened = await openDocument(page, await writeSample('sample.docx'));
    site = opened.site;
    await expect(pages(page)).toHaveCount(2);
    for (const sheet of await pages(page).all()) {
      await expect(sheet.locator('header')).toHaveText('Document Lens sample · header');
      await expect(sheet.locator('footer')).toHaveText('Footer · フッター');
      expect(await sheet.evaluate((element) => getComputedStyle(element).backgroundColor)).toBe('rgb(255, 255, 255)');
    }
    // A4 at 100%: 595.3 pt wide.
    expect((await pages(page).first().boundingBox())?.width).toBeCloseTo((11906 / 20) * (4 / 3), 0);

    const content = page.locator('#content');
    const heading = content.getByText('Formatting', { exact: true });
    expect(await heading.evaluate((element) => [getComputedStyle(element).color, getComputedStyle(element).fontWeight])).toEqual(['rgb(47, 84, 150)', '700']);
    expect(await content.getByText('bold', { exact: true }).evaluate((element) => getComputedStyle(element).fontWeight)).toBe('700');
    expect(await content.getByText('italic', { exact: true }).evaluate((element) => getComputedStyle(element).fontStyle)).toBe('italic');
    expect(await content.getByText('red', { exact: true }).evaluate((element) => getComputedStyle(element).color)).toBe('rgb(192, 0, 0)');
    await expect(content.getByText('これは日本語のテキストです。文書をそのまま表示します。')).toBeVisible();
    await expect(content.locator('ol li')).toHaveText('A footnote at the foot of the page.');

    const marker = (text: string) =>
      content.locator('p', { hasText: text }).evaluate((element) => getComputedStyle(element, '::before').content);
    expect(await marker('First bullet')).toContain('•');
    expect(await marker('Second step')).toContain('counter(');

    await expect(content.locator('td[colspan="2"]')).toHaveText('Merged across');
    await expect(content.locator('td[rowspan="2"]')).toHaveText('Merged down');
    // The cell under the vertical merge is hidden.
    await expect(content.locator('td:visible')).toHaveCount(7);

    const picture = content.locator('img');
    await expect(picture).toHaveCount(1);
    expect(await picture.evaluate(async (image: HTMLImageElement) => {
      await image.decode();
      return [image.src.startsWith('blob:'), image.naturalWidth, image.naturalHeight, image.getBoundingClientRect().width];
    })).toEqual([true, 200, 100, 200]);

    // The document's style sheets carry the nonce and stay inside the document.
    const sheets = await page.locator('head style').evaluateAll((styles) => styles.map((style) => [(style as HTMLStyleElement).nonce !== '', style.textContent ?? '']));
    expect(sheets.length).toBeGreaterThan(0);
    for (const [nonce, text] of sheets) {
      expect(nonce).toBe(true);
      expect(text).toMatch(/^(@font-face|@scope \(\.dl-docx\))/);
    }
    // The embedded font is declared from a blob: URL, under a name of the viewer's own.
    expect(sheets.map(([, text]) => text).join('\n')).toMatch(/@font-face \{ font-family: dl-docx-font-1; src: url\("blob:/);
    expect(opened.problems).toEqual([]);
  });

  test('opens web links through the host and scrolls to bookmarks', async ({ page }) => {
    const opened = await openDocument(page, await writeSample('sample.docx'));
    site = opened.site;
    const url = page.url();
    await page.locator('#content').getByText('Document Lens on GitHub').click();
    expect(await page.evaluate(() => (window as unknown as { previewLinks: string[] }).previewLinks)).toEqual([SAMPLE_LINK]);

    const details = page.locator('#content').getByText('Details', { exact: true });
    await expect(details).not.toBeInViewport();
    await page.locator('#content').getByText('Jump to the details').click();
    await expect(details).toBeInViewport();
    expect(await scrollTop(page)).toBeGreaterThan(500);
    expect(page.url()).toBe(url);
    expect(opened.problems).toEqual([]);
  });

  test('zooms around the view, fits the page width and keeps the position across reloads', async ({ page }) => {
    const opened = await openDocument(page, await writeSample('sample.docx'));
    site = opened.site;
    const content = page.locator('#content');

    await page.locator('#fit-width').click();
    const sheet = await pages(page).first().boundingBox();
    const area = await content.evaluate((element) => ({ left: element.getBoundingClientRect().left, width: element.clientWidth }));
    const fitted = { pageWidth: sheet?.width ?? 0, available: area.width, left: (sheet?.x ?? 0) - area.left };
    const zoom = Number((await page.locator('#zoom-value').textContent())?.replace('%', '')) / 100;
    expect(fitted.pageWidth + 2 * PAGE_GAP * zoom).toBeCloseTo(fitted.available, -1);
    expect(fitted.left).toBeCloseTo(PAGE_GAP * zoom, -1);

    await page.locator('#zoom-value').click();
    await content.evaluate((element) => {
      element.scrollTop = 900;
    });
    const before = await scrollTop(page);
    // The point at the centre of the view stays there when zooming in.
    const centre = await content.evaluate((element) => element.scrollTop + element.clientHeight / 2);
    await page.locator('#zoom-in').click();
    await expect(page.locator('#zoom-value')).toHaveText('110%');
    expect(await content.evaluate((element) => element.scrollTop + element.clientHeight / 2)).toBeCloseTo(centre * 1.1, -1);
    await page.locator('#zoom-out').click();
    await expect(page.locator('#zoom-value')).toHaveText('100%');
    expect(await scrollTop(page)).toBeCloseTo(before, -1);

    const styles = await page.locator('head style').count();
    await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
    await expect(pages(page)).toHaveCount(2);
    expect(await scrollTop(page)).toBeCloseTo(before, -1);
    expect(await page.locator('head style').count()).toBe(styles);

    // The webview itself reloading (e.g. after VS Code restarts) restores the position from the saved state.
    const savedTop = () => page.evaluate(() => (JSON.parse(sessionStorage.getItem('document-lens-state') ?? '{}') as { viewer?: { top?: number } }).viewer?.top);
    await expect.poll(savedTop).toBeCloseTo(before, -1);
    await page.reload();
    await expect(pages(page)).toHaveCount(2);
    await expect.poll(() => scrollTop(page)).toBeCloseTo(before, -1);
    expect(opened.problems).toEqual([]);
  });

  test('hostile sample: loads nothing from outside, neutralises links and clamps sizes', async ({ page }) => {
    const requests: string[] = [];
    page.on('request', (request) => requests.push(request.url()));
    const opened = await openDocument(page, await writeSample('hostile.docx'));
    site = opened.site;
    const origin = new URL(site.url).origin;
    for (const url of requests) {
      expect(url.startsWith(`${origin}/`) || url.startsWith('blob:') || url.startsWith('data:'), url).toBe(true);
    }

    const content = page.locator('#content');
    await expect(content.getByText('script link', { exact: true })).toBeVisible();
    const hrefs = await content.locator('a[href]').evaluateAll((links) => links.map((link) => link.getAttribute('href')));
    expect(hrefs).toEqual(['#dl-docx-error']);
    expect(await content.locator('[id]').evaluateAll((elements) => elements.map((element) => element.id))).toEqual(['dl-docx-error', 'dl-docx-content']);
    await expect(page.locator('#error')).toHaveCount(1);

    await expect(content.locator('.dl-docx-image-placeholder')).toHaveCount(2);
    expect(await content.locator('img').evaluateAll((images) => images.map((image) => image.getAttribute('src')?.startsWith('blob:')))).toEqual([true]);
    expect(await content.locator('svg image').evaluateAll((images) => images.map((image) => image.getAttribute('href')))).toEqual([null]);
    await expect(content.locator('iframe, script, object, embed')).toHaveCount(0);

    const sizes = await content.evaluate((element) => {
      const widths = [...element.querySelectorAll('section, img, svg, table')].map((node) => node.getBoundingClientRect().width);
      const fontSizes = [...element.querySelectorAll('span')].map((node) => parseFloat(getComputedStyle(node).fontSize));
      return { widest: Math.max(...widths), largestFont: Math.max(...fontSizes) };
    });
    expect(sizes.widest).toBeLessThanOrEqual(MAX_LENGTH_PX + 1);
    expect(sizes.largestFont).toBeLessThanOrEqual(MAX_LENGTH_PX + 1);

    // Injected rules cannot reach the viewer's own elements, nor move the document over them.
    expect(await page.locator('head style').evaluateAll((styles) => styles.some((style) => style.textContent?.includes(':scope')))).toBe(true);
    await expect(page.locator('#toolbar')).toBeVisible();
    expect(await page.locator('body').evaluate((body) => getComputedStyle(body).display)).toBe('block');
    const toolbar = await page.locator('#toolbar').boundingBox();
    const topmost = await page.evaluate(([x, y]) => document.elementFromPoint(x ?? 0, y ?? 0)?.closest('#toolbar') !== null, [
      (toolbar?.x ?? 0) + (toolbar?.width ?? 0) / 2,
      (toolbar?.y ?? 0) + (toolbar?.height ?? 0) / 2,
    ]);
    expect(topmost).toBe(true);
    const frame = await page.locator('.dl-docx-frame').evaluate((element) => [getComputedStyle(element).position, getComputedStyle(element).contain]);
    expect(frame).toEqual(['static', 'layout']);

    // var() substitutions from the theme neither load anything nor make huge lengths.
    const css = (await page.locator('head style').allTextContents()).join('\n');
    expect(css).not.toMatch(/example\.invalid|var\(--[uks]\)|--u:|--k:/);
    const margin = await content.locator('section.docx').first().evaluate((section) => parseFloat(getComputedStyle(section).marginLeft));
    expect(margin).toBeLessThanOrEqual(MAX_LENGTH_PX + 1);
    // The embedded font is renamed, so it cannot stand in for the UI font it claims to be.
    expect(FONT_INJECTION).toContain('BlinkMacSystemFont');
    expect(css).toMatch(/@font-face \{ font-family: dl-docx-font-1; src: url\("blob:[^"]+"\); \}/);
    expect(css).not.toMatch(/BlinkMacSystemFont|size-adjust/);
    expect(opened.problems).toEqual([]);
  });

  test('releases the previous render’s pictures and fonts when the file reloads', async ({ page }) => {
    await page.addInitScript(() => {
      const revoke = URL.revokeObjectURL.bind(URL);
      const revoked: string[] = [];
      Object.assign(window, { revoked });
      URL.revokeObjectURL = (url: string) => {
        revoked.push(url);
        revoke(url);
      };
    });
    const opened = await openDocument(page, await writeSample('sample.docx'));
    site = opened.site;
    const objectUrls = () =>
      page.evaluate(() => {
        const urls = [...document.querySelectorAll('#content img')].map((image) => image.getAttribute('src') ?? '');
        const styles = [...document.querySelectorAll('head style')].map((style) => style.textContent ?? '');
        return [...urls, ...styles.flatMap((text) => text.match(/blob:[^\s"')]+/g) ?? [])].sort();
      });
    const before = await objectUrls();
    expect(before).toHaveLength(2);
    await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
    await expect.poll(objectUrls).not.toEqual(before);
    expect((await page.evaluate(() => (window as unknown as { revoked: string[] }).revoked)).sort()).toEqual(before);
    expect(opened.problems).toEqual([]);
  });

  test('makes one URL per picture however often it is used, and releases it when the file reloads', async ({ page }) => {
    await page.addInitScript(() => {
      const create = URL.createObjectURL.bind(URL);
      const revoke = URL.revokeObjectURL.bind(URL);
      const created: string[] = [];
      const revoked: string[] = [];
      Object.assign(window, { created, revoked });
      URL.createObjectURL = (blob: Blob | MediaSource) => {
        const url = create(blob);
        created.push(url);
        return url;
      };
      URL.revokeObjectURL = (url: string) => {
        revoked.push(url);
        revoke(url);
      };
    });
    // 200 uses of one picture, some in superscript runs (which docx-preview renders twice, keeping one copy).
    const use = inlinePicture('rIdPicture', 9144 * 20, 9144 * 20);
    const superscript = use.replace('<w:r>', '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>');
    const body = Array.from({ length: 100 }, () => paragraph(use + superscript)).join('');
    const picture = { name: 'word/media/image1.png', contentType: 'image/png', data: png(64, 64, () => [200, 40, 40, 255]), relationship: { id: 'rIdPicture', type: RELATIONSHIP.image } };
    const opened = await openDocument(page, writeFile('repeated.docx', wordPackage({ body, media: [picture] })));
    site = opened.site;
    await expect(page.locator('#content img')).toHaveCount(200);
    const urls = () =>
      page.evaluate(() => {
        const { created, revoked } = window as unknown as { created: string[]; revoked: string[] };
        return { created: [...created], revoked: [...revoked] };
      });
    const first = await urls();
    expect(first.created).toHaveLength(1);
    expect(await page.locator('#content img').evaluateAll((images) => new Set(images.map((image) => image.getAttribute('src'))).size)).toBe(1);
    await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
    await expect.poll(async () => (await urls()).created.length).toBe(2);
    expect((await urls()).revoked).toEqual(first.created);
    expect(opened.problems).toEqual([]);
  });

  /** Opens `data` as `name` and expects the error banner to say `message`, with no pages shown. */
  async function expectRefused(page: Page, name: string, data: Uint8Array, message: string | RegExp): Promise<void> {
    site = await servePreview(writeFile(name, data));
    const problems = watchProblems(page);
    await page.goto(site.url);
    await expect(page.locator('#error')).toHaveText(message);
    await expect(page.locator('#content .docx-wrapper')).toHaveCount(0);
    expect(problems).toEqual([]);
  }

  /** XML (a run of empty paragraphs) just over MAX_XML_PART_SIZE. */
  const oversizedXml = (root: string) => `<w:${root} ${WORD_NAMESPACES}>${'<w:p/>'.repeat(Math.ceil(MAX_XML_PART_SIZE / 6))}</w:${root}>`;

  test('a damaged document shows the error banner and keeps nothing half-rendered', async ({ page }) => {
    await expectRefused(
      page,
      'damaged.docx',
      wordPackage({ body: `${paragraph(run('unclosed'))}<w:p>` }),
      'Could not display damaged.docx: the document is damaged (word/document.xml is not well-formed XML)',
    );
  });

  test('a part larger than MAX_XML_PART_SIZE is refused before it is parsed', async ({ page }) => {
    await expectRefused(
      page,
      'oversized.docx',
      wordPackage({ body: '<w:p/>'.repeat(Math.ceil(MAX_XML_PART_SIZE / 6)) }),
      'Could not display oversized.docx: its part word/document.xml is larger than 64 MB, the most Document Lens reads from one part',
    );
  });

  test('an XML part reached through a relationship is size-checked whatever its name', async ({ page }) => {
    const header = { name: 'word/media/header.png', contentType: CONTENT_TYPE.header, xml: oversizedXml('hdr'), relationship: { id: 'rIdHeader', type: RELATIONSHIP.header } };
    await expectRefused(
      page,
      'disguised.docx',
      wordPackage({ body: paragraph(run('text')), parts: [header] }),
      'Could not display disguised.docx: its part word/media/header.png is larger than 64 MB, the most Document Lens reads from one part',
    );
  });

  test('a picture larger than MAX_XML_PART_SIZE is not refused', async ({ page }) => {
    const picture = { name: 'word/media/large.png', contentType: 'image/png', data: Buffer.alloc(MAX_XML_PART_SIZE + 1), relationship: { id: 'rIdLarge', type: RELATIONSHIP.image } };
    const opened = await openDocument(page, writeFile('large-picture.docx', wordPackage({ body: paragraph(run('text')), media: [picture] })));
    site = opened.site;
    expect(opened.problems).toEqual([]);
  });

  test('an entry that inflates past its ZIP header is refused', async ({ page }) => {
    const forged = { name: 'word/media/forged.png', data: Buffer.alloc(100_000), declaredSize: 100 };
    await expectRefused(
      page,
      'forged.docx',
      wordPackage({ body: paragraph(run('text')), extraEntries: [forged] }),
      'Could not display forged.docx: the file is damaged: word/media/forged.png is larger than its ZIP header says',
    );
  });

  test('XML beyond the element budget is refused before docx-preview parses it', async ({ page }) => {
    // About 5 million elements in 30 MB: within the size limits, beyond DOCX_LIMITS.maxXmlElements.
    const xml = `<w:hdr ${WORD_NAMESPACES}>${'<w:p/>'.repeat(5_000_000)}</w:hdr>`;
    const header = { name: 'word/header1.xml', contentType: CONTENT_TYPE.header, xml, relationship: { id: 'rIdHeader', type: RELATIONSHIP.header } };
    await expectRefused(
      page,
      'dense.docx',
      wordPackage({ body: paragraph(run('x')), parts: [header] }),
      'Could not display dense.docx: it has more than 4,000,000 XML elements, the most Document Lens reads',
    );
  });

  test('archives JSZip would read differently from the checks are refused', async ({ page }) => {

    const refusal = (detail: string) => `Could not display tampered.docx: the file is damaged or uses ZIP features Document Lens does not read (${detail})`;
    // A well-formed document part, and a malformed one JSZip would put under the same name.
    const decoy = wordPackage({ body: paragraph(run('text')), extraEntries: [{ name: 'x/../word/document.xml', data: '<w:document><not-closed></w:document>' }] });
    await expectRefused(page, 'tampered.docx', decoy, refusal('unusual entry name "x/../word/document.xml"'));
    // A harmless archive in front of the real one, with the same entries: JSZip skips to the real one (whose picture
    // inflates far past its declared size), a reader that trusts the end record's offset stays on the decoy.
    const harmless = wordPackage({ body: paragraph(run('text')), extraEntries: [{ name: 'word/media/forged.png', data: 'x' }] });
    const real = wordPackage({
      body: paragraph(run(Array.from({ length: 5000 }, (_, i) => i).join(','))),
      extraEntries: [{ name: 'word/media/forged.png', data: Buffer.alloc(1_000_000), declaredSize: 100 }],
    });
    await expectRefused(page, 'tampered.docx', prependedDecoy(harmless, real), refusal('unexpected data around the ZIP central directory'));
    // Chrome's inflater stops at bytes after the end of a deflate stream (pako, in JSZip, ignores them).
    const junk = markDeflated(wordPackage({ body: paragraph(run('text')), extraEntries: [junkDeflateEntry('word/media/image1.png', 8 * 1024 * 1024)] }), 'word/media/image1.png');
    await expectRefused(page, 'tampered.docx', junk, refusal('word/media/image1.png cannot be inflated'));
  });

  test('a password-protected document is refused with a clear message', async ({ page }) => {
    site = await servePreview(writeFile('locked.docx', Buffer.concat([compoundFileHeaderStub(), Buffer.from('EncryptedPackage', 'utf16le')])));
    const problems = watchProblems(page);
    await page.goto(site.url);
    await expect(page.locator('#error')).toHaveText('This .docx file is password-protected; Document Lens cannot open encrypted files');
    expect(problems).toEqual([]);
  });

  const ownFiles = (process.env['E2E_FILES'] ?? '').split(path.delimiter).filter((file) => file.toLowerCase().endsWith('.docx'));
  for (const file of ownFiles) {
    test(`own file ${path.basename(file)}: renders its pages`, async ({ page }) => {
      test.setTimeout(120_000);
      const opened = await openDocument(page, file, 90_000);
      site = opened.site;
      expect(opened.problems).toEqual([]);
    });
  }
});
