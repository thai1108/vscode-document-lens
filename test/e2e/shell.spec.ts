// The shared viewer frame in Chrome, through the browser preview whose mock host replays VS Code's message
// serializer, for the synthetic sample of every format. Set E2E_FILES (paths joined by the OS path delimiter) to run
// the same checks on your own documents; they are only read locally.
import { expect, test } from '@playwright/test';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import { FORMATS } from '../../src/formats/index.ts';
import { expectShellLoaded, servePreview, watchProblems, writeSample } from './preview-server.ts';

const SAMPLE_NAMES = FORMATS.map((format) => `sample${format.extensions[0]}`);

for (const fileName of SAMPLE_NAMES) {
  test(`${fileName}: loads in the shared frame without console or CSP errors`, async ({ page }) => {
    const site = await servePreview(await writeSample(fileName));
    try {
      const problems = watchProblems(page);
      await page.goto(site.url);
      await expectShellLoaded(page, fileName);
      await expect(page.locator('#content')).not.toBeEmpty();
      expect(problems).toEqual([]);
    } finally {
      await site.close();
    }
  });
}

test('zoom controls, keys and state survive a reload of the file', async ({ page }) => {
  const site = await servePreview(await writeSample('sample.docx'));
  try {
    const problems = watchProblems(page);
    await page.goto(site.url);
    await expectShellLoaded(page, 'sample.docx');
    await page.locator('#zoom-in').click();
    await expect(page.locator('#zoom-value')).toHaveText('110%');
    await page.locator('#content').focus();
    await page.keyboard.press('-');
    await page.keyboard.press('-');
    await expect(page.locator('#zoom-value')).toHaveText('90%');
    await page.evaluate(() => (window as unknown as { previewReload(): void }).previewReload());
    await expect(page.locator('#zoom-value')).toHaveText('90%');
    await page.reload();
    await expect(page.locator('#zoom-value')).toHaveText('90%');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

test('a file that fails the host check shows the error banner', async ({ page }) => {
  const file = path.join(mkdtempSync(path.join(tmpdir(), 'document-lens-broken-')), 'broken.docx');
  writeFileSync(file, 'not a zip file');
  const site = await servePreview(file);
  try {
    const problems = watchProblems(page);
    await page.goto(site.url);
    await expect(page.locator('#file-name')).toHaveText('broken.docx');
    await expect(page.locator('#error')).toHaveText('Not a valid .docx file: Not a ZIP archive (end of central directory not found)');
    expect(problems).toEqual([]);
  } finally {
    await site.close();
  }
});

const ownFiles = (process.env['E2E_FILES'] ?? '').split(path.delimiter).filter(Boolean);
for (const file of ownFiles) {
  test(`own file ${path.basename(file)}: loads without errors`, async ({ page }) => {
    const site = await servePreview(file);
    try {
      const problems = watchProblems(page);
      await page.goto(site.url);
      await expectShellLoaded(page, path.basename(file));
      expect(problems).toEqual([]);
    } finally {
      await site.close();
    }
  });
}
