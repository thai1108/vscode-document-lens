// Keeps the format registry, package.json, the viewer bundles and the samples in step: adding a format means
// touching each of them, and this suite says which one is missing.
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import * as path from 'node:path';
import { before, describe, it } from 'node:test';
import { FORMATS, formatForFile } from '../src/formats/index.ts';
import { viewTypeOf } from '../src/formats/types.ts';
import { loadSamples, type Sample } from './support/samples.ts';

const ROOT = path.resolve(import.meta.dirname, '..');
const manifest = JSON.parse(readFileSync(path.join(ROOT, 'package.json'), 'utf8')) as {
  contributes: { customEditors: { viewType: string; selector: { filenamePattern: string }[]; priority: string }[] };
};

describe('format registry', () => {
  let samples: Sample[];
  before(async () => {
    samples = await loadSamples();
  });

  it('declares exactly one default custom editor per format, with the same extensions', () => {
    const editors = manifest.contributes.customEditors;
    assert.deepEqual(editors.map((editor) => editor.viewType).sort(), FORMATS.map(viewTypeOf).sort());
    for (const format of FORMATS) {
      const editor = editors.find((candidate) => candidate.viewType === viewTypeOf(format));
      assert.deepEqual(editor?.selector.map((selector) => selector.filenamePattern), format.extensions.map((extension) => `*${extension}`));
      assert.equal(editor?.priority, 'default');
    }
  });

  it('has unique ids and extensions, and a viewer bundle for every format', () => {
    assert.equal(new Set(FORMATS.map((format) => format.id)).size, FORMATS.length);
    const extensions = FORMATS.flatMap((format) => format.extensions);
    assert.equal(new Set(extensions).size, extensions.length);
    for (const format of FORMATS) {
      assert.ok(existsSync(path.join(ROOT, 'src/webview', format.webview, 'main.ts')), `src/webview/${format.webview}/main.ts`);
      assert.ok(format.maxFileSize > 0);
    }
  });

  it('picks formats by extension, ignoring case', () => {
    assert.equal(formatForFile('Report.PDF')?.id, 'pdf');
    assert.equal(formatForFile('a.docx')?.id, 'docx');
    assert.equal(formatForFile('a.doc')?.id, 'doc');
    assert.equal(formatForFile('a.pptx')?.id, 'pptx');
    assert.equal(formatForFile('Export.CSV')?.id, 'csv');
    assert.equal(formatForFile('a.tsv')?.id, 'tsv');
    assert.equal(formatForFile('a.txt'), undefined);
  });

  it('has a sample for every format that passes its own check and fails every other format’s', () => {
    for (const format of FORMATS) {
      const own = samples.filter((sample) => formatForFile(sample.fileName) === format);
      assert.ok(own.length, `a sample for ${format.id} in test/fixtures/`);
      for (const sample of own) {
        format.check(sample.build());
      }
    }
    for (const sample of samples) {
      const bytes = sample.build();
      // The compound file stubs are indistinguishable by signature alone, and so is a .csv from a .tsv.
      const families = [['doc', 'xls', 'ppt'], ['csv', 'tsv']];
      const sameFamily = (a: string, b: string) => families.some((family) => family.includes(a) && family.includes(b));
      for (const format of FORMATS.filter((candidate) => candidate !== formatForFile(sample.fileName))) {
        if (sameFamily(format.id, formatForFile(sample.fileName)?.id ?? '')) {
          continue;
        }
        assert.throws(() => format.check(bytes), `${sample.fileName} is not a ${format.id}`);
      }
    }
  });
});
