// Writes the files the VS Code suite opens: every synthetic sample, an oversized and a broken file.
import { mkdirSync, truncateSync, writeFileSync } from 'node:fs';
import * as path from 'node:path';
import { MAX_FILE_SIZE } from '../../src/limits.ts';
import { loadSamples } from '../support/samples.ts';

const out = path.resolve(import.meta.dirname, '../../dist-test/vscode/workspace');
mkdirSync(out, { recursive: true });
for (const sample of await loadSamples()) {
  writeFileSync(path.join(out, sample.fileName), sample.build());
}
writeFileSync(path.join(out, 'broken.docx'), 'not a zip file');
// Sparse, so the oversized file costs no disk space.
writeFileSync(path.join(out, 'oversized.pdf'), '');
truncateSync(path.join(out, 'oversized.pdf'), MAX_FILE_SIZE.pdf + 1);
console.log(`Wrote VS Code test fixtures to ${out}`);
