// Writes the synthetic sample of every format: `pnpm fixtures <out-dir>`.
import { mkdirSync, writeFileSync } from 'node:fs';
import * as path from 'node:path';
import { loadSamples } from '../test/support/samples.ts';

const [outDir] = process.argv.slice(2);
if (!outDir) {
  console.error('usage: pnpm fixtures <out-dir>');
  process.exit(2);
}
mkdirSync(outDir, { recursive: true });
const samples = await loadSamples();
for (const sample of samples) {
  writeFileSync(path.join(outDir, sample.fileName), sample.build());
}
console.log(`Wrote ${samples.map((sample) => sample.fileName).join(', ')} to ${outDir}`);
