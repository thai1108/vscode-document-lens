// Synthetic sample files, one module per format under test/fixtures/ (each exports `samples`). They are written by
// `pnpm fixtures <dir>` and used by the unit, browser and VS Code suites. Never put real documents here.
import { readdirSync } from 'node:fs';
import * as path from 'node:path';
import { pathToFileURL } from 'node:url';

export interface Sample {
  /** Name the sample is written under; its extension picks the format. */
  fileName: string;
  build(): Uint8Array;
}

const FIXTURES = path.resolve(import.meta.dirname, '../fixtures');

export async function loadSamples(): Promise<Sample[]> {
  const modules = readdirSync(FIXTURES)
    .filter((name) => name.endsWith('.ts'))
    .sort();
  const samples: Sample[] = [];
  for (const name of modules) {
    const module = (await import(pathToFileURL(path.join(FIXTURES, name)).href)) as { samples?: Sample[] };
    samples.push(...(module.samples ?? []));
  }
  return samples;
}
