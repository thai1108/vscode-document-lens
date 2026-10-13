import * as esbuild from 'esbuild';
import { existsSync, readdirSync } from 'node:fs';

const watch = process.argv.includes('--watch');

const common = {
  bundle: true,
  sourcemap: watch,
  minify: !watch,
  logLevel: 'info',
};

// Every folder under src/webview with a main.ts (except the shared shell) is one viewer bundle:
// src/webview/<name>/main.ts → dist/webview-<name>.js, plus dist/webview-<name>.css for the CSS it imports.
const viewers = readdirSync('src/webview', { withFileTypes: true })
  .filter((entry) => entry.isDirectory() && entry.name !== 'shell' && existsSync(`src/webview/${entry.name}/main.ts`))
  .map((entry) => entry.name);

const contexts = await Promise.all([
  esbuild.context({
    ...common,
    entryPoints: ['src/extension.ts'],
    outfile: 'dist/extension.cjs',
    platform: 'node',
    format: 'cjs',
    target: 'node20',
    external: ['vscode'],
  }),
  ...viewers.map((name) =>
    esbuild.context({
      ...common,
      entryPoints: [`src/webview/${name}/main.ts`],
      outfile: `dist/webview-${name}.js`,
      platform: 'browser',
      format: 'iife',
      target: 'chrome120',
    }),
  ),
]);

if (watch) {
  await Promise.all(contexts.map((ctx) => ctx.watch()));
} else {
  await Promise.all(contexts.map((ctx) => ctx.rebuild()));
  await Promise.all(contexts.map((ctx) => ctx.dispose()));
}
