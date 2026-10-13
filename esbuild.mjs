import * as esbuild from 'esbuild';
import { existsSync, readdirSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

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

// A viewer tunes its own bundle (loaders, plugins, define, ...) in src/webview/<name>/build.mjs, whose default export
// is merged over the defaults below, so viewers never need to edit this shared file.
async function viewerOptions(name) {
  const file = `src/webview/${name}/build.mjs`;
  return existsSync(file) ? ((await import(pathToFileURL(file).href)).default ?? {}) : {};
}

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
  ...(await Promise.all(viewers.map(viewerOptions))).map((options, i) =>
    esbuild.context({
      ...common,
      platform: 'browser',
      format: 'iife',
      target: 'chrome120',
      ...options,
      entryPoints: [`src/webview/${viewers[i]}/main.ts`],
      outfile: `dist/webview-${viewers[i]}.js`,
    }),
  ),
]);

if (watch) {
  await Promise.all(contexts.map((ctx) => ctx.watch()));
} else {
  await Promise.all(contexts.map((ctx) => ctx.rebuild()));
  await Promise.all(contexts.map((ctx) => ctx.dispose()));
}
