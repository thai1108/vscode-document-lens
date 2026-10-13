// Bundle options for the PDF viewer, merged over the defaults in esbuild.mjs.
//
// pdf.js needs more than its main module at runtime: a worker, the CMaps for CJK text, standard font data for Symbol
// and ZapfDingbats, and WebAssembly decoders for JBIG2/CCITT fax and JPEG 2000 images. The webview may not fetch
// anything (no connect-src) and cannot start a worker from an extension file, so all of it is embedded in the bundle:
//   pdfjs-assets:worker  the worker module's source, started from a blob: URL
//   pdfjs-assets:files   base64 of every file pdf.js may ask for, keyed by its option name and file name
// The licenses of what is embedded are written next to the bundle.
import { readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import * as path from 'node:path';

const PDFJS = path.dirname(createRequire(import.meta.url).resolve('pdfjs-dist/package.json'));
const pdfjsFile = (name) => path.join(PDFJS, name);

const WORKER = pdfjsFile('build/pdf.worker.min.mjs');
// With useSystemFonts (the default in a browser) pdf.js only asks for Symbol and ZapfDingbats; other standard fonts
// fall back to the system's.
const STANDARD_FONTS = ['FoxitSymbol.pfb', 'FoxitDingbats.pfb'];
// qcms (ICC profiles) is unused: pdf.js only loads it when the worker fetches its own files.
const WASM = ['jbig2.wasm', 'openjpeg.wasm'];
const LICENSES = [
  ['pdf.js', 'LICENSE'],
  ['CMaps (Adobe)', 'cmaps/LICENSE'],
  ['Standard font data (Foxit)', 'standard_fonts/LICENSE_FOXIT'],
  ['JBIG2 decoder (pdf.js)', 'wasm/LICENSE_PDFJS_JBIG2'],
  ['JBIG2 decoder (pdfium)', 'wasm/LICENSE_JBIG2'],
  ['OpenJPEG decoder (pdf.js)', 'wasm/LICENSE_PDFJS_OPENJPEG'],
  ['OpenJPEG', 'wasm/LICENSE_OPENJPEG'],
];

function base64Files(folder, names) {
  return Object.fromEntries(names.map((name) => [name, readFileSync(path.join(PDFJS, folder, name)).toString('base64')]));
}

function embeddedFiles() {
  const cMaps = readdirSync(pdfjsFile('cmaps')).filter((name) => name.endsWith('.bcmap'));
  return {
    cMapUrl: base64Files('cmaps', cMaps),
    standardFontDataUrl: base64Files('standard_fonts', STANDARD_FONTS),
    wasmUrl: base64Files('wasm', WASM),
  };
}

function licenses() {
  const { version } = JSON.parse(readFileSync(pdfjsFile('package.json'), 'utf8'));
  const sections = LICENSES.map(([title, file]) => `${title} (pdfjs-dist ${version}/${file})\n\n${readFileSync(pdfjsFile(file), 'utf8').trim()}\n`);
  return `Third-party software embedded in webview-pdf.js\n\n${sections.join(`\n${'='.repeat(80)}\n\n`)}`;
}

const pdfjsAssets = {
  name: 'pdfjs-assets',
  setup(build) {
    build.onResolve({ filter: /^pdfjs-assets:/ }, (args) => ({ path: args.path, namespace: 'pdfjs-assets' }));
    build.onLoad({ filter: /^pdfjs-assets:worker$/, namespace: 'pdfjs-assets' }, () => ({
      contents: readFileSync(WORKER, 'utf8'),
      loader: 'text',
      watchFiles: [WORKER],
    }));
    build.onLoad({ filter: /^pdfjs-assets:files$/, namespace: 'pdfjs-assets' }, () => ({
      contents: JSON.stringify(embeddedFiles()),
      loader: 'json',
    }));
    build.onEnd((result) => {
      if (result.errors.length === 0) {
        writeFileSync(build.initialOptions.outfile.replace(/\.js$/, '.licenses.txt'), licenses());
      }
    });
  },
};

export default {
  plugins: [pdfjsAssets],
  // pdf.js reads import.meta.url only on its Node.js code path, which never runs in the webview.
  define: { 'import.meta.url': 'undefined' },
  banner: { js: '/*! Includes pdf.js (Apache-2.0) with Adobe CMaps, Foxit fonts, JBIG2 and OpenJPEG decoders; see webview-pdf.licenses.txt */' },
};
