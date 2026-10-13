// Bundle options for the .docx viewer, merged over the defaults in esbuild.mjs.
//
// The bundle includes docx-preview (Apache-2.0) and JSZip's browser build (used under its MIT option), which itself
// contains pako, lie, immediate and setImmediate. Their license texts are written next to the bundle.
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import * as path from 'node:path';

const require = createRequire(import.meta.url);

/** The folder of `name`, resolved from `from` (a file inside the package that depends on it). */
function packageFolder(name, from = import.meta.url) {
  return path.dirname(createRequire(from).resolve(`${name}/package.json`));
}

function licenses() {
  // docx-preview's package exports do not include its package.json.
  const docxPreview = path.resolve(path.dirname(require.resolve('docx-preview')), '..');
  const jszip = packageFolder('jszip');
  const lie = packageFolder('lie', path.join(jszip, 'package.json'));
  const pako = packageFolder('pako', path.join(jszip, 'package.json'));
  const packages = [
    [docxPreview, 'LICENSE'],
    [jszip, 'LICENSE.markdown'],
    [pako, 'LICENSE'],
    [lie, 'license.md'],
    [packageFolder('immediate', path.join(lie, 'package.json')), 'LICENSE.txt'],
    [packageFolder('setimmediate', path.join(jszip, 'package.json')), 'LICENSE.txt'],
  ];
  const sections = packages.map(([folder, file]) => {
    const { name, version, license } = JSON.parse(readFileSync(path.join(folder, 'package.json'), 'utf8'));
    return `${name} ${version} (${license.replace(/^\((.*)\)$/, '$1')})\n\n${readFileSync(path.join(folder, file), 'utf8').trim()}\n`;
  });
  // pako's zlib port carries the zlib license in its source files rather than in LICENSE.
  const zlib = readFileSync(path.join(pako, 'lib/zlib/inflate.js'), 'utf8').match(/^\/\/ \(C\) 1995[\s\S]*?source distribution\.$/m)?.[0];
  if (!zlib) {
    throw new Error('pako zlib license notice not found');
  }
  sections.push(`zlib (ported in pako)\n\n${zlib.replace(/^\/\/ ?/gm, '')}\n`);
  return `Third-party software included in webview-docx.js\n\n${sections.join(`\n${'='.repeat(80)}\n\n`)}`;
}

const thirdPartyLicenses = {
  name: 'third-party-licenses',
  setup(build) {
    build.onEnd((result) => {
      if (result.errors.length === 0 && build.initialOptions.outfile) {
        writeFileSync(build.initialOptions.outfile.replace(/\.js$/, '.licenses.txt'), licenses());
      }
    });
  },
};

export default {
  plugins: [thirdPartyLicenses],
  banner: { js: '/*! Includes docx-preview (Apache-2.0) and JSZip with pako, lie, immediate and setImmediate (MIT); see webview-docx.licenses.txt */' },
};
