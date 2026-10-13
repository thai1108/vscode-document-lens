// Bundle options for the spreadsheet viewer (merged over the defaults in esbuild.mjs).
import { readFileSync, writeFileSync } from 'node:fs';
import * as path from 'node:path';
import { URL } from 'node:url';

// The bundle includes SheetJS Community Edition, whose Apache-2.0 license must travel with it.
const thirdPartyLicenses = {
  name: 'third-party-licenses',
  setup(build) {
    build.onEnd((result) => {
      if (result.errors.length || !build.initialOptions.outfile) {
        return;
      }
      const version = JSON.parse(readFileSync(new URL('../../../node_modules/xlsx/package.json', import.meta.url), 'utf8')).version;
      const license = readFileSync(new URL('../../../node_modules/xlsx/LICENSE', import.meta.url), 'utf8');
      const target = build.initialOptions.outfile.replace(/\.js$/, '.LICENSES.txt');
      writeFileSync(target, `${path.basename(build.initialOptions.outfile)} includes SheetJS Community Edition ${version} (https://sheetjs.com), licensed as follows.\n\n${license}`);
    });
  },
};

export default {
  plugins: [thirdPartyLicenses],
};
