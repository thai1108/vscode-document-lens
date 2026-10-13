// The host's preflight checks for .xls/.xlsx and the viewer state read back from the webview.
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { xls } from '../../../src/formats/xls.ts';
import { xlsx } from '../../../src/formats/xlsx.ts';
import { readSavedState } from '../../../src/webview/spreadsheet/viewer.ts';
import { sampleXls } from '../../fixtures/xls.ts';
import { sampleXlsx } from '../../fixtures/xlsx.ts';

describe('spreadsheet formats on the host', () => {
  it('accepts the samples', () => {
    xls.check(sampleXls());
    xlsx.check(sampleXlsx());
  });

  it('explains web pages and XML spreadsheets saved as .xls', () => {
    const message = /This \.xls file is a web page or XML spreadsheet with an \.xls extension/;
    assert.throws(() => xls.check(Buffer.from('<html><body><table><tr><td>1</td></tr></table></body></html>')), message);
    assert.throws(() => xls.check(Buffer.from('﻿\r\n  <?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"/>')), message);
  });

  it('keeps the other messages for files that are not .xls', () => {
    assert.throws(() => xls.check(sampleXlsx()), /This \.xls file is a ZIP package, probably a \.xlsx file/);
    assert.throws(() => xls.check(Buffer.from('a\tb\n1\t2\n')), /Not a \.xls file \(no compound file header\)/);
    assert.throws(() => xlsx.check(sampleXls()), /older binary file, probably a \.xls file/);
  });
});

describe('saved viewer state', () => {
  it('keeps what is well-formed', () => {
    assert.deepEqual(readSavedState({ sheet: 'Data', positions: { Data: { top: 120, left: 0, row: 4, col: 2 }, Other: { top: 5, left: 6 } } }), {
      sheet: 'Data',
      positions: { Data: { top: 120, left: 0, row: 4, col: 2 }, Other: { top: 5, left: 6 } },
    });
    assert.deepEqual(readSavedState({ headerRow: false }), { headerRow: false });
  });

  it('drops anything else', () => {
    assert.deepEqual(readSavedState(undefined), {});
    assert.deepEqual(readSavedState('Data'), {});
    assert.deepEqual(readSavedState({ headerRow: 'no' }), {});
    assert.deepEqual(readSavedState({ sheet: 3, positions: { A: null, B: { top: 'x', left: 1 }, C: { top: -5, left: Infinity }, D: { top: -5, left: 2, row: 1.5, col: 1 } } }), {
      positions: { D: { top: 0, left: 2 } },
    });
  });
});
