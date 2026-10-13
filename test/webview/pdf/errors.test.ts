import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { MAX_PDF_PAGES } from '../../../src/limits.ts';
import { loadErrorMessage, pageCountError } from '../../../src/webview/pdf/errors.ts';

describe('pageCountError', () => {
  it('accepts documents up to the page limit', () => {
    assert.equal(pageCountError(1), undefined);
    assert.equal(pageCountError(MAX_PDF_PAGES), undefined);
  });

  it('refuses empty documents and page trees that claim more pages than the limit', () => {
    assert.equal(pageCountError(0), 'This .pdf file has no pages');
    assert.equal(pageCountError(Number.NaN), 'This .pdf file has no pages');
    assert.equal(pageCountError(MAX_PDF_PAGES + 1), `This .pdf file has too many pages (${MAX_PDF_PAGES + 1}; limit ${MAX_PDF_PAGES})`);
    assert.equal(pageCountError(100_000_000), `This .pdf file has too many pages (100000000; limit ${MAX_PDF_PAGES})`);
  });
});

/** pdf.js rejects with exceptions identified by `name` (rebuilt on the main thread from the worker's). */
function exception(name: string, message: string): Error {
  return Object.assign(new Error(message), { name });
}

describe('loadErrorMessage', () => {
  it('explains password-protected files', () => {
    assert.equal(
      loadErrorMessage(exception('PasswordException', 'No password given')),
      'This .pdf file is password-protected; Document Lens cannot open encrypted files',
    );
  });

  it('marks files without a PDF structure as invalid', () => {
    assert.equal(loadErrorMessage(exception('InvalidPDFException', 'Invalid PDF structure.')), 'Not a valid .pdf file: Invalid PDF structure.');
  });

  it('passes other messages through and copes with values that are not errors', () => {
    assert.equal(loadErrorMessage(new Error('This .pdf file has no pages')), 'This .pdf file has no pages');
    assert.equal(loadErrorMessage(exception('UnknownErrorException', '')), 'UnknownErrorException');
    assert.equal(loadErrorMessage('worker crashed'), 'worker crashed');
    assert.equal(loadErrorMessage(undefined), 'undefined');
  });
});
