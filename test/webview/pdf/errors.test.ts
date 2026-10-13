import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { loadErrorMessage } from '../../../src/webview/pdf/errors.ts';

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
