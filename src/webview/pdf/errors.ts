import { MAX_PDF_PAGES } from '../../limits.ts';

/** Message for a document pdf.js could not open; the shell shows it as "Could not display <file>: <message>". */
export function loadErrorMessage(error: unknown): string {
  const { name, message } = (typeof error === 'object' && error !== null ? error : {}) as { name?: unknown; message?: unknown };
  const detail = typeof message === 'string' && message ? message : String(error);
  switch (name) {
    case 'PasswordException':
      return 'This .pdf file is password-protected; Document Lens cannot open encrypted files';
    case 'InvalidPDFException':
      return `Not a valid .pdf file: ${detail}`;
    default:
      return detail;
  }
}

/**
 * Why a document with `pages` pages (pdf.js takes the count from the file's page tree) is not shown, if it is not.
 * The viewer lays out a placeholder per page up front, so a claimed count of millions would exhaust memory.
 */
export function pageCountError(pages: number): string | undefined {
  if (!(pages >= 1)) {
    return 'This .pdf file has no pages';
  }
  return pages > MAX_PDF_PAGES ? `This .pdf file has too many pages (${pages}; limit ${MAX_PDF_PAGES})` : undefined;
}
