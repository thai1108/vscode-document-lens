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
