/** Why a .ppt file cannot be shown; the message is what the user reads in the error banner. */
export type PptErrorKind = 'encrypted' | 'too-old' | 'damaged' | 'not-ppt' | 'too-complex';

export class PptError extends Error {
  readonly kind: PptErrorKind;

  constructor(kind: PptErrorKind, message: string) {
    super(message);
    this.name = 'PptError';
    this.kind = kind;
  }
}

export function damaged(detail: string): PptError {
  return new PptError('damaged', `This .ppt file is damaged: ${detail}`);
}

export function encrypted(): PptError {
  return new PptError('encrypted', 'This .ppt file is password-protected; Document Lens cannot open encrypted files');
}

export function tooOld(): PptError {
  return new PptError('too-old', 'This .ppt file was saved by PowerPoint 95 or older, a format too old for Document Lens to read');
}
