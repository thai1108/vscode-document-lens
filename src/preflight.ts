// Cheap structural checks run on the host before a file's bytes are handed to a viewer, so an obviously wrong or
// hostile file gets a clear message instead of a library failing (or unpacking a ZIP bomb) inside the webview.
import { sniffEncoding } from './csv/encoding.ts';
import { OOXML_ZIP_LIMITS } from './limits.ts';
import { ZipArchive } from './zip.ts';

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const PDF_SIGNATURE = '%PDF-';
/** Readers accept a PDF header preceded by up to 1 KB of junk (some generators prepend one). */
const PDF_HEADER_WINDOW = 1024;
/** Stream name in the compound file that wraps a password-protected OOXML package, as UTF-16LE. */
const ENCRYPTED_PACKAGE = Buffer.from('EncryptedPackage', 'utf16le');

const OOXML_FOLDERS: Record<string, string> = { 'word/': '.docx', 'xl/': '.xlsx', 'ppt/': '.pptx' };

/** OLE compound file (.doc, .xls, .ppt, and encrypted OOXML packages). */
export function isCompoundFile(bytes: Uint8Array): boolean {
  return CFB_SIGNATURE.every((byte, i) => bytes[i] === byte);
}

export function isZip(bytes: Uint8Array): boolean {
  return bytes[0] === 0x50 && bytes[1] === 0x4b && (bytes[2] === 0x03 || bytes[2] === 0x05) && (bytes[3] === 0x04 || bytes[3] === 0x06);
}

export function checkPdf(bytes: Uint8Array): void {
  const head = Buffer.from(bytes.buffer, bytes.byteOffset, Math.min(bytes.byteLength, PDF_HEADER_WINDOW + PDF_SIGNATURE.length));
  if (!head.includes(PDF_SIGNATURE, 0, 'latin1')) {
    throw new Error('Not a .pdf file (no %PDF- header)');
  }
}

/** .csv and .tsv files are text; a workbook or other binary file saved under their extension is named for what it is. */
export function checkText(bytes: Uint8Array, extension: string): void {
  if (isZip(bytes)) {
    throw new Error(`This ${extension} file is a ZIP package, probably an .xlsx workbook with the wrong extension`);
  }
  if (isCompoundFile(bytes)) {
    throw new Error(`This ${extension} file is an Office binary file, probably an .xls workbook with the wrong extension`);
  }
  if (Buffer.from(bytes.buffer, bytes.byteOffset, Math.min(bytes.byteLength, PDF_SIGNATURE.length)).toString('latin1') === PDF_SIGNATURE) {
    throw new Error(`This ${extension} file is a PDF document`);
  }
  if (sniffEncoding(bytes) === 'binary') {
    throw new Error(`This ${extension} file contains binary data, not text`);
  }
}

/** .doc, .xls and .ppt are compound files; a ZIP here is usually the newer format under the old extension. */
export function checkCompoundFile(bytes: Uint8Array, extension: string): void {
  if (isCompoundFile(bytes)) {
    return;
  }
  if (isZip(bytes)) {
    throw new Error(`This ${extension} file is a ZIP package, probably a ${extension}x file with the wrong extension`);
  }
  throw new Error(`Not a ${extension} file (no compound file header)`);
}

/**
 * .docx, .xlsx and .pptx are ZIP packages: the central directory must stay within OOXML_ZIP_LIMITS, and the package
 * relationship part must point at a main document part that exists in the archive.
 */
export function checkOoxmlPackage(bytes: Uint8Array, extension: string): void {
  const buffer = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (isCompoundFile(bytes)) {
    throw new Error(
      buffer.includes(ENCRYPTED_PACKAGE)
        ? `This ${extension} file is password-protected; Document Lens cannot open encrypted files`
        : `This ${extension} file is an older binary file, probably a ${extension.slice(0, -1)} file with the wrong extension`,
    );
  }
  let zip: ZipArchive;
  try {
    zip = ZipArchive.open(buffer, OOXML_ZIP_LIMITS);
  } catch (error) {
    throw new Error(`Not a valid ${extension} file: ${(error as Error).message}`, { cause: error });
  }
  if (!zip.has('[Content_Types].xml') || !zip.has('_rels/.rels')) {
    throw new Error(`Not a ${extension} file (package parts are missing)`);
  }
  const main = officeDocumentTarget(zip.readText('_rels/.rels'));
  if (!main || !zip.has(main)) {
    throw new Error(`Not a ${extension} file (the main document part is missing)`);
  }
  const folder = Object.keys(OOXML_FOLDERS).find((prefix) => main.startsWith(prefix));
  const actual = folder ? OOXML_FOLDERS[folder] : undefined;
  if (actual && actual !== extension) {
    throw new Error(`This ${extension} file is actually a ${actual} file; rename it to open it`);
  }
}

/** The package's main part (relationship type …/officeDocument) as a ZIP entry name. */
export function officeDocumentTarget(rels: string): string | undefined {
  for (const [relationship] of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const attributes = Object.fromEntries([...relationship.matchAll(/\b(Type|Target)\s*=\s*(["'])(.*?)\2/g)].map(([, name, , value]) => [name, value]));
    if (attributes['Type']?.endsWith('/officeDocument') && attributes['Target']) {
      return attributes['Target'].replaceAll('&amp;', '&').replace(/^\//, '');
    }
  }
  return undefined;
}
