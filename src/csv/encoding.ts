// Text encodings of delimited text files (.csv, .tsv). Shared by the extension host and the webview, so it uses no
// Node or DOM APIs beyond TextDecoder, which both have (Shift_JIS included).

export type TextEncodingName = 'utf-8' | 'utf-16le' | 'utf-16be' | 'shift_jis';

export interface TextEncoding {
  readonly name: TextEncodingName;
  /** Length in bytes of the byte order mark the file starts with; 0 without one. */
  readonly bom: number;
}

export interface DecodedText {
  readonly encoding: TextEncoding;
  /** The text without the byte order mark. */
  readonly text: string;
  /** Some bytes were not valid in the encoding and are shown as U+FFFD. */
  readonly lossy: boolean;
}

/** How much of the start of the file is sniffed for UTF-16 without a byte order mark and for binary data. */
const SNIFF_BYTES = 64 * 1024;

/**
 * The encoding a file declares or plainly shows: a byte order mark, or UTF-16 recognized by its zero bytes. Returns
 * undefined for other files (decodeText() then tells UTF-8 from Shift_JIS) and 'binary' for files with zero bytes
 * that are not UTF-16.
 */
export function sniffEncoding(bytes: Uint8Array): TextEncoding | 'binary' | undefined {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return { name: 'utf-8', bom: 3 };
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return { name: 'utf-16le', bom: 2 };
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return { name: 'utf-16be', bom: 2 };
  }
  const length = Math.min(bytes.length, SNIFF_BYTES) & ~1;
  let evenZeros = 0;
  let oddZeros = 0;
  for (let i = 0; i < length; i += 2) {
    if (bytes[i] === 0) {
      evenZeros++;
    }
    if (bytes[i + 1] === 0) {
      oddZeros++;
    }
  }
  if (!evenZeros && !oddZeros) {
    return undefined;
  }
  const pairs = length / 2;
  // Text in UTF-16 has a zero in the high byte of every ASCII character (delimiters, digits, line breaks) and almost
  // never in the low byte.
  if (oddZeros >= pairs * 0.3 && evenZeros <= pairs * 0.02) {
    return { name: 'utf-16le', bom: 0 };
  }
  if (evenZeros >= pairs * 0.3 && oddZeros <= pairs * 0.02) {
    return { name: 'utf-16be', bom: 0 };
  }
  // A stray zero byte does not make a text file binary; pictures, archives and executables are full of them.
  return evenZeros + oddZeros > length / 1000 ? 'binary' : undefined;
}

/**
 * Decodes a text file: a byte order mark or UTF-16 zero bytes decide; otherwise valid UTF-8 is UTF-8, and a file
 * that is not but is valid Shift_JIS (what Excel writes on Japanese Windows) is Shift_JIS. Anything else is read as
 * UTF-8 with the invalid bytes replaced. Throws for binary files.
 */
export function decodeText(bytes: Uint8Array): DecodedText {
  const sniffed = sniffEncoding(bytes);
  if (sniffed === 'binary') {
    throw new Error('The file contains binary data, not text');
  }
  if (sniffed) {
    return decodeAs(bytes, sniffed);
  }
  const utf8 = tryDecode(bytes, 'utf-8');
  if (utf8 !== undefined) {
    return { encoding: { name: 'utf-8', bom: 0 }, text: utf8, lossy: false };
  }
  const shiftJis = tryDecode(bytes, 'shift_jis');
  if (shiftJis !== undefined) {
    return { encoding: { name: 'shift_jis', bom: 0 }, text: shiftJis, lossy: false };
  }
  return decodeAs(bytes, { name: 'utf-8', bom: 0 });
}

/** Decodes bytes in a known encoding, after its byte order mark. */
export function decodeAs(bytes: Uint8Array, encoding: TextEncoding): DecodedText {
  const body = bytes.subarray(encoding.bom);
  const strict = tryDecode(body, encoding.name);
  if (strict !== undefined) {
    return { encoding, text: strict, lossy: false };
  }
  return { encoding, text: new TextDecoder(encoding.name, { ignoreBOM: true }).decode(body), lossy: true };
}

/** The text if every byte is valid in the encoding, else undefined. */
function tryDecode(bytes: Uint8Array, name: TextEncodingName): string | undefined {
  try {
    return new TextDecoder(name, { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch {
    return undefined;
  }
}

export function encodingLabel(encoding: TextEncoding): string {
  switch (encoding.name) {
    case 'utf-8':
      return encoding.bom ? 'UTF-8 with BOM' : 'UTF-8';
    case 'utf-16le':
      return encoding.bom ? 'UTF-16 LE' : 'UTF-16 LE (no BOM)';
    case 'utf-16be':
      return encoding.bom ? 'UTF-16 BE' : 'UTF-16 BE (no BOM)';
    case 'shift_jis':
      return 'Shift_JIS';
  }
}
