// Caps that keep a hostile or simply huge file from exhausting memory in the extension host or the webview.
// Shared by the host (size checks, ZIP preflight) and the viewers (XML parsing), so nothing here may import Node APIs.

const MB = 1024 * 1024;

/** Largest file each format opens; bigger files are refused before they are read. */
export const MAX_FILE_SIZE = {
  pdf: 256 * MB,
  docx: 100 * MB,
  doc: 64 * MB,
  xlsx: 100 * MB,
  xls: 64 * MB,
  pptx: 200 * MB,
  ppt: 128 * MB,
  // Every field becomes a string in the webview, so a text file costs several times its size there.
  csv: 64 * MB,
  tsv: 64 * MB,
} as const;

/** Bounds on an OOXML package (.docx/.xlsx/.pptx) checked from its ZIP central directory before a viewer unpacks it. */
export const OOXML_ZIP_LIMITS = {
  maxEntries: 20_000,
  maxEntrySize: 256 * MB,
  maxTotalSize: 1024 * MB,
  /** Declared uncompressed/compressed size of one entry; deflate cannot exceed about 1032:1, so more is a forged header. */
  maxCompressionRatio: 1100,
};

/**
 * Most pages a PDF viewer lays out: it creates a placeholder per page, and a tiny file's page tree can claim any
 * page count.
 */
export const MAX_PDF_PAGES = 50_000;

/** Largest single XML part a viewer parses into a DOM (larger parts must be streamed or refused). */
export const MAX_XML_PART_SIZE = 64 * MB;

export function formatMegabytes(bytes: number): string {
  return `${Math.round(bytes / MB)} MB`;
}
