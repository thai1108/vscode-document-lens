// Caps that keep a hostile or simply huge workbook from exhausting the webview. Past a cap the viewer shows what it
// has read so far with a notice instead of failing; ZIP and XML part sizes are bounded by src/limits.ts.

export const SHEET_LIMITS = {
  /** Excel's own grid. */
  maxRows: 1_048_576,
  maxCols: 16_384,
  /** Cells kept per sheet (values and formatted blanks); later rows are dropped with a notice. */
  maxCells: 6_000_000,
  /** Formula texts kept per sheet for the cell bar; cells past it still show their value. */
  maxFormulas: 2_000_000,
  maxSheets: 1_000,
  maxSharedStrings: 4_000_000,
  /** Excel's limit on the text of one cell. */
  maxStringLength: 32_767,
  maxRichRuns: 256,
  maxStyles: 65_536,
  maxMerges: 100_000,
  maxLinks: 100_000,
  /** Rows with their own height or visibility. */
  maxRowEntries: 1_048_576,
  /** One XML tag or text run; anything longer is not a spreadsheet part. */
  maxXmlToken: 16 * 1024 * 1024,
  /** Rendered height of the grid; Chrome cannot lay out much taller elements (2^25 px). */
  maxCanvasPx: 30_000_000,
} as const;
