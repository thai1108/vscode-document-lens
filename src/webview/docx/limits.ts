// Caps on what a .docx can make the viewer build, on top of the host's ZIP checks and MAX_XML_PART_SIZE
// (src/limits.ts). docx-preview parses every part and renders the whole document at once, recursively, so the XML is
// budgeted before it is parsed and the parsed model is measured before it is turned into DOM.
export const DOCX_LIMITS = {
  /** All the XML parts docx-preview parses, together (each is also within MAX_XML_PART_SIZE). */
  maxXmlTotalSize: 96 * 1024 * 1024,
  /** Elements in all those parts together. */
  maxXmlElements: 4_000_000,
  /** Pictures and fonts the document refers to, each counted once. */
  maxMediaSize: 256 * 1024 * 1024,
  /** Estimated elements the render would create (body, notes, and a header and footer per page). */
  maxRenderedNodes: 2_000_000,
  /** Nesting of the parsed model (tables in tables, text boxes, ...). */
  maxDepth: 200,
  /** Largest absolute length a style may use, in points. Word's own maximum page and picture size is 22 in. */
  maxLengthPt: 5000,
  /** Largest em/rem/ex/ch, vw/vh and % values. */
  maxRelativeLength: 100,
  maxPercentage: 1000,
  /** Largest unitless line height (a multiple of the font size). */
  maxLineHeight: 20,
  maxColumns: 45,
} as const;
