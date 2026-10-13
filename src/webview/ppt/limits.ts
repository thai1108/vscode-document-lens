// Caps that keep a hostile .ppt from exhausting the webview: every count or size a file declares is checked against
// these before it is trusted. Real presentations stay far below them.

const MB = 1024 * 1024;

export const PPT_LIMITS = {
  /** Directory entries read from the compound file (each is 128 bytes; real files have a few dozen). */
  maxDirectoryEntries: 100_000,
  /** The Current User stream is normally under 200 bytes. */
  maxCurrentUserSize: 64 * 1024,
  /** UserEditAtoms followed back through the edit history (one per incremental save). */
  maxEdits: 10_000,
  /** Persist object references collected from all persist directories. */
  maxPersistEntries: 200_000,
  /** Records visited while parsing one file. */
  maxRecords: 2_000_000,
  /** Container nesting (groups of groups of shapes); deeper content is left out. */
  maxDepth: 48,
  maxSlides: 10_000,
  /** Characters of slide text kept across the whole presentation. */
  maxTextChars: 4_000_000,
  /** Picture elements shown across the whole presentation. */
  maxPictures: 2_000,
  /** One picture's data. */
  maxPictureSize: 64 * MB,
  /** All distinct pictures' data together. */
  maxTotalPictureSize: 256 * MB,
  /** Decoded size of one picture (width × height); bigger images are not handed to the browser's decoder. */
  maxPicturePixels: 100_000_000,
} as const;

export type PptLimits = Readonly<Record<keyof typeof PPT_LIMITS, number>>;
