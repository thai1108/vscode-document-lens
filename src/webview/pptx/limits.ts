// Caps on what one presentation may make the viewer build. The ZIP and XML caps live in src/limits.ts (shared with the
// host's preflight); these bound the work per slide so a hostile file cannot freeze the webview.

const MB = 1024 * 1024;

export const PPTX_LIMITS = {
  maxSlides: 5_000,
  /** Shapes drawn on one slide, counting group members and the master's and layout's own shapes. */
  maxShapesPerSlide: 5_000,
  /**
   * Everything else built for one slide, counted in elements: path segments, gradient stops, pattern cells,
   * arrowheads, table cells. Small files can describe a lot of these (empty paths, shared theme gradients).
   */
  maxElementsPerSlide: 200_000,
  maxGroupDepth: 32,
  /** Paragraphs and text runs drawn on one slide. */
  maxRunsPerSlide: 50_000,
  maxTableCells: 10_000,
  /** Guides and path commands (each path counts as one) of one shape geometry. */
  maxGuides: 1_000,
  maxPathCommands: 20_000,
  maxGradientStops: 64,
  /** Bytes of chart and SmartArt parts read for one slide; each part is read once per presentation. */
  maxPartBytesPerSlide: 64 * MB,
  /** Picture bytes turned into blob: URLs, in total and per picture. */
  maxMediaBytes: 512 * MB,
  maxPictureBytes: 128 * MB,
};
