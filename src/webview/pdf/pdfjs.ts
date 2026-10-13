// The one place the viewer imports pdf.js from. The viewer components (pdf_viewer.mjs) read the core library from
// globalThis.pdfjsLib, which pdf.mjs sets when it is evaluated, so the core must be imported first.
export { AnnotationEditorType, AnnotationMode, getDocument, PDFWorker, VerbosityLevel } from 'pdfjs-dist';
export type { PDFDocumentLoadingTask, PDFDocumentProxy } from 'pdfjs-dist';
export { EventBus, FindState, PDFFindController, PDFLinkService, PDFViewer } from 'pdfjs-dist/web/pdf_viewer.mjs';
