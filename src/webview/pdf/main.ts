import workerSource from 'pdfjs-assets:worker';
import files from 'pdfjs-assets:files';
import '../shell/shell.css';
import './pdf.css';
import { startViewer } from '../shell/app.ts';
import { startWorker } from './assets.ts';
import { createPdfViewer } from './viewer.ts';

startViewer(createPdfViewer({ startWorker: () => startWorker(workerSource), files }));
