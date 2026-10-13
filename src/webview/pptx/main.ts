import '../shell/shell.css';
import './pptx.css';
import { startViewer } from '../shell/app.ts';
import { createPptxViewer } from './viewer.ts';

startViewer(createPptxViewer());
