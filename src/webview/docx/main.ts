import '../shell/shell.css';
import './docx.css';
import { startViewer } from '../shell/app.ts';
import { docxViewer } from './viewer.ts';

startViewer(docxViewer());
