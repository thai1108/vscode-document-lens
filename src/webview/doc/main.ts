import '../shell/shell.css';
import './doc.css';
import { startViewer } from '../shell/app.ts';
import { docViewer } from './viewer.ts';

startViewer(docViewer());
