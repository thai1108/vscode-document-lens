import '../shell/shell.css';
import './ppt.css';
import { startViewer } from '../shell/app.ts';
import { pptViewer } from './viewer.ts';

startViewer(pptViewer());
