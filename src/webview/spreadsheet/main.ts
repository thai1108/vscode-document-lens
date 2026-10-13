import '../shell/shell.css';
import './spreadsheet.css';
import { startViewer } from '../shell/app.ts';
import { spreadsheetViewer } from './viewer.ts';

startViewer(spreadsheetViewer());
