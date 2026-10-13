import type * as vscode from 'vscode';
import { DocumentLensProvider } from './documentProvider.ts';

export function activate(context: vscode.ExtensionContext): void {
  context.subscriptions.push(...DocumentLensProvider.register(context));
}

export function deactivate(): void {}
