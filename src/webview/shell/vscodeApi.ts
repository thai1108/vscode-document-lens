import type { WebviewMessage } from '../../protocol.ts';

/** What the shell keeps across webview reloads; `viewer` belongs to the format's viewer. */
export interface ShellState {
  zoom?: number;
  viewer?: unknown;
}

export interface VsCodeApi {
  postMessage(message: WebviewMessage): void;
  getState(): ShellState | undefined;
  setState(state: ShellState): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

export function connectToHost(): VsCodeApi {
  return acquireVsCodeApi();
}
