// Runs the webview shell (src/webview/shell) in a happy-dom window with a fake VS Code host. happy-dom does no
// layout, so these tests cover DOM wiring and state; rendering is covered by the Playwright suite (test/e2e).
import { Window } from 'happy-dom';
import type { HostMessage, WebviewMessage } from '../../../src/protocol.ts';
import { startViewer, type Shell, type Viewer } from '../../../src/webview/shell/app.ts';
import type { ShellState } from '../../../src/webview/shell/vscodeApi.ts';
import { deserializeInWebview, serializeForWebview } from '../../support/vscodeWebviewSerializer.ts';

export interface FakeHost {
  window: Window;
  shell: Shell;
  posted: WebviewMessage[];
  state: ShellState | undefined;
  /** Delivers a host message the way VS Code does (serialized in the extension host, revived in the webview). */
  deliver(message: HostMessage): void;
  until(predicate: () => boolean, what: string, timeoutMs?: number): Promise<void>;
  $<E extends Element = HTMLElement>(selector: string): E;
  click(target: Element): MouseEvent;
  key(key: string, modifiers?: Modifiers, target?: EventTarget): void;
  close(): Promise<void>;
}

type Modifiers = Partial<Record<'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey', boolean>>;

const GLOBALS = ['window', 'document', 'HTMLElement', 'HTMLScriptElement', 'KeyboardEvent', 'MouseEvent', 'WheelEvent', 'MessageEvent', 'getComputedStyle'] as const;

export function documentMessage(fileName: string, text: string, format = 'pdf'): HostMessage {
  return { type: 'document', format, fileName, data: new TextEncoder().encode(text) };
}

export async function bootShell(viewer: Viewer, options: { state?: ShellState } = {}): Promise<FakeHost> {
  const window = new Window({ url: 'https://webview.test/', width: 1200, height: 800 });
  const scope = globalThis as unknown as Record<string, unknown>;
  const source = window as unknown as Record<string, unknown>;
  for (const name of GLOBALS) {
    scope[name] = name === 'window' ? window : source[name];
  }
  const document = window.document as unknown as Document;
  document.body.innerHTML = '<div id="app"></div><script nonce="style-nonce" src="webview.js"></script>';

  const posted: WebviewMessage[] = [];
  const host = {
    window,
    posted,
    state: options.state,
  } as FakeHost;
  host.shell = startViewer(viewer, {
    postMessage: (message) => posted.push(structuredClone(message)),
    getState: () => host.state,
    setState: (state) => {
      host.state = structuredClone(state);
    },
  });
  Object.assign(host, {
    deliver(message: HostMessage) {
      const data = deserializeInWebview(serializeForWebview(message));
      window.dispatchEvent(new window.MessageEvent('message', { data }));
    },
    async until(predicate: () => boolean, what: string, timeoutMs = 2000) {
      const deadline = Date.now() + timeoutMs;
      while (!predicate()) {
        if (Date.now() > deadline) {
          throw new Error(`Timed out waiting for ${what}`);
        }
        await new Promise((resolve) => setTimeout(resolve, 5));
      }
    },
    $(selector: string) {
      const found = document.querySelector(selector);
      if (!found) {
        throw new Error(`No element matches ${selector}`);
      }
      return found;
    },
    click(target: Element) {
      const event = new window.MouseEvent('click', { bubbles: true, cancelable: true });
      target.dispatchEvent(event as unknown as Event);
      return event as unknown as MouseEvent;
    },
    key(key: string, modifiers: Modifiers = {}, target: EventTarget = window as unknown as EventTarget) {
      target.dispatchEvent(new window.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...modifiers }) as unknown as Event);
    },
    async close() {
      host.shell.dispose();
      await window.happyDOM.close();
    },
  });
  return host;
}
