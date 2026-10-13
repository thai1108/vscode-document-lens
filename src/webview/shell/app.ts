// The frame every viewer runs in: toolbar (file name, viewer controls, zoom), notice/error banners and the scrolling
// content area. A viewer only turns the file's bytes into DOM inside `content`; the shell talks to the host.
import { externalUrl, type HostMessage } from '../../protocol.ts';
import { element } from './dom.ts';
import { connectToHost, type ShellState, type VsCodeApi } from './vscodeApi.ts';
import { clampZoom, formatZoom, stepZoom } from './zoom.ts';

export interface LoadedDocument {
  format: string;
  fileName: string;
  data: Uint8Array;
}

export interface ViewerContext {
  /** The scrolling area the viewer renders into; the viewer owns its children. */
  readonly content: HTMLElement;
  /** Toolbar slot for the viewer's own controls (page number, sheet picker, ...). */
  readonly tools: HTMLElement;
  /** Nonce for <style> elements the viewer creates (the CSP refuses styles without it). */
  readonly styleNonce: string;
  zoom(): number;
  setZoom(zoom: number): void;
  /** The viewer's part of the persisted webview state. */
  viewerState<T>(): T | undefined;
  saveViewerState(state: unknown): void;
  /** Banner under the toolbar; null removes it. */
  showNotice(text: string | null): void;
  /** The standard banner for formats whose preview leaves out formatting/layout. */
  showSimplifiedNotice(detail: string): void;
  showError(message: string): void;
  /** Asks the host to open an http(s) URL in the browser; anything else is ignored. */
  openLink(url: string): void;
}

export interface Viewer {
  /** Renders the document; called again with the new bytes when the file changes on disk (`reload` true). */
  render(document: LoadedDocument, context: ViewerContext, reload: boolean): void | Promise<void>;
  /** Applies the shared zoom; without it the shell scales the content with CSS zoom. */
  applyZoom?(zoom: number): void;
  /** Zoom at which the document fills the content width; shows the Fit width button. */
  fitWidthZoom?(): number | undefined;
}

export interface Shell {
  readonly context: ViewerContext;
  dispose(): void;
}

const SIMPLIFIED_PREFIX = 'Simplified preview: ';

export function startViewer(viewer: Viewer, api: VsCodeApi = connectToHost()): Shell {
  const app = document.getElementById('app');
  if (!app) {
    throw new Error('#app missing from the webview page');
  }
  const title = element('span', { className: 'file-name', attributes: { id: 'file-name' } });
  const tools = element('div', { className: 'viewer-tools', attributes: { id: 'viewer-tools' } });
  const zoomOut = element('button', { className: 'icon-button', text: '−', attributes: { id: 'zoom-out', title: 'Zoom out (-)', 'aria-label': 'Zoom out' } });
  const zoomValue = element('button', { className: 'zoom-value', text: '100%', attributes: { id: 'zoom-value', title: 'Actual size (0)' } });
  const zoomIn = element('button', { className: 'icon-button', text: '+', attributes: { id: 'zoom-in', title: 'Zoom in (+)', 'aria-label': 'Zoom in' } });
  const fitWidth = element('button', { className: 'text-button', text: 'Fit width', attributes: { id: 'fit-width', title: 'Fit width (W)' } });
  fitWidth.hidden = !viewer.fitWidthZoom;
  const zoomGroup = element('div', { className: 'zoom-group' });
  zoomGroup.append(zoomOut, zoomValue, zoomIn, fitWidth);
  const toolbar = element('header', { className: 'toolbar', attributes: { id: 'toolbar' } });
  toolbar.append(title, tools, element('div', { className: 'spacer' }), zoomGroup);
  const notice = element('div', { className: 'banner notice', attributes: { id: 'notice', role: 'status' } });
  const error = element('div', { className: 'banner error', attributes: { id: 'error', role: 'alert' } });
  notice.hidden = true;
  error.hidden = true;
  const content = element('main', { className: 'content', attributes: { id: 'content', tabindex: '0' } });
  app.replaceChildren(toolbar, notice, error, content);

  let state: ShellState = api.getState() ?? {};
  let zoom = clampZoom(state.zoom ?? 1);
  let rendered = false;
  let renderQueue = Promise.resolve();
  const save = (patch: ShellState) => {
    state = { ...state, ...patch };
    api.setState(state);
  };

  const applyZoom = () => {
    zoomValue.textContent = formatZoom(zoom);
    if (viewer.applyZoom) {
      viewer.applyZoom(zoom);
    } else {
      content.style.setProperty('zoom', String(zoom));
    }
  };

  const context: ViewerContext = {
    content,
    tools,
    styleNonce: scriptNonce(),
    zoom: () => zoom,
    setZoom(next) {
      zoom = clampZoom(next);
      save({ zoom });
      applyZoom();
    },
    viewerState: <T>() => state.viewer as T | undefined,
    saveViewerState: (viewerState) => save({ viewer: viewerState }),
    showNotice(text) {
      notice.textContent = text ?? '';
      notice.hidden = text === null;
    },
    showSimplifiedNotice(detail) {
      context.showNotice(SIMPLIFIED_PREFIX + detail);
    },
    showError(message) {
      error.textContent = message;
      error.hidden = false;
    },
    openLink(url) {
      const external = externalUrl(url);
      if (external) {
        api.postMessage({ type: 'openLink', url: external });
      }
    },
  };

  const show = (message: HostMessage) => {
    title.textContent = message.fileName;
    title.title = message.fileName;
    if (message.type === 'error') {
      context.showError(message.message);
      return;
    }
    const reload = rendered;
    rendered = true;
    // Renders run one at a time so a quick second save on disk cannot interleave with the first render.
    renderQueue = renderQueue.then(async () => {
      try {
        await viewer.render({ format: message.format, fileName: message.fileName, data: message.data }, context, reload);
        error.hidden = true;
        applyZoom();
      } catch (failure) {
        context.showError(`Could not display ${message.fileName}: ${(failure as Error).message}`);
      }
    });
  };

  const onMessage = (event: MessageEvent) => {
    const message = event.data as HostMessage | undefined;
    if (message?.type === 'document' || message?.type === 'error') {
      show(message);
    }
  };
  const onKey = (event: KeyboardEvent) => {
    const target = event.target as Element | null;
    if (event.metaKey || event.ctrlKey || event.altKey || target?.closest?.('input, textarea, select, [contenteditable="true"]')) {
      return;
    }
    if (event.key === '+' || event.key === '=') {
      context.setZoom(stepZoom(zoom, 1));
    } else if (event.key === '-') {
      context.setZoom(stepZoom(zoom, -1));
    } else if (event.key === '0') {
      context.setZoom(1);
    } else if ((event.key === 'w' || event.key === 'W') && viewer.fitWidthZoom) {
      fitToWidth();
    } else {
      return;
    }
    event.preventDefault();
  };
  // Pinch gestures arrive as wheel events with ctrlKey set.
  const onWheel = (event: WheelEvent) => {
    if (event.ctrlKey || event.metaKey) {
      event.preventDefault();
      context.setZoom(zoom * Math.exp(-event.deltaY / 300));
    }
  };
  const onClick = (event: MouseEvent) => {
    const link = (event.target as Element | null)?.closest?.('a[href]');
    const href = link?.getAttribute('href');
    if (!href || href.startsWith('#')) {
      return;
    }
    // Links in a document never navigate the webview; http(s) ones open in the browser through the host.
    event.preventDefault();
    event.stopPropagation();
    context.openLink(href);
  };
  const fitToWidth = () => {
    const fit = viewer.fitWidthZoom?.();
    if (fit) {
      context.setZoom(fit);
    }
  };

  zoomOut.addEventListener('click', () => context.setZoom(stepZoom(zoom, -1)));
  zoomIn.addEventListener('click', () => context.setZoom(stepZoom(zoom, 1)));
  zoomValue.addEventListener('click', () => context.setZoom(1));
  fitWidth.addEventListener('click', fitToWidth);
  window.addEventListener('message', onMessage);
  window.addEventListener('keydown', onKey);
  content.addEventListener('wheel', onWheel, { passive: false });
  document.addEventListener('click', onClick, true);
  zoomValue.textContent = formatZoom(zoom);
  api.postMessage({ type: 'ready' });

  return {
    context,
    dispose() {
      window.removeEventListener('message', onMessage);
      window.removeEventListener('keydown', onKey);
      content.removeEventListener('wheel', onWheel);
      document.removeEventListener('click', onClick, true);
    },
  };
}

/** Chrome hides the nonce attribute once the page has loaded and keeps it in the `nonce` property. */
function scriptNonce(): string {
  const script = document.querySelector<HTMLScriptElement>('script[nonce]');
  return script?.nonce || script?.getAttribute('nonce') || '';
}
