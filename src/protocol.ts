// Messages between the extension host and a viewer webview. The webview renders untrusted files, so everything it
// posts is validated here before the host acts on it.

/** Extension host → webview. */
export type HostMessage =
  /** The file to show; sent again (with new bytes) whenever the file changes on disk. */
  | { type: 'document'; format: string; fileName: string; data: Uint8Array }
  /** The file could not be read or failed its preflight check. */
  | { type: 'error'; fileName: string; message: string };

/** Webview → extension host. */
export type WebviewMessage =
  | { type: 'ready' }
  /** A link the user clicked in the document; opened only if it is http(s). */
  | { type: 'openLink'; url: string };

export const MESSAGE_LIMITS = {
  urlLength: 8192,
};

/** Validates a message from the webview at runtime; anything malformed or oversized is dropped. */
export function parseWebviewMessage(raw: unknown): WebviewMessage | null {
  if (typeof raw !== 'object' || raw === null) {
    return null;
  }
  const message = raw as Record<string, unknown>;
  switch (message['type']) {
    case 'ready':
      return { type: 'ready' };
    case 'openLink': {
      const url = externalUrl(message['url']);
      return url ? { type: 'openLink', url } : null;
    }
    default:
      return null;
  }
}

/** The normalized URL if `value` is an absolute http(s) URL without embedded credentials, else null. */
export function externalUrl(value: unknown): string | null {
  if (typeof value !== 'string' || value.length > MESSAGE_LIMITS.urlLength) {
    return null;
  }
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    return null;
  }
  if ((url.protocol !== 'http:' && url.protocol !== 'https:') || url.username || url.password) {
    return null;
  }
  return url.href;
}
