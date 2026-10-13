// Pictures and fonts of a parsed document (docx-preview's WordDocument). docx-preview inflates a picture and makes a
// new blob: URL for it every time the document refers to it (a header picture is loaded again for every page), so a
// small file repeating one large picture could fill the webview's memory. Here every path is inflated once, gets one
// Blob and one URL, and the distinct pictures and fonts are held to DOCX_LIMITS.maxMediaSize together.
import { DOCX_LIMITS } from './limits.ts';

/** What docx-preview's WordDocument does with its package's binary parts; checked at runtime. */
interface WordDocumentMedia {
  _package?: { load?: (path: string, type?: string) => Promise<unknown> };
  blobToURL?: (blob: Blob | null, path?: string) => unknown;
}

export interface MediaLoads {
  /** Every blob: URL made for the document, recorded as it is made (revoke them when the render is replaced). */
  readonly objectUrls: string[];
  /** Whether pictures or fonts were left out because together they exceeded DOCX_LIMITS.maxMediaSize. */
  readonly overBudget: boolean;
}

/**
 * Makes `parsed` load each binary part once, using `sizes` (the entries' verified sizes) for the budget. Call it
 * between docx-preview's parseAsync and renderDocument.
 */
export function shareMedia(parsed: unknown, sizes: ReadonlyMap<string, number>, budget: number = DOCX_LIMITS.maxMediaSize): MediaLoads {
  const doc = parsed as WordDocumentMedia;
  const packageParts = doc._package;
  const load = packageParts?.load;
  const toUrl = doc.blobToURL;
  if (!packageParts || typeof load !== 'function' || typeof toUrl !== 'function') {
    throw new Error('docx-preview no longer exposes the package loader this viewer relies on');
  }
  const objectUrls: string[] = [];
  const state = { objectUrls, overBudget: false };
  const loads = new Map<string, Promise<unknown>>();
  const counted = new Set<string>();
  let total = 0;

  packageParts.load = (path: string, type = 'string') => {
    if (type !== 'blob' && type !== 'uint8array') {
      return load.call(packageParts, path, type);
    }
    // OpenXmlPackage.get drops a leading slash.
    const name = path.startsWith('/') ? path.slice(1) : path;
    if (!counted.has(name)) {
      const size = sizes.get(name) ?? 0;
      if (total + size > budget) {
        state.overBudget = true;
        return Promise.resolve(null);
      }
      counted.add(name);
      total += size;
    }
    const key = `${type}:${name}`;
    let pending = loads.get(key);
    if (!pending) {
      pending = load.call(packageParts, path, type);
      loads.set(key, pending);
    }
    // docx-preview deobfuscates an embedded font's bytes in place, so every use gets its own copy.
    return type === 'uint8array' ? pending.then((bytes) => (bytes instanceof Uint8Array ? bytes.slice() : bytes)) : pending;
  };

  const urls = new Map<string, unknown>();
  doc.blobToURL = (blob: Blob | null, path?: string) => {
    if (!blob) {
      return null;
    }
    if (path !== undefined && urls.has(path)) {
      return urls.get(path);
    }
    const url = toUrl.call(doc, blob, path);
    if (typeof url === 'string') {
      objectUrls.push(url);
    }
    if (path !== undefined) {
      urls.set(path, url);
    }
    return url;
  };
  return state;
}
