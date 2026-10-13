// Pictures from the package as blob: URLs (the only image source the webview's CSP allows besides data:). The type
// is sniffed from the bytes, never taken from the file name; formats a browser cannot draw (EMF, WMF, TIFF) are
// reported by name so the slide can show a labelled placeholder instead.
import { PPTX_LIMITS } from './limits.ts';
import type { PptxPackage } from './package.ts';

export type Picture = { readonly url: string; readonly width?: number; readonly height?: number } | { readonly unsupported: string };

const BROWSER_TYPES: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
};

const UNSUPPORTED_NAMES: Readonly<Record<string, string>> = {
  emf: 'EMF picture',
  wmf: 'WMF picture',
  tiff: 'TIFF picture',
  unknown: 'Picture',
};

export class MediaStore {
  private readonly pkg: PptxPackage;
  private readonly pictures = new Map<string, Promise<Picture>>();
  private readonly urls: string[] = [];
  private used = 0;

  constructor(pkg: PptxPackage) {
    this.pkg = pkg;
  }

  picture(part: string): Promise<Picture> {
    let loading = this.pictures.get(part);
    if (!loading) {
      loading = this.load(part);
      this.pictures.set(part, loading);
    }
    return loading;
  }

  private async load(part: string): Promise<Picture> {
    const size = this.pkg.size(part);
    if (size === undefined) {
      return { unsupported: 'Missing picture' };
    }
    if (size > PPTX_LIMITS.maxPictureBytes || this.used + size > PPTX_LIMITS.maxMediaBytes) {
      return { unsupported: 'Picture not shown (too large)' };
    }
    this.used += size;
    const bytes = await this.pkg.bytes(part);
    const kind = sniffImage(bytes);
    const type = BROWSER_TYPES[kind];
    if (!type) {
      return { unsupported: UNSUPPORTED_NAMES[kind] ?? 'Picture' };
    }
    const url = URL.createObjectURL(new Blob([bytes as Uint8Array<ArrayBuffer>], { type }));
    this.urls.push(url);
    return { url, ...imageSize(bytes, kind) };
  }

  /** Revokes every URL handed out; the store must not be used afterwards. */
  dispose(): void {
    for (const url of this.urls) {
      URL.revokeObjectURL(url);
    }
    this.urls.length = 0;
    this.pictures.clear();
  }
}

export function sniffImage(bytes: Uint8Array): string {
  const at = (offset: number, ...values: number[]) => values.every((value, i) => bytes[offset + i] === value);
  if (at(0, 0x89, 0x50, 0x4e, 0x47)) {
    return 'png';
  }
  if (at(0, 0xff, 0xd8, 0xff)) {
    return 'jpeg';
  }
  if (at(0, 0x47, 0x49, 0x46, 0x38)) {
    return 'gif';
  }
  if (at(0, 0x42, 0x4d)) {
    return 'bmp';
  }
  if (at(0, 0x52, 0x49, 0x46, 0x46) && at(8, 0x57, 0x45, 0x42, 0x50)) {
    return 'webp';
  }
  if (at(0, 0x01, 0x00, 0x00, 0x00) && at(40, 0x20, 0x45, 0x4d, 0x46)) {
    return 'emf';
  }
  if (at(0, 0xd7, 0xcd, 0xc6, 0x9a) || at(0, 0x01, 0x00, 0x09, 0x00) || at(0, 0x02, 0x00, 0x09, 0x00)) {
    return 'wmf';
  }
  if (at(0, 0x49, 0x49, 0x2a, 0x00) || at(0, 0x4d, 0x4d, 0x00, 0x2a)) {
    return 'tiff';
  }
  const head = new TextDecoder().decode(bytes.subarray(0, 512)).trimStart();
  if (/^(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*)*<svg[\s>]/i.test(head)) {
    return 'svg';
  }
  return 'unknown';
}

/** Pixel size from the image header, for tiled picture fills. */
export function imageSize(bytes: Uint8Array, kind: string): { width?: number; height?: number } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  try {
    if (kind === 'png' && bytes.byteLength >= 24) {
      return { width: view.getUint32(16), height: view.getUint32(20) };
    }
    if (kind === 'gif' && bytes.byteLength >= 10) {
      return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
    }
    if (kind === 'bmp' && bytes.byteLength >= 26) {
      return { width: Math.abs(view.getInt32(18, true)), height: Math.abs(view.getInt32(22, true)) };
    }
    if (kind === 'jpeg') {
      let offset = 2;
      while (offset + 9 < bytes.byteLength) {
        if (bytes[offset] !== 0xff) {
          break;
        }
        const marker = bytes[offset + 1] ?? 0;
        const length = view.getUint16(offset + 2);
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
          return { height: view.getUint16(offset + 5), width: view.getUint16(offset + 7) };
        }
        offset += 2 + length;
      }
    }
  } catch {
    // A truncated header only means the size is unknown.
  }
  return {};
}
