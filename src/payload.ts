import type { FormatDefinition } from './formats/types.ts';
import type { HostMessage } from './protocol.ts';

/**
 * The document message for a file's bytes.
 *
 * The bytes go out as a plain, exactly-sized Uint8Array: VS Code's webview serializer JSON.stringifies the message
 * with a replacer that turns typed arrays into transferred buffers, but a Node Buffer's toJSON() runs first (it would
 * arrive as `{ type: 'Buffer', data: [...] }`), and a view into a larger buffer would transfer all of that buffer.
 */
export function documentMessage(format: FormatDefinition, fileName: string, bytes: Uint8Array): HostMessage {
  return { type: 'document', format: format.id, fileName, data: plainBytes(bytes) };
}

export function plainBytes(bytes: Uint8Array): Uint8Array {
  const exact = bytes.constructor === Uint8Array && bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength;
  return exact ? bytes : new Uint8Array(bytes);
}
