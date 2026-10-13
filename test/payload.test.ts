import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { pdf } from '../src/formats/pdf.ts';
import { documentMessage, plainBytes } from '../src/payload.ts';
import type { HostMessage } from '../src/protocol.ts';
import { deserializeInWebview, serializeForWebview } from './support/vscodeWebviewSerializer.ts';

function delivered(message: HostMessage): HostMessage {
  return deserializeInWebview(serializeForWebview(message)) as HostMessage;
}

describe('documentMessage', () => {
  it('reaches the webview as an exactly sized Uint8Array, also when read into a Node Buffer view', () => {
    const file = Buffer.from('%PDF-1.4 example');
    const pooled = Buffer.concat([Buffer.from('other data before '), file]).subarray(18);
    for (const bytes of [file, pooled, new Uint8Array(file)]) {
      const sent = documentMessage(pdf, 'a.pdf', bytes);
      const serialized = serializeForWebview(sent);
      assert.equal(serialized.buffers.length, 1);
      assert.equal(serialized.buffers[0]?.byteLength, file.byteLength, 'only the file is transferred');
      const received = delivered(sent);
      assert.equal(received.type, 'document');
      if (received.type === 'document') {
        assert.equal(received.data.constructor, Uint8Array);
        assert.deepEqual([...received.data], [...file]);
        assert.equal(received.format, 'pdf');
        assert.equal(received.fileName, 'a.pdf');
      }
    }
  });
});

describe('plainBytes', () => {
  it('keeps an exact plain Uint8Array and copies anything else', () => {
    const exact = new Uint8Array([1, 2, 3]);
    assert.equal(plainBytes(exact), exact);
    const buffer = Buffer.from([1, 2, 3]);
    assert.notEqual(plainBytes(buffer), buffer);
    assert.equal(plainBytes(buffer).constructor, Uint8Array);
    const view = new Uint8Array(new ArrayBuffer(8), 2, 3);
    assert.equal(plainBytes(view).buffer.byteLength, 3);
  });
});
