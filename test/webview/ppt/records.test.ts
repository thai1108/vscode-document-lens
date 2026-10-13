import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { PptError } from '../../../src/webview/ppt/errors.ts';
import { RecordStream } from '../../../src/webview/ppt/records.ts';
import { atom, container, recordHeader, u32 } from '../../support/pptWriter.ts';

const stream = (bytes: Uint8Array, records = 1000) => new RecordStream(new Uint8Array(bytes), { records });

describe('RecordStream', () => {
  it('splits the header into version, instance, type and data range', () => {
    const records = stream(Buffer.concat([atom(0x0f9f, u32(1), 0x123, 2)]));
    assert.deepEqual(records.at(0), { version: 2, instance: 0x123, type: 0x0f9f, offset: 0, start: 8, end: 12 });
    assert.equal(records.u32(8), 1);
  });

  it('lists the children of a container and finds them by type and instance', () => {
    const bytes = container(0x0ff0, [atom(0x03f3, u32(5)), atom(0x0f9f, u32(0), 0), atom(0x0f9f, u32(1), 1)]);
    const records = stream(bytes);
    const parent = records.at(0);
    assert.deepEqual(
      records.children(parent).map((child) => child.type),
      [0x03f3, 0x0f9f, 0x0f9f],
    );
    assert.equal(records.child(parent, 0x0f9f, 1)?.offset, 8 + 12 + 12);
    assert.equal(records.child(parent, 0x0fa0), undefined);
  });

  it('ignores trailing bytes too short for a header', () => {
    const bytes = Buffer.concat([container(0x03e8, [atom(0x03ea, Buffer.alloc(0))]), Buffer.alloc(3)]);
    const records = stream(bytes);
    assert.equal(records.children({ start: 8, end: bytes.length }).length, 1);
  });

  it('reports a record whose length runs past its container or stream', () => {
    const child = atom(0x0f9f, u32(1));
    child.writeUInt32LE(5000, 4);
    const records = stream(container(0x0ff0, [child]));
    assert.throws(() => records.children(records.at(0)), (error: Error) => error instanceof PptError && error.kind === 'damaged' && /runs past the end/.test(error.message));
    const huge = recordHeader(0x03e8, 0xffffffff, 0, 0xf);
    assert.throws(() => stream(huge).at(0), /runs past the end/);
  });

  it('reports headers outside the stream', () => {
    const records = stream(atom(0x03ea, Buffer.alloc(0)));
    assert.throws(() => records.at(4), /lies outside its stream/);
    assert.throws(() => records.at(-1), /lies outside its stream/);
    assert.throws(() => records.at(0.5), /lies outside its stream/);
  });

  it('stops after the record budget', () => {
    const bytes = container(0x0ff0, Array.from({ length: 10 }, () => atom(0x03ea, Buffer.alloc(0))));
    const records = stream(bytes, 5);
    assert.throws(() => records.children(records.at(0)), (error: Error) => error instanceof PptError && error.kind === 'too-complex');
  });

  it('checks that an atom is long enough before its fields are read', () => {
    const records = stream(atom(0x0ff5, u32(1, 2)));
    assert.throws(() => records.expect(records.at(0), 28, 'an edit record'), /an edit record is too short/);
  });
});
