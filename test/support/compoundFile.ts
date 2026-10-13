/**
 * Only the 512-byte header of an OLE compound file (version 3): enough for the host's preflight signature check,
 * not a readable file. Viewers that parse .doc/.xls/.ppt need real fixtures from a compound file writer.
 */
export function compoundFileHeaderStub(): Buffer {
  const header = Buffer.alloc(512);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(header, 0);
  header.writeUInt16LE(0x003e, 24);
  header.writeUInt16LE(0x0003, 26);
  header.writeUInt16LE(0xfffe, 28);
  header.writeUInt16LE(9, 30);
  header.writeUInt16LE(6, 32);
  return header;
}
