// Shift_JIS encoding for test files. Node can decode Shift_JIS (TextDecoder) but not encode it, so the table is the
// decoder's own mapping turned around: every lead/trail byte pair is decoded once and the first pair giving a
// character wins, as in the WHATWG encoder (which skips the NEC-selected IBM duplicates at pointers 8272–8835).

let table: Map<number, number> | undefined;

function encoderTable(): Map<number, number> {
  if (table) {
    return table;
  }
  table = new Map();
  const decoder = new TextDecoder('shift_jis', { fatal: true });
  for (let pointer = 0; pointer < 11280; pointer++) {
    if (pointer >= 8272 && pointer <= 8835) {
      continue;
    }
    const lead = Math.floor(pointer / 188);
    const trail = pointer % 188;
    const bytes = [lead + (lead < 0x1f ? 0x81 : 0xc1), trail + (trail < 0x3f ? 0x40 : 0x41)];
    let text: string;
    try {
      text = decoder.decode(new Uint8Array(bytes));
    } catch {
      continue;
    }
    const code = text.codePointAt(0);
    if (code !== undefined && text.length === 1 && !table.has(code)) {
      table.set(code, (bytes[0] ?? 0) * 256 + (bytes[1] ?? 0));
    }
  }
  return table;
}

/** Encodes text as Shift_JIS; throws for a character Shift_JIS cannot hold. */
export function encodeShiftJis(text: string): Uint8Array {
  const out: number[] = [];
  const codes = encoderTable();
  for (const char of text) {
    const code = char.codePointAt(0) ?? 0;
    if (code < 0x80) {
      out.push(code);
    } else if (code >= 0xff61 && code <= 0xff9f) {
      out.push(code - 0xff61 + 0xa1);
    } else {
      const pair = codes.get(code === 0x2212 ? 0xff0d : code);
      if (pair === undefined) {
        throw new Error(`${char} cannot be written in Shift_JIS`);
      }
      out.push(pair >> 8, pair & 0xff);
    }
  }
  return Uint8Array.from(out);
}
