import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { PptError, type PptErrorKind } from '../../../src/webview/ppt/errors.ts';
import { PPT_LIMITS, type PptLimits } from '../../../src/webview/ppt/limits.ts';
import { parsePresentation, type Presentation, type Slide } from '../../../src/webview/ppt/presentation.ts';
import { SAMPLE_PPT, samplePicture } from '../../fixtures/ppt.ts';
import { rgbPng } from '../../support/png.ts';
import {
  atom,
  container,
  currentUserStream,
  pptFile,
  pptStreams,
  slidePersistId,
  textBoxShape,
  u32,
  writePpt,
  type PptSpec,
  type PptStreams,
  type PptWriteOptions,
} from '../../support/pptWriter.ts';

function fileOf(streams: PptStreams, extra: Record<string, Uint8Array | null> = {}): Uint8Array {
  return new Uint8Array(pptFile({ 'PowerPoint Document': streams.document, 'Current User': streams.currentUser, Pictures: streams.pictures, ...extra }));
}

function parse(spec: PptSpec, options?: PptWriteOptions, limits?: PptLimits): Presentation {
  return parsePresentation(new Uint8Array(writePpt(spec, options)), limits);
}

/** Each text block as its paragraphs joined by " / ". */
function texts(slide: Slide | undefined): string[] {
  return (slide?.texts ?? []).map((block) => block.paragraphs.map((paragraph) => paragraph.text).join(' / '));
}

function assertPptError(run: () => unknown, kind: PptErrorKind, message: RegExp): void {
  assert.throws(run, (error: Error) => {
    assert.ok(error instanceof PptError, `${error.name}: ${error.message}`);
    assert.equal(error.kind, kind);
    assert.match(error.message, message);
    return true;
  });
}

/** Offset of the persist directory entry for `persistId` (the writer emits one run of consecutive ids per save). */
function persistEntryOffset(streams: PptStreams, persistId: number, save = 0): number {
  const edit = streams.edits[save] ?? 0;
  const directory = streams.document.readUInt32LE(edit + 8 + 12);
  const firstId = streams.document.readUInt32LE(directory + 8) & 0xfffff;
  return directory + 8 + 4 + (persistId - firstId) * 4;
}

describe('parsePresentation', () => {
  it('reads every slide of the sample in order with its text, picture and hidden flag', () => {
    const presentation = parse(SAMPLE_PPT);
    assert.equal(presentation.width / presentation.height, 16 / 9);
    assert.deepEqual(presentation.warnings, []);
    const [first, second, third] = presentation.slides;
    assert.equal(presentation.slides.length, 3);

    assert.deepEqual(texts(first), ['Quarterly review', '四半期レビュー / Document Lens sample']);
    assert.deepEqual(
      first?.texts.map((block) => block.kind),
      ['title', 'body'],
    );
    assert.equal(first?.texts[1]?.paragraphs.some((paragraph) => paragraph.bullet), false, 'subtitles have no bullets');

    // The footer placeholder repeats master content and is left out.
    assert.deepEqual(texts(second), ['Highlights', 'Revenue grew 12% / New customers in Osaka / Roadmap\nsecond line / 東京オフィス開設', '注: 速報値です']);
    assert.deepEqual(
      second?.texts[1]?.paragraphs.map((paragraph) => [paragraph.level, paragraph.bullet]),
      [
        [0, true],
        [1, true],
        [0, true],
        [1, true],
      ],
    );
    assert.equal(second?.texts[2]?.kind, 'other');
    assert.equal(second?.pictures.length, 1);
    assert.equal(second?.pictures[0]?.mime, 'image/png');
    assert.deepEqual(second?.pictures[0]?.data, new Uint8Array(samplePicture()));
    assert.deepEqual([second?.pictures[0]?.width, second?.pictures[0]?.height], [120, 90]);

    assert.equal(first?.hidden, false);
    assert.equal(second?.hidden, false);
    assert.equal(third?.hidden, true);
    assert.deepEqual(texts(third), ['Backup slide', 'Only for questions']);
    assert.deepEqual(
      presentation.slides.map((slide) => [slide.number, slide.damaged]),
      [
        [1, false],
        [2, false],
        [3, false],
      ],
    );
  });

  it('reads the newest save of an incrementally saved file and keeps unchanged objects from older saves', () => {
    const edited: PptSpec = {
      ...SAMPLE_PPT,
      slides: SAMPLE_PPT.slides.map((slide, i) =>
        i === 0
          ? { placeholders: [{ type: 'centerTitle', text: 'Annual review' }, { type: 'subtitle', text: '年次レビュー' }] }
          : i === 1
            ? { ...slide, textBoxes: [{ type: 'other', text: 'Final numbers' }] }
            : { placeholders: [{ type: 'title', text: 'Not saved' }] },
      ),
    };
    const options = { update: { spec: edited, persistIds: [1, slidePersistId(1)] } };
    assert.equal(pptStreams(SAMPLE_PPT, options).edits.length, 2);
    const presentation = parse(SAMPLE_PPT, options);
    assert.deepEqual(texts(presentation.slides[0]), ['Annual review', '年次レビュー']);
    assert.deepEqual(texts(presentation.slides[1]).at(-1), 'Final numbers');
    // Slide 3's container was not saved again: the old one is still used.
    assert.equal(presentation.slides[2]?.hidden, true);
  });

  it('finds the document by scanning the stream when the Current User stream is missing or wrong', () => {
    const streams = pptStreams(SAMPLE_PPT);
    const expected = texts(parse(SAMPLE_PPT).slides[1]);
    const withoutUser = parsePresentation(fileOf(streams, { 'Current User': null }));
    assert.deepEqual(texts(withoutUser.slides[1]), expected);
    const pointingNowhere = parsePresentation(fileOf({ ...streams, currentUser: currentUserStream(3) }));
    assert.deepEqual(texts(pointingNowhere.slides[1]), expected);
    const outside = parsePresentation(fileOf({ ...streams, currentUser: currentUserStream(0x7fffffff) }));
    assert.equal(outside.slides.length, 3);
    const tooShort = parsePresentation(fileOf({ ...streams, currentUser: Buffer.alloc(20) }));
    assert.equal(tooShort.slides.length, 3);
  });

  it('finds the document by scanning when the edit history loops or its directory is missing', () => {
    const looping = pptStreams(SAMPLE_PPT);
    const edit = looping.edits[0] ?? 0;
    looping.document.writeUInt32LE(edit, edit + 8 + 8);
    assert.equal(parsePresentation(fileOf(looping)).slides.length, 3);

    const noDirectory = pptStreams(SAMPLE_PPT);
    noDirectory.document.writeUInt32LE(0, (noDirectory.edits[0] ?? 0) + 8 + 12);
    assert.equal(parsePresentation(fileOf(noDirectory)).slides.length, 3);
  });

  it('puts titles first and the other text from top to bottom', () => {
    const presentation = parse({
      slides: [
        {
          placeholders: [{ type: 'body', text: 'Body', top: 1500 }, { type: 'title', text: 'Title', top: 2500 }],
          textBoxes: [
            { type: 'other', text: 'Note at the bottom', top: 3000 },
            { type: 'other', text: 'Note at the top', top: 100 },
          ],
        },
      ],
    });
    assert.deepEqual(texts(presentation.slides[0]), ['Title', 'Note at the top', 'Body', 'Note at the bottom']);
  });

  it('shows slide-list text no shape refers to, once', () => {
    const presentation = parse({
      slides: [
        {
          placeholders: [{ type: 'title', text: 'Shown by a shape' }],
          listOnly: [{ type: 'body', text: 'Only in the list' }, { type: 'other', text: 'Shown by a shape' }],
        },
      ],
    });
    assert.deepEqual(texts(presentation.slides[0]), ['Shown by a shape', 'Only in the list']);
  });

  it('marks a slide whose container is missing as damaged and still lists its slide-list text', () => {
    const streams = pptStreams(SAMPLE_PPT);
    // Point slide 1 at the DocumentContainer, which is not a slide.
    streams.document.writeUInt32LE(0, persistEntryOffset(streams, slidePersistId(0)));
    const presentation = parsePresentation(fileOf(streams));
    assert.equal(presentation.slides[0]?.damaged, true);
    assert.deepEqual(texts(presentation.slides[0]), ['Quarterly review', '四半期レビュー / Document Lens sample']);
    assert.equal(presentation.slides[1]?.damaged, false);
  });

  it('reads pictures from the Pictures stream or the picture store, and shares one picture between slides', () => {
    const png = rgbPng(2, 2, () => [9, 9, 9]);
    const dib = Buffer.concat([u32(40, 1, 1), Buffer.from([1, 0, 24, 0]), u32(0, 4, 0, 0, 0, 0), Buffer.alloc(4)]);
    const presentation = parse({
      pictures: [
        { format: 'png', data: png },
        { format: 'png', data: png, embedded: true },
        { format: 'dib', data: dib },
        { format: 'emf', data: Buffer.alloc(100) },
        { format: 'jpeg', data: Buffer.from('not a jpeg') },
      ],
      slides: [{ pictures: [1, 2, 3] }, { pictures: [4, 5, 9, 1] }],
    });
    const [first, second] = presentation.slides;
    assert.deepEqual(
      first?.pictures.map((picture) => picture.mime),
      ['image/png', 'image/png', 'image/bmp'],
    );
    assert.deepEqual(
      second?.pictures.map((picture) => picture.note),
      ['EMF picture (not shown)', 'Damaged picture (not shown)', 'Missing picture', undefined],
    );
    assert.equal(second?.pictures[3], first?.pictures[0]);
  });

  it('shows a placeholder when the Pictures stream is missing or points nowhere', () => {
    const streams = pptStreams({ pictures: [{ format: 'png', data: rgbPng(1, 1, () => [0, 0, 0]) }], slides: [{ pictures: [1] }] });
    assert.equal(parsePresentation(fileOf(streams, { Pictures: null })).slides[0]?.pictures[0]?.note, 'Missing picture');
    const shifted = fileOf({ ...streams, pictures: Buffer.concat([Buffer.alloc(3), streams.pictures]) });
    assert.match(parsePresentation(shifted).slides[0]?.pictures[0]?.note ?? '', /Missing picture|Damaged picture/);
  });

  it('still shows the slides when the Pictures stream is damaged', () => {
    const streams = pptStreams(SAMPLE_PPT);
    const bytes = pptFile(
      { 'PowerPoint Document': streams.document, 'Current User': streams.currentUser, Pictures: streams.pictures },
      {
        patchFat: (fat, layout) => {
          const start = layout.streams.get('Pictures')?.start ?? 0;
          fat[start + 2] = start;
        },
      },
    );
    const presentation = parsePresentation(new Uint8Array(bytes));
    assert.equal(presentation.slides.length, 3);
    assert.equal(presentation.slides[1]?.pictures[0]?.note, 'Missing picture');
    assert.deepEqual(presentation.warnings, ['The pictures are damaged and left out.']);
  });

  it('reads a presentation without slides', () => {
    const presentation = parse({ slides: [] });
    assert.deepEqual(presentation.slides, []);
    assert.deepEqual([presentation.width, presentation.height], [5760, 4320]);
  });

  it('falls back to 4:3 when the slide size is implausible', () => {
    const presentation = parse({ slides: [{}], slideSize: { width: 5760, height: 0 } });
    assert.deepEqual([presentation.width, presentation.height], [5760, 4320]);
  });
});

describe('parsePresentation limits', () => {
  it('cuts text after the character limit and says so', () => {
    const presentation = parse(SAMPLE_PPT, {}, { ...PPT_LIMITS, maxTextChars: 20 });
    assert.deepEqual(texts(presentation.slides[0]), ['Quarterly review', '四半期レ']);
    assert.deepEqual(texts(presentation.slides[1]), []);
    assert.deepEqual(presentation.warnings, ['Text after the first 20 characters is left out.']);
  });

  it('shows at most the picture limit, and leaves out pictures beyond the data limit', () => {
    const png = rgbPng(2, 2, () => [0, 0, 0]);
    const spec: PptSpec = { pictures: [{ format: 'png', data: png }, { format: 'png', data: rgbPng(3, 3, () => [1, 1, 1]) }], slides: [{ pictures: [1, 2, 1] }] };
    const counted = parse(spec, {}, { ...PPT_LIMITS, maxPictures: 2 });
    assert.equal(counted.slides[0]?.pictures.length, 2);
    assert.deepEqual(counted.warnings, ['Only the first 2 pictures are shown.']);
    const sized = parse(spec, {}, { ...PPT_LIMITS, maxTotalPictureSize: png.length + 10 });
    assert.deepEqual(
      sized.slides[0]?.pictures.map((picture) => picture.note),
      [undefined, 'Picture left out (too much picture data)', undefined],
    );
    assert.deepEqual(sized.warnings, ['Some pictures are left out: the presentation holds too much picture data.']);
  });

  it('shows at most the slide limit', () => {
    const presentation = parse(SAMPLE_PPT, {}, { ...PPT_LIMITS, maxSlides: 2 });
    assert.equal(presentation.slides.length, 2);
    assert.deepEqual(presentation.warnings, ['Only the first 2 slides are shown.']);
  });

  it('leaves out shapes nested deeper than the limit and keeps the rest of the slide', () => {
    let nested = textBoxShape({ type: 'other', text: 'Too deep' }, 100);
    for (let i = 0; i < 200; i++) {
      nested = container(0xf003, [container(0xf004, [atom(0xf009, Buffer.alloc(16), 0, 1), atom(0xf00a, u32(0, 0x201), 0, 2)]), nested]);
    }
    const presentation = parse({ slides: [{ placeholders: [{ type: 'title', text: 'Still here' }], extraShapes: [nested] }] });
    assert.deepEqual(texts(presentation.slides[0]), ['Still here']);
    assert.deepEqual(presentation.warnings, ['Some deeply nested shapes are left out.']);
  });

  it('refuses a file with more records than the limit', () => {
    assertPptError(() => parse(SAMPLE_PPT, {}, { ...PPT_LIMITS, maxRecords: 50 }), 'too-complex', /too many records/);
  });
});

describe('parsePresentation with files it cannot show', () => {
  it('explains password-protected files', () => {
    assertPptError(() => parse({ ...SAMPLE_PPT, encrypted: { token: true } }), 'encrypted', /password-protected/);
    assertPptError(() => parse({ ...SAMPLE_PPT, encrypted: { session: true } }), 'encrypted', /password-protected/);
    const streams = pptStreams(SAMPLE_PPT);
    assertPptError(() => parsePresentation(fileOf(streams, { EncryptedSummary: new Uint8Array(100) })), 'encrypted', /password-protected/);
    assertPptError(() => parsePresentation(new Uint8Array(pptFile({ EncryptedPackage: new Uint8Array(5000), EncryptionInfo: new Uint8Array(200) }))), 'encrypted', /password-protected/);
  });

  it('explains files saved by PowerPoint 95 or older', () => {
    assertPptError(() => parsePresentation(new Uint8Array(pptFile({ PP40: new Uint8Array(6000), 'Current User': new Uint8Array(12) }))), 'too-old', /PowerPoint 95 or older/);
    const shortUser = Buffer.concat([u32(10), Buffer.alloc(10)]);
    const streams = pptStreams(SAMPLE_PPT);
    assertPptError(() => parsePresentation(fileOf({ ...streams, currentUser: shortUser })), 'too-old', /PowerPoint 95 or older/);
  });

  it('names the format of other compound files', () => {
    assertPptError(() => parsePresentation(new Uint8Array(pptFile({ WordDocument: new Uint8Array(5000) }))), 'not-ppt', /actually a \.doc file/);
    assertPptError(() => parsePresentation(new Uint8Array(pptFile({ Workbook: new Uint8Array(5000) }))), 'not-ppt', /actually a \.xls file/);
    assertPptError(() => parsePresentation(new Uint8Array(pptFile({ Other: new Uint8Array(10) }))), 'not-ppt', /no PowerPoint Document stream/);
    assertPptError(() => parsePresentation(new TextEncoder().encode('PK\u0003\u0004 not a compound file')), 'not-ppt', /no compound file header/);
  });

  it('reports a document record whose length runs past the stream', () => {
    const streams = pptStreams(SAMPLE_PPT);
    streams.document.writeUInt32LE(0x7fffffff, (streams.offsets.get(1) ?? 0) + 4);
    assertPptError(() => parsePresentation(fileOf(streams)), 'damaged', /document record could not be found/);
  });

  it('reports a slide list record whose length runs past the document record', () => {
    const streams = pptStreams(SAMPLE_PPT);
    const document = streams.document;
    // The slide list is the last SlideListWithText (instance 0) inside the DocumentContainer.
    const listOffset = document.lastIndexOf(Buffer.from([0x0f, 0x00, 0xf0, 0x0f]), streams.offsets.get(2));
    document.writeUInt32LE(0x00ffffff, listOffset + 4);
    assertPptError(() => parsePresentation(fileOf(streams)), 'damaged', /runs past the end of its container/);
  });

  it('reports a compound file whose sector chain loops', () => {
    const streams = pptStreams(SAMPLE_PPT);
    const bytes = pptFile(
      { 'PowerPoint Document': Buffer.concat([streams.document, Buffer.alloc(5000)]), 'Current User': streams.currentUser, Pictures: streams.pictures },
      {
        patchFat: (fat, layout) => {
          const start = layout.streams.get('PowerPoint Document')?.start ?? 0;
          fat[start + 4] = start;
        },
      },
    );
    assertPptError(() => parsePresentation(new Uint8Array(bytes)), 'damaged', /This \.ppt file is damaged: compound file sector chain loops/);
  });
});

/** Deterministic pseudo-random integers below `limit` (mulberry32, whose low bits are as random as the high ones). */
function randomSource(seed: number): (limit: number) => number {
  let state = seed >>> 0;
  return (limit) => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return (((t ^ (t >>> 14)) >>> 0) / 2 ** 32) * limit >>> 0;
  };
}

/** Overwrites 1–8 random bytes of `bytes` in [from, to). */
function corrupt(bytes: Buffer, random: (limit: number) => number, from = 0, to = bytes.length): Buffer {
  const copy = Buffer.from(bytes);
  for (let flips = 1 + random(8); flips > 0; flips--) {
    copy[from + random(to - from)] = random(256);
  }
  return copy;
}

/** Parses each corrupted file: it either opens or fails with a PptError, and no file takes long. */
function fuzz(rounds: number, build: (random: (limit: number) => number, round: number) => Uint8Array): void {
  const random = randomSource(rounds);
  for (let round = 0; round < rounds; round++) {
    const bytes = build(random, round);
    const started = performance.now();
    try {
      parsePresentation(bytes);
    } catch (error) {
      assert.ok(error instanceof PptError, `round ${round}: ${(error as Error).stack}`);
    }
    assert.ok(performance.now() - started < 2000, `round ${round} took ${Math.round(performance.now() - started)} ms`);
  }
}

describe('parsePresentation on corrupted bytes', () => {
  const streams = pptStreams(SAMPLE_PPT);

  it('copes with a corrupted or truncated document stream', () => {
    fuzz(400, (random, round) => {
      const document = corrupt(streams.document, random);
      return fileOf({ ...streams, document: round % 5 === 0 ? document.subarray(0, random(document.length)) : document });
    });
  });

  it('copes with a corrupted compound file header', () => {
    const file = Buffer.from(fileOf(streams));
    // Mostly the fields after the signature: version, byte order, sector sizes, counts, first sectors, DIFAT.
    fuzz(400, (random) => new Uint8Array(corrupt(file, random, 24, 80)));
    fuzz(100, (random) => new Uint8Array(corrupt(file, random, 0, 512)));
  });

  it('copes with a corrupted Pictures stream', () => {
    fuzz(200, (random) => fileOf({ ...streams, pictures: corrupt(streams.pictures, random, 0, 64) }));
    fuzz(100, (random) => fileOf({ ...streams, pictures: corrupt(streams.pictures, random) }));
  });

  it('copes with a corrupted Current User stream', () => {
    // Mostly the record header, token and edit offset; the rest is the user name.
    fuzz(300, (random) => fileOf({ ...streams, currentUser: corrupt(streams.currentUser, random, 0, 28) }));
    fuzz(100, (random) => fileOf({ ...streams, currentUser: corrupt(streams.currentUser, random) }));
  });
});

/** Fails if `run` takes longer than `ms`; the hostile files below took seconds to minutes before their caps. */
function within<T>(ms: number, run: () => T): T {
  const started = performance.now();
  const result = run();
  const elapsed = performance.now() - started;
  assert.ok(elapsed < ms, `took ${Math.round(elapsed)} ms (limit ${ms} ms)`);
  return result;
}

/** Re-points every picture store entry of `format` at offset 0 of the Pictures stream. */
function pointStoreEntriesAtFirstBlip(document: Buffer, storeType: number): number {
  const header = Buffer.from([(storeType << 4) | 2, 0x00, 0x07, 0xf0]);
  let patched = 0;
  for (let at = document.indexOf(header); at >= 0; at = document.indexOf(header, at + 1)) {
    document.writeUInt32LE(0, at + 8 + 28);
    patched += 1;
  }
  return patched;
}

/** One big blip (`body` after a UID and tag) as the whole Pictures stream, shown by `count` store entries. */
function sharedBlipFile(format: 'jpeg' | 'dib', count: number, body: Buffer): Uint8Array {
  const placeholder = format === 'jpeg' ? Buffer.from([0xff, 0xd8, 0xff]) : Buffer.alloc(60);
  const streams = pptStreams({
    pictures: Array.from({ length: count }, () => ({ format, data: placeholder })),
    slides: [{ pictures: Array.from({ length: count }, (_, i) => i + 1) }],
  });
  assert.equal(pointStoreEntriesAtFirstBlip(streams.document, format === 'jpeg' ? 5 : 7), count);
  const blip = atom(format === 'jpeg' ? 0xf01d : 0xf01f, Buffer.concat([Buffer.alloc(16), Buffer.from([0xff]), body]), format === 'jpeg' ? 0x46a : 0x7a8);
  return new Uint8Array(pptFile({ 'PowerPoint Document': streams.document, 'Current User': streams.currentUser, Pictures: blip }));
}

describe('parsePresentation against files built to waste time', () => {
  it('counts every persist directory entry it reads, also when many saves share one directory', () => {
    // 10,000 saves all pointing at one directory of 64 × 4,095 entries: 2.6 billion entries to read.
    const parts: Buffer[] = [atom(0x0001, Buffer.alloc(0))];
    const offsetOf = () => parts.reduce((sum, part) => sum + part.length, 0);
    const documentOffset = offsetOf();
    parts.push(container(0x03e8, [container(0x0ff0, [], 0)]));
    const directoryOffset = offsetOf();
    const group = Buffer.concat([u32(((4095 << 20) | 2) >>> 0), Buffer.alloc(4095 * 4)]);
    parts.push(atom(0x1772, Buffer.concat([u32((1 << 20) | 1, documentOffset), ...Array.from({ length: 64 }, () => group)])));
    let previous = 0;
    for (let i = 0; i < 10_000; i++) {
      const offset = offsetOf();
      parts.push(atom(0x0ff5, Buffer.concat([u32(256), Buffer.from([0, 0, 0, 3]), u32(previous, directoryOffset, 1, 5000), Buffer.from([1, 0, 0, 0])])));
      previous = offset;
    }
    const bytes = new Uint8Array(pptFile({ 'PowerPoint Document': Buffer.concat(parts), 'Current User': currentUserStream(previous) }));
    within(1000, () => assertPptError(() => parsePresentation(bytes), 'too-complex', /too many objects/));
  });

  it('counts persist entries that a later save repeats', () => {
    const options = { update: { spec: SAMPLE_PPT, persistIds: [1, slidePersistId(0)] } };
    const first = pptStreams(SAMPLE_PPT, options);
    const firstSave = first.document.readUInt32LE((first.edits[0] ?? 0) + 8 + 12);
    const entries = first.document.readUInt32LE(firstSave + 8) >>> 20;
    // The second save repeats two ids: they are read again.
    assert.equal(parse(SAMPLE_PPT, options, { ...PPT_LIMITS, maxPersistEntries: entries + 2 }).slides.length, 3);
    assertPptError(() => parse(SAMPLE_PPT, options, { ...PPT_LIMITS, maxPersistEntries: entries + 1 }), 'too-complex', /too many objects/);
  });

  it('pairs paragraphs with their style runs in one pass', () => {
    // 200,000 one-letter paragraphs, one style run each.
    const presentation = within(2000, () => parse({ slides: [{ textBoxes: [{ type: 'other', text: 'a\r'.repeat(200_000), bytes: true }] }] }));
    assert.equal(presentation.slides[0]?.texts[0]?.paragraphs.length, PPT_LIMITS.maxParagraphs);
    assert.deepEqual(presentation.warnings, ['Only the first 100,000 paragraphs are shown.']);
  });

  it('caps the paragraphs shown across the presentation', () => {
    const presentation = parse(SAMPLE_PPT, {}, { ...PPT_LIMITS, maxParagraphs: 3 });
    assert.deepEqual(texts(presentation.slides[0]), ['Quarterly review', '四半期レビュー / Document Lens sample']);
    assert.deepEqual(texts(presentation.slides[1]), []);
    assert.deepEqual(presentation.warnings, ['Only the first 3 paragraphs are shown.']);
  });

  it('examines a picture once however many store entries point at it', () => {
    // 500 entries showing one 16 MB "JPEG" of fill bytes, whose frame header would be searched for each time.
    const jpeg = sharedBlipFile('jpeg', 500, Buffer.concat([Buffer.from([0xff, 0xd8, 0xff]), Buffer.alloc(16 * 1024 * 1024, 0xff)]));
    const pictures = within(2000, () => parsePresentation(jpeg)).slides[0]?.pictures ?? [];
    assert.equal(pictures.length, 500);
    assert.equal(new Set(pictures).size, 1);
    assert.equal(pictures[0]?.note, 'Damaged picture (not shown)');

    // The same with a 1 × 1 DIB followed by 16 MB, which would be copied each time.
    const header = Buffer.concat([u32(40, 1, 1), Buffer.from([1, 0, 24, 0]), u32(0, 4, 0, 0, 0, 0)]);
    const dib = sharedBlipFile('dib', 500, Buffer.concat([header, Buffer.alloc(16 * 1024 * 1024)]));
    const shown = within(2000, () => parsePresentation(dib)).slides[0]?.pictures ?? [];
    assert.equal(new Set(shown).size, 1);
    assert.equal(shown[0]?.mime, 'image/bmp');
  });

  it('charges picture data to the budget before scanning or converting it, damaged pictures included', () => {
    const fake = (fill: number) => Buffer.concat([Buffer.from([0xff, 0xd8, 0xff]), Buffer.alloc(100_000, fill)]);
    const spec: PptSpec = {
      pictures: [
        { format: 'jpeg', data: fake(0xff) },
        { format: 'jpeg', data: fake(0xfe) },
        { format: 'png', data: rgbPng(2, 2, () => [0, 0, 0]) },
      ],
      slides: [{ pictures: [1, 2, 3] }],
    };
    const presentation = parse(spec, {}, { ...PPT_LIMITS, maxTotalPictureSize: 150_000 });
    assert.deepEqual(
      presentation.slides[0]?.pictures.map((picture) => picture.note),
      ['Damaged picture (not shown)', 'Picture left out (too much picture data)', undefined],
    );
    assert.deepEqual(presentation.warnings, ['Some pictures are left out: the presentation holds too much picture data.']);
  });

  it('counts shape property tables as work, also when many slides point at the same slide', () => {
    // 200 slide-list entries all pointing at one slide with 100 shapes of 4,095 properties each.
    const properties = Buffer.alloc(4095 * 6);
    for (let i = 0; i < 4095; i++) {
      properties.writeUInt16LE(0x0181, i * 6);
    }
    const shape = container(0xf004, [atom(0xf00b, properties, 4095, 3)]);
    const streams = pptStreams({ slides: [{ extraShapes: Array.from({ length: 100 }, () => shape) }, ...Array.from({ length: 199 }, () => ({}))] });
    for (let i = 1; i < 200; i++) {
      streams.document.writeUInt32LE(streams.offsets.get(slidePersistId(0)) ?? 0, persistEntryOffset(streams, slidePersistId(i)));
    }
    within(1000, () => assertPptError(() => parsePresentation(fileOf(streams)), 'too-complex', /too many records/));
  });
});
