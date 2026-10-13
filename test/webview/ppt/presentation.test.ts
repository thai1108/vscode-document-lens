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

  it('never fails with anything but a PptError on corrupted bytes', () => {
    const streams = pptStreams(SAMPLE_PPT);
    let seed = 1;
    const random = (limit: number) => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return seed % limit;
    };
    for (let round = 0; round < 400; round++) {
      const document = Buffer.from(streams.document);
      for (let flips = 1 + random(8); flips > 0; flips--) {
        document[random(document.length)] = random(256);
      }
      const truncated = round % 5 === 0 ? document.subarray(0, random(document.length)) : document;
      try {
        parsePresentation(fileOf({ ...streams, document: Buffer.from(truncated) }));
      } catch (error) {
        assert.ok(error instanceof PptError, `round ${round}: ${(error as Error).stack}`);
      }
    }
  });
});
