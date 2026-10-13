// Reads a PowerPoint 97–2003 presentation ([MS-PPT]) into what the viewer shows: each slide's text and pictures.
//
// The "PowerPoint Document" stream is a log of saves. The Current User stream points at the last UserEditAtom; each
// UserEditAtom points at its PersistDirectoryAtom (persist id → stream offset) and at the previous save's
// UserEditAtom, and newer saves win. The DocumentContainer found that way lists the slides in order
// (SlideListWithText); each slide's own container holds its drawing, whose shapes carry text (directly, or as a
// reference into the slide list) and picture references into the document's picture store. When the chain is broken,
// the stream's top-level records are scanned instead.
import { CompoundFile, CompoundFileError } from './cfb.ts';
import { damaged, encrypted, PptError, tooOld } from './errors.ts';
import { PPT_LIMITS, type PptLimits } from './limits.ts';
import { isBlip, placeholder, readBlip, type Picture } from './pictures.ts';
import { RecordStream, RT, type PptRecord } from './records.ts';
import { decodeBytes, decodeChars, paragraphRuns, textKind, toParagraphs, type TextBlock } from './text.ts';

export interface Slide {
  /** 1-based position in the slide show. */
  readonly number: number;
  readonly hidden: boolean;
  /** Title blocks first, then the rest from top to bottom. */
  readonly texts: readonly TextBlock[];
  readonly pictures: readonly Picture[];
  /** Part of the slide could not be read; what could be read is still listed. */
  readonly damaged: boolean;
}

export interface Presentation {
  /** Slide size in master units (576 per inch). */
  readonly width: number;
  readonly height: number;
  readonly slides: readonly Slide[];
  /** What was left out (limits reached, damaged parts), for the notice banner. */
  readonly warnings: readonly string[];
}

const DOCUMENT_STREAM = 'PowerPoint Document';
const CURRENT_USER_STREAM = 'Current User';
const PICTURES_STREAM = 'Pictures';
const HEADER_TOKEN_ENCRYPTED = 0xf3d1c4df;
const CURRENT_USER_MIN_SIZE = 28;
const USER_EDIT_MIN_SIZE = 28;
const USER_EDIT_ENCRYPTED_SIZE = 32;
const FBSE_FIXED_SIZE = 36;
/** 10 × 7.5 inches. */
const DEFAULT_SLIDE_SIZE = { width: 5760, height: 4320 };
/** SlideListWithText instances: 0 lists the slides (1 the masters, 2 the notes). */
const SLIDE_LIST = 0;
const SHAPE_DELETED = 0x8;
const SHAPE_BACKGROUND = 0x400;
const PROPERTY_ID = 0x3fff;
const PROPERTY_COMPLEX = 0x8000;
const PROPERTY_PIB = 0x0104;
/** Slide placeholders that repeat master content: date, slide number, footer, header. */
const MASTER_PLACEHOLDERS = new Set([0x07, 0x08, 0x09, 0x0a]);
const SLIDE_HIDDEN = 0x4;

export function parsePresentation(bytes: Uint8Array, limits: PptLimits = PPT_LIMITS): Presentation {
  try {
    return readPresentation(bytes, limits);
  } catch (error) {
    if (error instanceof CompoundFileError) {
      throw damaged(error.message);
    }
    if (error instanceof RangeError) {
      throw damaged('a record is truncated');
    }
    throw error;
  }
}

function readPresentation(bytes: Uint8Array, limits: PptLimits): Presentation {
  if (!CompoundFile.isCompoundFile(bytes)) {
    throw new PptError('not-ppt', 'Not a .ppt file (no compound file header)');
  }
  const file = CompoundFile.open(bytes, limits);
  const documentBytes = file.stream(DOCUMENT_STREAM);
  if (!documentBytes) {
    throw missingDocument(file);
  }
  if (file.has('EncryptedSummary')) {
    throw encrypted();
  }
  const budget = { records: limits.maxRecords };
  const document = new RecordStream(documentBytes, budget);
  const { persist, documentOffset } = locateDocument(document, readCurrentUser(file, limits), limits);
  const container = document.at(documentOffset);
  const children = document.children(container);
  const warnings = new Set<string>();
  const picturesBytes = readPicturesStream(file, warnings);
  const context: Context = {
    limits,
    document,
    pictures: picturesBytes && new RecordStream(picturesBytes, budget),
    blipStore: [],
    pictureCache: new Map(),
    warnings,
    charsLeft: limits.maxTextChars,
    picturesLeft: limits.maxPictures,
    pictureBytesLeft: limits.maxTotalPictureSize,
  };
  context.blipStore = readBlipStore(context, children.find((child) => child.type === RT.DrawingGroup));

  const slideList = children.find((child) => child.type === RT.SlideListWithText && child.instance === SLIDE_LIST);
  let entries = slideList ? readSlideList(context, slideList) : [];
  if (entries.length > limits.maxSlides) {
    context.warnings.add(`Only the first ${count(limits.maxSlides)} slides are shown.`);
    entries = entries.slice(0, limits.maxSlides);
  }
  const slides = entries.map((entry, i) => new SlideReader(context, entry, i + 1).read(persist.get(entry.persistRef)));
  return { ...slideSize(document, children), slides, warnings: [...context.warnings] };
}

/** The Pictures stream, if any; when it is damaged the slides are still shown, without pictures. */
function readPicturesStream(file: CompoundFile, warnings: Set<string>): Uint8Array | undefined {
  try {
    return file.stream(PICTURES_STREAM);
  } catch (error) {
    if (!(error instanceof CompoundFileError)) {
      throw error;
    }
    warnings.add('The pictures are damaged and left out.');
    return undefined;
  }
}

/** What a compound file without a PowerPoint Document stream is instead. */
function missingDocument(file: CompoundFile): PptError {
  if (file.has('PP40')) {
    return tooOld();
  }
  if (file.has('EncryptedPackage')) {
    return encrypted();
  }
  if (file.has('WordDocument')) {
    return new PptError('not-ppt', 'This .ppt file is actually a .doc file; rename it to open it');
  }
  if (file.has('Workbook') || file.has('Book')) {
    return new PptError('not-ppt', 'This .ppt file is actually a .xls file; rename it to open it');
  }
  return new PptError('not-ppt', 'Not a .ppt file (no PowerPoint Document stream)');
}

/** CurrentUserAtom.offsetToCurrentEdit, or undefined when the stream is missing or unusable. */
function readCurrentUser(file: CompoundFile, limits: PptLimits): number | undefined {
  let data: Uint8Array | undefined;
  try {
    data = file.stream(CURRENT_USER_STREAM, limits.maxCurrentUserSize);
  } catch (error) {
    if (error instanceof CompoundFileError) {
      return undefined;
    }
    throw error;
  }
  if (!data) {
    return undefined;
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (data.byteLength < CURRENT_USER_MIN_SIZE) {
    // PowerPoint 95 wrote a short Current User stream: a 4-byte size, then that many bytes.
    if (data.byteLength >= 4 && view.getUint32(0, true) + 4 === data.byteLength) {
      throw tooOld();
    }
    return undefined;
  }
  if (view.getUint16(2, true) !== RT.CurrentUserAtom) {
    return undefined;
  }
  if (view.getUint32(12, true) === HEADER_TOKEN_ENCRYPTED) {
    throw encrypted();
  }
  return view.getUint32(16, true);
}

interface DocumentLocation {
  persist: Map<number, number>;
  documentOffset: number;
}

function locateDocument(document: RecordStream, currentEdit: number | undefined, limits: PptLimits): DocumentLocation {
  if (currentEdit !== undefined) {
    try {
      const chain = readEditChain(document, currentEdit, limits);
      const session = chain.encryptRef === undefined ? undefined : chain.persist.get(chain.encryptRef);
      if (session !== undefined && isRecordAt(document, session, RT.CryptSession10Container)) {
        throw encrypted();
      }
      const documentOffset = chain.persist.get(chain.documentRef);
      if (documentOffset !== undefined && isRecordAt(document, documentOffset, RT.Document)) {
        return { persist: chain.persist, documentOffset };
      }
    } catch (error) {
      if (!isDamage(error)) {
        throw error;
      }
    }
  }
  const scan = scanTopLevel(document, limits);
  if (scan.encrypted) {
    throw encrypted();
  }
  if (scan.documentOffset === undefined) {
    throw damaged('its document record could not be found');
  }
  return { persist: scan.persist, documentOffset: scan.documentOffset };
}

/** Follows the UserEditAtoms from the newest save back, merging their persist directories (newer entries win). */
function readEditChain(document: RecordStream, firstEdit: number, limits: PptLimits): { persist: Map<number, number>; documentRef: number; encryptRef?: number } {
  const persist = new Map<number, number>();
  const seen = new Set<number>();
  let newest: { documentRef: number; encryptRef?: number } | undefined;
  for (let edit = firstEdit; ; ) {
    if (seen.has(edit)) {
      throw damaged('its edit history loops');
    }
    if (seen.size >= limits.maxEdits) {
      throw new PptError('too-complex', 'This .ppt file has too many saved edits to preview');
    }
    seen.add(edit);
    const atom = document.at(edit);
    if (atom.type !== RT.UserEditAtom) {
      throw damaged('its edit history points at the wrong record');
    }
    document.expect(atom, USER_EDIT_MIN_SIZE, 'an edit record');
    newest ??= {
      documentRef: document.u32(atom.start + 16),
      encryptRef: atom.end - atom.start >= USER_EDIT_ENCRYPTED_SIZE ? document.u32(atom.start + 28) : undefined,
    };
    readPersistDirectory(document, document.u32(atom.start + 12), persist, false, limits);
    const previous = document.u32(atom.start + 8);
    if (previous === 0) {
      return { persist, ...newest };
    }
    edit = previous;
  }
}

function readPersistDirectory(document: RecordStream, offset: number, persist: Map<number, number>, overwrite: boolean, limits: PptLimits): void {
  const directory = document.at(offset);
  if (directory.type !== RT.PersistDirectoryAtom) {
    throw damaged('a persist directory is missing');
  }
  for (let position = directory.start; position + 4 <= directory.end; ) {
    const entry = document.u32(position);
    const firstId = entry & 0xfffff;
    const count = entry >>> 20;
    position += 4;
    if (position + count * 4 > directory.end) {
      throw damaged('a persist directory is truncated');
    }
    for (let i = 0; i < count; i++) {
      if (overwrite || !persist.has(firstId + i)) {
        persist.set(firstId + i, document.u32(position + i * 4));
      }
    }
    if (persist.size > limits.maxPersistEntries) {
      throw new PptError('too-complex', 'This .ppt file has too many objects to preview');
    }
    position += count * 4;
  }
}

/** Fallback for a broken edit chain: the last DocumentContainer in the stream and every persist directory in order. */
function scanTopLevel(document: RecordStream, limits: PptLimits): { persist: Map<number, number>; documentOffset?: number; encrypted: boolean } {
  const persist = new Map<number, number>();
  let documentOffset: number | undefined;
  let isEncrypted = false;
  for (let offset = 0; offset + 8 <= document.bytes.byteLength; ) {
    let record: PptRecord;
    try {
      record = document.at(offset);
    } catch (error) {
      if (isDamage(error)) {
        break;
      }
      throw error;
    }
    if (record.type === RT.Document) {
      documentOffset = offset;
    } else if (record.type === RT.CryptSession10Container) {
      isEncrypted = true;
    } else if (record.type === RT.PersistDirectoryAtom) {
      try {
        readPersistDirectory(document, offset, persist, true, limits);
      } catch (error) {
        if (!isDamage(error)) {
          throw error;
        }
      }
    }
    offset = record.end;
  }
  return { persist, documentOffset, encrypted: isEncrypted };
}

function count(value: number): string {
  return value.toLocaleString('en-US');
}

function isRecordAt(stream: RecordStream, offset: number, type: number): boolean {
  try {
    return stream.at(offset).type === type;
  } catch (error) {
    if (isDamage(error)) {
      return false;
    }
    throw error;
  }
}

function isDamage(error: unknown): boolean {
  return (error instanceof PptError && error.kind === 'damaged') || error instanceof RangeError || error instanceof CompoundFileError;
}

function slideSize(document: RecordStream, children: readonly PptRecord[]): { width: number; height: number } {
  const atom = children.find((child) => child.type === RT.DocumentAtom);
  if (!atom || atom.end - atom.start < 8) {
    return DEFAULT_SLIDE_SIZE;
  }
  const width = document.i32(atom.start);
  const height = document.i32(atom.start + 4);
  const ratio = width / height;
  return width > 0 && height > 0 && ratio > 0.1 && ratio < 10 ? { width, height } : DEFAULT_SLIDE_SIZE;
}

interface Context {
  readonly limits: PptLimits;
  readonly document: RecordStream;
  readonly pictures: RecordStream | undefined;
  /** OfficeArtBStoreContainer entries; a shape's pib property is a 1-based index into it. */
  blipStore: readonly PptRecord[];
  readonly pictureCache: Map<number, Picture>;
  readonly warnings: Set<string>;
  charsLeft: number;
  picturesLeft: number;
  pictureBytesLeft: number;
}

function readBlipStore(context: Context, drawingGroup: PptRecord | undefined): PptRecord[] {
  const { document } = context;
  try {
    const dgg = drawingGroup && document.child(drawingGroup, RT.DggContainer);
    const store = dgg && document.child(dgg, RT.BStoreContainer);
    return store ? document.children(store) : [];
  } catch (error) {
    if (!isDamage(error)) {
      throw error;
    }
    context.warnings.add('The list of pictures is damaged; pictures are left out.');
    return [];
  }
}

interface ListText {
  readonly kind: 'text';
  readonly type: number;
  text: string;
  style?: PptRecord;
}

type TextItem = ListText | { readonly kind: 'ref'; readonly index: number };

interface SlideEntry {
  readonly persistRef: number;
  /** Placeholder text kept in the slide list; shapes refer to it by index (OutlineTextRefAtom). */
  readonly texts: readonly ListText[];
}

/** Each SlidePersistAtom is followed by the text of that slide's placeholders. */
function readSlideList(context: Context, list: PptRecord): SlideEntry[] {
  const { document } = context;
  const groups: { atom: PptRecord; records: PptRecord[] }[] = [];
  for (const child of document.children(list)) {
    if (child.type === RT.SlidePersistAtom) {
      groups.push({ atom: child, records: [] });
    } else {
      groups.at(-1)?.records.push(child);
    }
  }
  return groups.slice(0, context.limits.maxSlides + 1).map(({ atom, records }) => ({
    persistRef: atom.end - atom.start >= 4 ? document.u32(atom.start) : 0,
    texts: textItems(context, records).filter((item) => item.kind === 'text'),
  }));
}

/** Text records in order: a TextHeaderAtom starts a text, the chars/bytes atom and style atom after it belong to it. */
function textItems(context: Context, records: readonly PptRecord[]): TextItem[] {
  const { document } = context;
  const items: TextItem[] = [];
  let current: ListText | undefined;
  for (const record of records) {
    switch (record.type) {
      case RT.TextHeaderAtom:
        current = { kind: 'text', type: record.end - record.start >= 4 ? document.u32(record.start) : -1, text: '' };
        items.push(current);
        break;
      case RT.TextCharsAtom:
      case RT.TextBytesAtom:
        if (!current) {
          current = { kind: 'text', type: -1, text: '' };
          items.push(current);
        }
        current.text = decodeText(context, record);
        break;
      case RT.StyleTextPropAtom:
        if (current) {
          current.style = record;
        }
        break;
      case RT.OutlineTextRefAtom:
        if (record.end - record.start >= 4) {
          items.push({ kind: 'ref', index: document.i32(record.start) });
        }
        current = undefined;
        break;
    }
  }
  return items;
}

function decodeText(context: Context, record: PptRecord): string {
  const wide = record.type === RT.TextCharsAtom;
  const bytes = context.document.data(record);
  const length = wide ? bytes.byteLength >> 1 : bytes.byteLength;
  if (length > context.charsLeft) {
    context.warnings.add(`Text after the first ${count(context.limits.maxTextChars)} characters is left out.`);
  }
  const kept = Math.min(length, context.charsLeft);
  context.charsLeft -= kept;
  return wide ? decodeChars(bytes.subarray(0, kept * 2)) : decodeBytes(bytes.subarray(0, kept));
}

function toBlock(context: Context, text: ListText): TextBlock {
  const runs = text.style ? paragraphRuns(context.document.data(text.style), text.text.length) : [];
  return { kind: textKind(text.type), paragraphs: toParagraphs(text.text, text.type, runs) };
}

interface Anchor {
  readonly top: number;
  readonly left: number;
}

interface Placed<T> {
  readonly item: T;
  readonly anchor: Anchor | undefined;
  readonly order: number;
}

/** Top to bottom, then left to right; things without a position keep their order after the rest. */
function inReadingOrder<T>(placed: readonly Placed<T>[]): T[] {
  return [...placed]
    .sort((a, b) => {
      if (!a.anchor || !b.anchor) {
        return (a.anchor ? 0 : 1) - (b.anchor ? 0 : 1) || a.order - b.order;
      }
      return a.anchor.top - b.anchor.top || a.anchor.left - b.anchor.left || a.order - b.order;
    })
    .map((entry) => entry.item);
}

class SlideReader {
  private readonly context: Context;
  private readonly entry: SlideEntry;
  private readonly number: number;
  private readonly texts: Placed<ListText>[] = [];
  private readonly pictures: Placed<Picture>[] = [];
  private readonly usedListTexts = new Set<number>();
  private order = 0;

  constructor(context: Context, entry: SlideEntry, number: number) {
    this.context = context;
    this.entry = entry;
    this.number = number;
  }

  read(offset: number | undefined): Slide {
    const { document } = this.context;
    let hidden = false;
    let isDamaged = false;
    try {
      if (offset === undefined) {
        throw damaged('a slide is missing');
      }
      const slide = document.at(offset);
      if (slide.type !== RT.Slide) {
        throw damaged('a slide reference points at the wrong record');
      }
      for (const child of document.children(slide)) {
        if (child.type === RT.SlideShowSlideInfoAtom && child.end - child.start >= 12) {
          hidden = (document.u16(child.start + 10) & SLIDE_HIDDEN) !== 0;
        } else if (child.type === RT.Drawing) {
          this.readDrawing(child);
        }
      }
    } catch (error) {
      if (!isDamage(error)) {
        throw error;
      }
      isDamaged = true;
    }
    this.addUnplacedListTexts();
    const blocks = inReadingOrder(this.texts)
      .map((text) => toBlock(this.context, text))
      .filter((block) => block.paragraphs.length > 0);
    return {
      number: this.number,
      hidden,
      texts: [...blocks.filter((block) => block.kind === 'title'), ...blocks.filter((block) => block.kind !== 'title')],
      pictures: inReadingOrder(this.pictures),
      damaged: isDamaged,
    };
  }

  private readDrawing(drawing: PptRecord): void {
    const { document } = this.context;
    const dg = document.child(drawing, RT.DgContainer);
    // The DgContainer's own SpContainer is the slide background; only the shape tree (SpgrContainer) is read.
    for (const group of dg ? document.children(dg) : []) {
      if (group.type === RT.SpgrContainer) {
        this.readGroup(group, 1, undefined);
      }
    }
  }

  /** A group's first SpContainer describes the group itself; its anchor positions the group's members. */
  private readGroup(group: PptRecord, depth: number, outerAnchor: Anchor | undefined): void {
    if (depth > this.context.limits.maxDepth) {
      this.context.warnings.add('Some deeply nested shapes are left out.');
      return;
    }
    let anchor = outerAnchor;
    this.context.document.children(group).forEach((child, i) => {
      if (child.type === RT.SpContainer) {
        const own = this.readShape(child, anchor);
        if (i === 0) {
          anchor = own ?? outerAnchor;
        }
      } else if (child.type === RT.SpgrContainer) {
        this.readGroup(child, depth + 1, anchor);
      }
    });
  }

  /** Collects the shape's text and picture; returns its own anchor (top-level shapes have one). */
  private readShape(shape: PptRecord, groupAnchor: Anchor | undefined): Anchor | undefined {
    const { document } = this.context;
    let flags = 0;
    let anchor: Anchor | undefined;
    let pib = 0;
    let placeholderType = -1;
    let textbox: PptRecord | undefined;
    for (const record of document.children(shape)) {
      switch (record.type) {
        case RT.FSP:
          flags = record.end - record.start >= 8 ? document.u32(record.start + 4) : 0;
          break;
        case RT.ClientAnchor:
          anchor = readAnchor(document, record);
          break;
        case RT.FOPT:
        case RT.TertiaryFOPT:
          pib ||= readPib(document, record);
          break;
        case RT.ClientData: {
          const atom = document.child(record, RT.PlaceholderAtom);
          if (atom && atom.end - atom.start >= 8) {
            placeholderType = document.u8(atom.start + 4);
          }
          break;
        }
        case RT.ClientTextbox:
          textbox = record;
          break;
      }
    }
    if (flags & (SHAPE_DELETED | SHAPE_BACKGROUND) || MASTER_PLACEHOLDERS.has(placeholderType)) {
      return anchor;
    }
    const position = anchor ?? groupAnchor;
    const order = this.order++;
    for (const item of textbox ? textItems(this.context, document.children(textbox)) : []) {
      const text = item.kind === 'text' ? item : this.usedListTexts.has(item.index) ? undefined : this.entry.texts[item.index];
      if (text) {
        if (item.kind === 'ref') {
          this.usedListTexts.add(item.index);
        }
        this.texts.push({ item: text, anchor: position, order });
      }
    }
    if (pib > 0) {
      this.addPicture(pib, position, order);
    }
    return anchor;
  }

  private addPicture(pib: number, anchor: Anchor | undefined, order: number): void {
    const { context } = this;
    if (context.picturesLeft <= 0) {
      context.warnings.add(`Only the first ${count(context.limits.maxPictures)} pictures are shown.`);
      return;
    }
    context.picturesLeft -= 1;
    this.pictures.push({ item: pictureAt(context, pib), anchor, order });
  }

  /** Slide-list text no shape refers to (e.g. the slide's drawing is damaged), unless a shape shows the same text. */
  private addUnplacedListTexts(): void {
    const shown = new Set(this.texts.map((placed) => placed.item.text));
    this.entry.texts.forEach((text, index) => {
      if (!this.usedListTexts.has(index) && !shown.has(text.text)) {
        this.texts.push({ item: text, anchor: undefined, order: this.order++ });
      }
    });
  }
}

/** OfficeArtClientAnchor: a SmallRectStruct (16-bit) or RectStruct (32-bit), both top, left, right, bottom. */
function readAnchor(document: RecordStream, record: PptRecord): Anchor | undefined {
  const length = record.end - record.start;
  if (length >= 16) {
    return { top: document.i32(record.start), left: document.i32(record.start + 4) };
  }
  return length >= 8 ? { top: document.i16(record.start), left: document.i16(record.start + 2) } : undefined;
}

/** The pib property (blip store index of the shape's picture) from an OfficeArtFOPT's property table. */
function readPib(document: RecordStream, options: PptRecord): number {
  for (let i = 0; i < options.instance; i++) {
    const position = options.start + i * 6;
    if (position + 6 > options.end) {
      break;
    }
    const id = document.u16(position);
    if ((id & PROPERTY_ID) === PROPERTY_PIB && !(id & PROPERTY_COMPLEX)) {
      return document.u32(position + 2);
    }
  }
  return 0;
}

function pictureAt(context: Context, index: number): Picture {
  const cached = context.pictureCache.get(index);
  if (cached) {
    return cached;
  }
  let picture: Picture;
  try {
    picture = loadPicture(context, index);
  } catch (error) {
    if (!isDamage(error)) {
      throw error;
    }
    picture = placeholder('unknown', 'Damaged picture (not shown)');
  }
  if (picture.data) {
    if (picture.data.byteLength > context.pictureBytesLeft) {
      context.warnings.add('Some pictures are left out: the presentation holds too much picture data.');
      picture = placeholder(picture.format, 'Picture left out (too much picture data)');
    } else {
      context.pictureBytesLeft -= picture.data.byteLength;
    }
  }
  context.pictureCache.set(index, picture);
  return picture;
}

/** A blip store entry is an OfficeArtFBSE pointing into the Pictures stream (or holding the blip), or a blip itself. */
function loadPicture(context: Context, index: number): Picture {
  const { document, pictures } = context;
  const entry = context.blipStore[index - 1];
  const missing = placeholder('unknown', 'Missing picture');
  if (!entry) {
    return missing;
  }
  if (isBlip(entry)) {
    return readBlip(document, entry, context.limits);
  }
  if (entry.type !== RT.FBSE) {
    return missing;
  }
  document.expect(entry, FBSE_FIXED_SIZE, 'a picture entry');
  const embedded = entry.start + FBSE_FIXED_SIZE + document.u8(entry.start + 33);
  if (embedded + 8 <= entry.end) {
    const blip = document.at(embedded, entry.end);
    return isBlip(blip) ? readBlip(document, blip, context.limits) : missing;
  }
  if (!pictures) {
    return missing;
  }
  const blip = pictures.at(document.u32(entry.start + 28));
  return isBlip(blip) ? readBlip(pictures, blip, context.limits) : missing;
}
