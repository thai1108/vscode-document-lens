/**
 * Writes PowerPoint 97–2003 presentations ([MS-PPT]) for tests: a compound file with the "PowerPoint Document",
 * "Current User" and "Pictures" streams, laid out like PowerPoint saves them. The document stream holds the
 * DocumentContainer (fonts and default text styles, picture store, slide list with placeholder text, view settings),
 * the main master, notes master and handout, and one SlideContainer per slide (placeholders, text boxes and picture
 * frames), then a PersistDirectoryAtom and the UserEditAtom that the Current User stream points at. An optional second
 * save appends changed objects the way an incremental save does. macOS Quick Look opens the result.
 */
import { createHash } from 'node:crypto';
import { writeCompoundFile, type CompoundFileOptions } from './cfbWriter.ts';

export type TextType = 'title' | 'body' | 'centerTitle' | 'subtitle' | 'other';

export interface TextSpec {
  type: TextType;
  /** Paragraphs separated by \r and line breaks inside a paragraph by \v, as PowerPoint stores them. */
  text: string;
  /** Store as a TextBytesAtom (8-bit) instead of a TextCharsAtom (UTF-16); only for text up to U+00FF. */
  bytes?: boolean;
  /** Indent level of each paragraph. */
  levels?: number[];
  /** Explicit bullet flag per paragraph; paragraphs without one follow the placeholder's default. */
  bullets?: (boolean | undefined)[];
  /** Vertical position of the shape in master units (default: below the previous shape). */
  top?: number;
  /**
   * For a placeholder: keep its text in the shape, as PowerPoint 2003 and later do, instead of in the slide list with
   * a reference from the shape (PowerPoint 97).
   */
  inShape?: boolean;
}

export interface SlideSpec {
  /** Placeholder text: in the slide list (SlideListWithText) and referenced from the shape, unless `inShape`. */
  placeholders?: TextSpec[];
  /** Text boxes whose text is stored in the shape itself. */
  textBoxes?: TextSpec[];
  /** Picture frames, each showing a 1-based index into PptSpec.pictures. */
  pictures?: number[];
  hidden?: boolean;
  /** Text of a footer placeholder, which repeats master content. */
  footer?: string;
  /** Records added to the slide's shape tree as they are (e.g. hostile nesting). */
  extraShapes?: Buffer[];
  /** Text in the slide list that no shape refers to. */
  listOnly?: TextSpec[];
}

export type PictureFormat = 'png' | 'jpeg' | 'dib' | 'emf' | 'wmf' | 'tiff';

export interface PictureSpec {
  format: PictureFormat;
  data: Uint8Array;
  /** Keep the picture inside its picture store entry instead of in the Pictures stream. */
  embedded?: boolean;
}

export interface PptSpec {
  slides: SlideSpec[];
  pictures?: PictureSpec[];
  /** Slide size in master units (576 per inch); default 10 × 7.5 inches. */
  slideSize?: { width: number; height: number };
  /** Marks the file as password-protected the way PowerPoint does (it does not encrypt anything). */
  encrypted?: { token?: boolean; session?: boolean };
}

export interface PptWriteOptions {
  /** A second save that rewrites the objects with these persist ids (1 = document, slidePersistId(i)) from another spec. */
  update?: { spec: PptSpec; persistIds: number[] };
}

export interface PptStreams {
  document: Buffer;
  currentUser: Buffer;
  pictures: Buffer;
  /** Offset of each persist object in the document stream (latest save). */
  offsets: Map<number, number>;
  /** Offsets of the UserEditAtoms, oldest first. */
  edits: number[];
}

/** Record types written here; the viewer has its own table. */
export const RECORD = {
  Document: 0x03e8,
  DocumentAtom: 0x03e9,
  EndDocumentAtom: 0x03ea,
  Slide: 0x03ee,
  SlideAtom: 0x03ef,
  Notes: 0x03f0,
  NotesAtom: 0x03f1,
  Handout: 0x0fc9,
  SlidePersistAtom: 0x03f3,
  Environment: 0x03f2,
  MainMaster: 0x03f8,
  SlideShowSlideInfoAtom: 0x03f9,
  DrawingGroup: 0x040b,
  Drawing: 0x040c,
  ColorSchemeAtom: 0x07f0,
  DocInfoList: 0x07d0,
  SlideViewInfo: 0x03fa,
  ZoomViewInfoAtom: 0x03fd,
  SlideViewInfoAtom: 0x03fe,
  FontCollection: 0x07d5,
  PlaceholderAtom: 0x0bc3,
  TextMasterStyleAtom: 0x0fa3,
  TextCFExceptionAtom: 0x0fa4,
  TextPFExceptionAtom: 0x0fa5,
  TextSIExceptionAtom: 0x0fa9,
  OutlineTextRefAtom: 0x0f9e,
  TextHeaderAtom: 0x0f9f,
  TextCharsAtom: 0x0fa0,
  StyleTextPropAtom: 0x0fa1,
  TextBytesAtom: 0x0fa8,
  FontEntityAtom: 0x0fb7,
  SlideListWithText: 0x0ff0,
  UserEditAtom: 0x0ff5,
  CurrentUserAtom: 0x0ff6,
  PersistDirectoryAtom: 0x1772,
  CryptSession10Container: 0x2f14,
  DggContainer: 0xf000,
  BStoreContainer: 0xf001,
  DgContainer: 0xf002,
  SpgrContainer: 0xf003,
  SpContainer: 0xf004,
  FDGGBlock: 0xf006,
  FBSE: 0xf007,
  FDG: 0xf008,
  FSPGR: 0xf009,
  FSP: 0xf00a,
  FOPT: 0xf00b,
  ClientTextbox: 0xf00d,
  ClientAnchor: 0xf010,
  ClientData: 0xf011,
} as const;

const TEXT_TYPES: Record<TextType, number> = { title: 0, body: 1, other: 4, subtitle: 5, centerTitle: 6 };
const PLACEHOLDER_IDS: Record<TextType, number> = { title: 0x0d, body: 0x0e, centerTitle: 0x0f, subtitle: 0x10, other: 0x13 };
const FOOTER_PLACEHOLDER = 0x09;
const SLIDE_LAYOUT_TITLE = 0x00;
const SLIDE_LAYOUT_TITLE_BODY = 0x01;
const SLIDE_LAYOUT_TITLE_ONLY = 0x07;
const MASTER_ID = 0x80000000;
const FIRST_SLIDE_ID = 256;
const DOCUMENT_PERSIST_ID = 1;
const MASTER_PERSIST_ID = 2;
const NOTES_MASTER_PERSIST_ID = 3;
const HANDOUT_PERSIST_ID = 4;
const FIRST_SLIDE_PERSIST_ID = 5;
/** Drawings: 1 main master, 2 notes master, 3 handout, then one per slide. */
const FIRST_SLIDE_DRAWING = 4;
const HEADER_TOKEN = 0xe391c05f;
const HEADER_TOKEN_ENCRYPTED = 0xf3d1c4df;
const SHAPE_RECTANGLE = 1;
const SHAPE_PICTURE_FRAME = 75;
const SHAPE_TEXT_BOX = 202;
const SHAPE_FLAGS = 0x0a00;
const PATRIARCH_FLAGS = 0x0005;
const BACKGROUND_FLAGS = 0x0c00;
const PROPERTY_PIB = 0x4104;
const PROPERTY_FILL_COLOR = 0x0181;
const PROPERTY_LINE_STYLE = 0x01ff;

/** Persist id of the slide at `index` (0-based), for PptWriteOptions.update. */
export function slidePersistId(index: number): number {
  return FIRST_SLIDE_PERSIST_ID + index;
}

const BLIPS: Record<PictureFormat, { recordType: number; instance: number; storeType: number; metafile: boolean }> = {
  emf: { recordType: 0xf01a, instance: 0x3d4, storeType: 2, metafile: true },
  wmf: { recordType: 0xf01b, instance: 0x216, storeType: 3, metafile: true },
  jpeg: { recordType: 0xf01d, instance: 0x46a, storeType: 5, metafile: false },
  png: { recordType: 0xf01e, instance: 0x6e0, storeType: 6, metafile: false },
  dib: { recordType: 0xf01f, instance: 0x7a8, storeType: 7, metafile: false },
  tiff: { recordType: 0xf029, instance: 0x6e4, storeType: 0x11, metafile: false },
};

export function recordHeader(type: number, length: number, instance = 0, version = 0): Buffer {
  const header = Buffer.alloc(8);
  header.writeUInt16LE(((instance & 0xfff) << 4) | (version & 0xf), 0);
  header.writeUInt16LE(type, 2);
  header.writeUInt32LE(length >>> 0, 4);
  return header;
}

export function atom(type: number, data: Uint8Array, instance = 0, version = 0): Buffer {
  return Buffer.concat([recordHeader(type, data.byteLength, instance, version), data]);
}

export function container(type: number, children: readonly Uint8Array[], instance = 0): Buffer {
  const data = Buffer.concat(children);
  return Buffer.concat([recordHeader(type, data.length, instance, 0xf), data]);
}

export function u8(...values: number[]): Buffer {
  return Buffer.from(values);
}

export function u16(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 2);
  values.forEach((value, i) => buffer.writeUInt16LE(value & 0xffff, i * 2));
  return buffer;
}

export function u32(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 4);
  values.forEach((value, i) => buffer.writeUInt32LE(value >>> 0, i * 4));
  return buffer;
}

function i16(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 2);
  values.forEach((value, i) => buffer.writeInt16LE(value, i * 2));
  return buffer;
}

/** A StyleTextPropAtom with one paragraph run per paragraph (varying optional fields) and one character run. */
function styleAtom(text: TextSpec): Buffer {
  const paragraphs = text.text.split('\r');
  const runs = paragraphs.map((paragraph, i) => {
    const level = text.levels?.[i] ?? 0;
    const bullet = text.bullets?.[i];
    let masks = 0x0800;
    const fields: Buffer[] = [];
    if (bullet !== undefined) {
      masks |= 0x0001;
      fields.push(u16(bullet ? 1 : 0));
    }
    fields.push(u16(0));
    if (level > 0) {
      masks |= 0x0100 | 0x0400;
      fields.push(u16(level * 576), u16(level * 576 - 288));
    }
    if (i === 0) {
      masks |= 0x100000;
      fields.push(u16(2), u16(720, 0), u16(1440, 0));
    }
    return Buffer.concat([u32(paragraph.length + 1), u16(level), u32(masks), ...fields]);
  });
  return atom(RECORD.StyleTextPropAtom, Buffer.concat([...runs, u32(text.text.length + 1), u32(0)]));
}

function textRecords(text: TextSpec): Buffer[] {
  const characters = text.bytes ? atom(RECORD.TextBytesAtom, Buffer.from(text.text, 'latin1')) : atom(RECORD.TextCharsAtom, Buffer.from(text.text, 'utf16le'));
  return [atom(RECORD.TextHeaderAtom, u32(TEXT_TYPES[text.type])), characters, styleAtom(text)];
}

interface ShapeOptions {
  spid: number;
  type: number;
  top: number;
  height?: number;
  properties?: [number, number][];
  /** Placeholder type and its position in the layout. */
  placeholder?: { type: number; position: number };
  textbox?: Buffer[];
}

function shape(options: ShapeOptions): Buffer {
  const properties = options.properties ?? [[PROPERTY_LINE_STYLE, 0x00080000]];
  const children = [
    atom(RECORD.FSP, u32(options.spid, SHAPE_FLAGS), options.type, 2),
    atom(RECORD.FOPT, Buffer.concat(properties.map(([id, value]) => Buffer.concat([u16(id), u32(value)]))), properties.length, 3),
    atom(RECORD.ClientAnchor, i16(options.top, 432, 5328, options.top + (options.height ?? 720))),
  ];
  if (options.placeholder !== undefined) {
    children.push(container(RECORD.ClientData, [atom(RECORD.PlaceholderAtom, Buffer.concat([u32(options.placeholder.position), u8(options.placeholder.type, 0), u16(0)]))]));
  }
  if (options.textbox) {
    children.push(container(RECORD.ClientTextbox, options.textbox));
  }
  return container(RECORD.SpContainer, children);
}

/** A text box shape with its text inside, at `top` (master units). */
export function textBoxShape(text: TextSpec, top: number): Buffer {
  return shape({ spid: 0, type: SHAPE_TEXT_BOX, top, textbox: textRecords(text) });
}

function drawing(drawingId: number, shapes: readonly Buffer[]): Buffer {
  const base = drawingId * 1024;
  const patriarch = container(RECORD.SpContainer, [atom(RECORD.FSPGR, Buffer.alloc(16), 0, 1), atom(RECORD.FSP, u32(base, PATRIARCH_FLAGS), 0, 2)]);
  const background = container(RECORD.SpContainer, [
    atom(RECORD.FSP, u32(base + shapes.length + 1, BACKGROUND_FLAGS), SHAPE_RECTANGLE, 2),
    atom(RECORD.FOPT, Buffer.concat([u16(PROPERTY_FILL_COLOR), u32(0x08000000)]), 1, 3),
  ]);
  return container(RECORD.Drawing, [
    container(RECORD.DgContainer, [
      atom(RECORD.FDG, u32(shapes.length + 2, base + shapes.length + 1), drawingId),
      container(RECORD.SpgrContainer, [patriarch, ...shapes]),
      background,
    ]),
  ]);
}

function slideAtom(placeholders: number[], masterId: number): Buffer {
  const types = Buffer.alloc(8);
  placeholders.slice(0, 8).forEach((type, i) => types.writeUInt8(type, i));
  const layout = placeholders[0] === PLACEHOLDER_IDS.centerTitle ? SLIDE_LAYOUT_TITLE : placeholders.length > 1 ? SLIDE_LAYOUT_TITLE_BODY : SLIDE_LAYOUT_TITLE_ONLY;
  return atom(RECORD.SlideAtom, Buffer.concat([u32(masterId ? layout : SLIDE_LAYOUT_TITLE_BODY), types, u32(masterId, 0), u16(masterId ? 0x0007 : 0, 0)]), 0, 2);
}

/** Instance 1: the slide's colour scheme; instance 6: an entry of the master's scheme list. */
function colorScheme(instance = 1): Buffer {
  return atom(RECORD.ColorSchemeAtom, u32(0x00ffffff, 0x00000000, 0x00808080, 0x00000000, 0x00ccffff, 0x00993333, 0x00999900, 0x00cc9900), instance);
}

function slideContainer(slide: SlideSpec, index: number): Buffer {
  const base = (FIRST_SLIDE_DRAWING + index) * 1024;
  let next = 0;
  let top = 288;
  const place = (spec?: TextSpec) => {
    const value = spec?.top ?? top;
    top = value + 864;
    return value;
  };
  let listIndex = 0;
  const shapes = [
    ...(slide.placeholders ?? []).map((text, i) =>
      shape({
        spid: base + ++next,
        type: SHAPE_RECTANGLE,
        top: place(text),
        placeholder: { type: PLACEHOLDER_IDS[text.type], position: i },
        textbox: text.inShape ? textRecords(text) : [atom(RECORD.OutlineTextRefAtom, u32(listIndex++))],
      }),
    ),
    ...(slide.textBoxes ?? []).map((text) => shape({ spid: base + ++next, type: SHAPE_TEXT_BOX, top: place(text), textbox: textRecords(text) })),
    ...(slide.pictures ?? []).map((picture) => shape({ spid: base + ++next, type: SHAPE_PICTURE_FRAME, top: place(), height: 1440, properties: [[PROPERTY_PIB, picture]] })),
    ...(slide.footer === undefined
      ? []
      : [shape({ spid: base + ++next, type: SHAPE_RECTANGLE, top: 4000, placeholder: { type: FOOTER_PLACEHOLDER, position: 2 }, textbox: textRecords({ type: 'other', text: slide.footer }) })]),
    ...(slide.extraShapes ?? []),
  ];
  const children = [slideAtom((slide.placeholders ?? []).map((text) => PLACEHOLDER_IDS[text.type]), MASTER_ID)];
  if (slide.hidden !== undefined) {
    children.push(atom(RECORD.SlideShowSlideInfoAtom, Buffer.concat([u32(0, 0), u8(0, 0), u16(slide.hidden ? 0x0004 : 0), u8(1, 0, 0, 0)])));
  }
  children.push(drawing(FIRST_SLIDE_DRAWING + index, shapes), colorScheme());
  return container(RECORD.Slide, children);
}

/** Master text style for one text type: font 0, a size and a scheme colour per level, no paragraph properties. */
function textMasterStyle(textType: number, fontSize: number, levels = 1): Buffer {
  const level = (i: number) =>
    Buffer.concat([textType >= 5 ? u16(i) : Buffer.alloc(0), u32(0), u32(0x10000 | 0x20000 | 0x40000), u16(0, fontSize - i * 4), u32(textType === 0 || textType === 6 ? 0x03000000 : 0x01000000)]);
  return atom(RECORD.TextMasterStyleAtom, Buffer.concat([u16(levels), ...Array.from({ length: levels }, (_, i) => level(i))]), textType);
}

function mainMaster(): Buffer {
  return container(RECORD.MainMaster, [
    slideAtom([0x01, 0x02, 0x07, 0x08, 0x09], 0),
    colorScheme(6),
    textMasterStyle(0, 44),
    textMasterStyle(1, 32, 5),
    textMasterStyle(2, 12),
    textMasterStyle(5, 32),
    textMasterStyle(6, 44),
    textMasterStyle(7, 28, 5),
    textMasterStyle(8, 24, 5),
    drawing(1, [
      shape({ spid: 1025, type: SHAPE_RECTANGLE, top: 288, placeholder: { type: 0x01, position: 0 }, textbox: textRecords({ type: 'title', text: 'Master title' }) }),
      shape({ spid: 1026, type: SHAPE_RECTANGLE, top: 1152, placeholder: { type: 0x02, position: 1 }, textbox: textRecords({ type: 'body', text: 'Master text' }) }),
    ]),
    colorScheme(),
  ]);
}

function fontEntity(name: string): Buffer {
  const face = Buffer.alloc(64);
  face.write(name, 0, 'utf16le');
  return atom(RECORD.FontEntityAtom, Buffer.concat([face, u8(0, 0, 0x04, 0x22)]));
}

/** The slide view's settings (guides off, 100% zoom); some readers refuse a presentation without them. */
function docInfoList(): Buffer {
  const zoom = Buffer.concat([u32(1, 1, 1, 1), Buffer.alloc(24), u32(0, 0), u8(0, 0), u16(0)]);
  return container(RECORD.DocInfoList, [
    container(RECORD.SlideViewInfo, [atom(RECORD.SlideViewInfoAtom, u8(0, 0, 0)), atom(RECORD.ZoomViewInfoAtom, zoom)]),
  ]);
}

function slidePersistAtom(persistId: number, texts: number, slideId: number): Buffer {
  return atom(RECORD.SlidePersistAtom, u32(persistId, 0, texts, slideId, 0));
}

interface StoredPicture {
  /** The blip record. */
  blip: Buffer;
  /** The picture store entry (with the blip inside when embedded). */
  entry: (pictureOffset: number) => Buffer;
}

function storedPicture(picture: PictureSpec): StoredPicture {
  const kind = BLIPS[picture.format];
  const uid = createHash('md5').update(picture.data).digest();
  const body = kind.metafile
    ? Buffer.concat([uid, u32(picture.data.byteLength, 0, 0, 1000, 1000, 9525000, 9525000, picture.data.byteLength), u8(0xfe, 0xfe), picture.data])
    : Buffer.concat([uid, u8(0xff), picture.data]);
  const blip = atom(kind.recordType, body, kind.instance);
  const entry = (pictureOffset: number) =>
    atom(
      RECORD.FBSE,
      Buffer.concat([
        u8(kind.storeType, kind.storeType),
        uid,
        u16(0xff),
        u32(blip.length, 1, picture.embedded ? 0 : pictureOffset),
        u8(0, 0, 0, 0),
        picture.embedded ? blip : Buffer.alloc(0),
      ]),
      kind.storeType,
      2,
    );
  return { blip, entry };
}

function listTexts(slide: SlideSpec): TextSpec[] {
  return [...(slide.placeholders ?? []).filter((text) => !text.inShape), ...(slide.listOnly ?? [])];
}

function documentContainer(spec: PptSpec, pictureEntries: readonly Buffer[]): Buffer {
  const size = spec.slideSize ?? { width: 5760, height: 4320 };
  const documentAtom = atom(
    RECORD.DocumentAtom,
    Buffer.concat([u32(size.width, size.height, 4320, 5760, 1, 2, NOTES_MASTER_PERSIST_ID, HANDOUT_PERSIST_ID), u16(1, 0), u8(0, 0, 0, 1)]),
    0,
    1,
  );
  const drawingCount = spec.slides.length + FIRST_SLIDE_DRAWING - 1;
  const clusters = Array.from({ length: drawingCount }, (_, i) => u32(i + 1, (i + 2) * 1024));
  const drawingGroup = container(RECORD.DrawingGroup, [
    container(RECORD.DggContainer, [
      atom(RECORD.FDGGBlock, Buffer.concat([u32((drawingCount + 1) * 1024, drawingCount + 1, 0, drawingCount), ...clusters])),
      container(RECORD.BStoreContainer, pictureEntries, pictureEntries.length),
    ]),
  ]);
  const masterList = container(RECORD.SlideListWithText, [slidePersistAtom(MASTER_PERSIST_ID, 0, MASTER_ID)], 1);
  const slideList = container(
    RECORD.SlideListWithText,
    spec.slides.flatMap((slide, i) => [
      slidePersistAtom(FIRST_SLIDE_PERSIST_ID + i, listTexts(slide).length, FIRST_SLIDE_ID + i),
      ...listTexts(slide).flatMap(textRecords),
    ]),
    0,
  );
  return container(RECORD.Document, [
    documentAtom,
    container(RECORD.Environment, [
      container(RECORD.FontCollection, [fontEntity('Arial')]),
      atom(RECORD.TextCFExceptionAtom, Buffer.concat([u32(0x10000 | 0x20000 | 0x40000), u16(0, 18), u32(0xfe000000)])),
      atom(RECORD.TextPFExceptionAtom, Buffer.concat([u16(0), u32(0x0800 | 0x1000), u16(0, 100)])),
      atom(RECORD.TextSIExceptionAtom, u32(0)),
      textMasterStyle(4, 18),
    ]),
    drawingGroup,
    masterList,
    docInfoList(),
    slideList,
    atom(RECORD.EndDocumentAtom, Buffer.alloc(0)),
  ]);
}

/** Persist objects of one save, by persist id. */
function persistObjects(spec: PptSpec, pictures: readonly StoredPicture[], pictureOffsets: readonly number[]): Map<number, Buffer> {
  const objects = new Map<number, Buffer>();
  objects.set(DOCUMENT_PERSIST_ID, documentContainer(spec, pictures.map((picture, i) => picture.entry(pictureOffsets[i] ?? 0))));
  objects.set(MASTER_PERSIST_ID, mainMaster());
  objects.set(NOTES_MASTER_PERSIST_ID, container(RECORD.Notes, [atom(RECORD.NotesAtom, u32(MASTER_ID, 0), 0, 1), drawing(2, []), colorScheme()]));
  objects.set(HANDOUT_PERSIST_ID, container(RECORD.Handout, [drawing(3, []), colorScheme()]));
  spec.slides.forEach((slide, i) => objects.set(FIRST_SLIDE_PERSIST_ID + i, slideContainer(slide, i)));
  if (spec.encrypted?.session) {
    objects.set(FIRST_SLIDE_PERSIST_ID + spec.slides.length, container(RECORD.CryptSession10Container, [Buffer.alloc(16)]));
  }
  return objects;
}

function persistDirectory(offsets: ReadonlyMap<number, number>): Buffer {
  const ids = [...offsets.keys()].sort((a, b) => a - b);
  const entries: Buffer[] = [];
  for (let i = 0; i < ids.length; ) {
    let end = i;
    while (end + 1 < ids.length && ids[end + 1] === (ids[end] ?? 0) + 1) {
      end += 1;
    }
    const run = ids.slice(i, end + 1);
    entries.push(u32(((run.length << 20) | (run[0] ?? 0)) >>> 0), ...run.map((id) => u32(offsets.get(id) ?? 0)));
    i = end + 1;
  }
  return atom(RECORD.PersistDirectoryAtom, Buffer.concat(entries));
}

function userEdit(previous: number, directory: number, seed: number, encryptRef?: number): Buffer {
  const fields = [u32(FIRST_SLIDE_ID), u16(0), u8(0, 3), u32(previous, directory, DOCUMENT_PERSIST_ID, seed), u16(1, 0)];
  if (encryptRef !== undefined) {
    fields.push(u32(encryptRef));
  }
  return atom(RECORD.UserEditAtom, Buffer.concat(fields));
}

export function currentUserStream(offsetToCurrentEdit: number, encrypted = false): Buffer {
  const user = 'Tester';
  return atom(
    RECORD.CurrentUserAtom,
    Buffer.concat([
      u32(0x14, encrypted ? HEADER_TOKEN_ENCRYPTED : HEADER_TOKEN, offsetToCurrentEdit),
      u16(user.length, 0x03f4),
      u8(3, 0),
      u16(0),
      Buffer.from(user, 'latin1'),
      u32(8),
      Buffer.from(user, 'utf16le'),
    ]),
  );
}

export function pptStreams(spec: PptSpec, options: PptWriteOptions = {}): PptStreams {
  const pictures = (spec.pictures ?? []).map(storedPicture);
  const pictureOffsets: number[] = [];
  let picturesLength = 0;
  for (const [i, picture] of pictures.entries()) {
    pictureOffsets.push(picturesLength);
    picturesLength += spec.pictures?.[i]?.embedded ? 0 : picture.blip.length;
  }
  const picturesStream = Buffer.concat(pictures.filter((_, i) => !spec.pictures?.[i]?.embedded).map((picture) => picture.blip));

  const parts: Buffer[] = [];
  let length = 0;
  const append = (record: Buffer) => {
    const offset = length;
    parts.push(record);
    length += record.length;
    return offset;
  };
  const offsets = new Map<number, number>();
  const edits: number[] = [];
  const save = (objects: Map<number, Buffer>, previous: number) => {
    const saved = new Map<number, number>();
    for (const [id, record] of objects) {
      saved.set(id, append(record));
      offsets.set(id, saved.get(id) ?? 0);
    }
    const directory = append(persistDirectory(saved));
    const sessionId = spec.encrypted?.session ? FIRST_SLIDE_PERSIST_ID + spec.slides.length : undefined;
    edits.push(append(userEdit(previous, directory, Math.max(...offsets.keys()) + 1, sessionId)));
  };

  save(persistObjects(spec, pictures, pictureOffsets), 0);
  if (options.update) {
    const updated = persistObjects(options.update.spec, pictures, pictureOffsets);
    save(new Map(options.update.persistIds.map((id) => [id, updated.get(id) ?? Buffer.alloc(0)])), edits.at(-1) ?? 0);
  }
  return {
    document: Buffer.concat(parts),
    currentUser: currentUserStream(edits.at(-1) ?? 0, spec.encrypted?.token),
    pictures: picturesStream,
    offsets,
    edits,
  };
}

/** The compound file around a presentation's streams; a stream left out (or null) is not written. */
export function pptFile(streams: Partial<Record<string, Uint8Array | null>>, options: CompoundFileOptions = {}): Buffer {
  return writeCompoundFile(
    Object.entries(streams)
      .filter((entry): entry is [string, Uint8Array] => entry[1] instanceof Uint8Array && (entry[0] !== 'Pictures' || entry[1].byteLength > 0))
      .map(([name, data]) => ({ name, data })),
    options,
  );
}

/** CLSID of PowerPoint 97–2003 presentations ({64818D10-4F9B-11CF-86EA-00AA00B929E8}), set on the root storage. */
export const PRESENTATION_CLSID = Buffer.from('108d81649b4fcf1186ea00aa00b929e8', 'hex');

export function writePpt(spec: PptSpec, options: PptWriteOptions & { compound?: CompoundFileOptions } = {}): Buffer {
  const streams = pptStreams(spec, options);
  return pptFile(
    { 'PowerPoint Document': streams.document, 'Current User': streams.currentUser, Pictures: streams.pictures },
    { rootClsid: PRESENTATION_CLSID, ...options.compound },
  );
}
