// The shape tree (p:spTree) as positioned DOM: each shape is a box at its frame, rotated with CSS, holding an SVG of
// its geometry (fill, outline, picture) and an HTML text layer. Placeholders on a slide inherit position, geometry,
// outline and text properties from their layout and master placeholders.
import { matchPlaceholder, placeholderKind, placeholderOf, type SlideInfo } from './deck.ts';
import { buildGeometry, customGeometry, formulasOf, presetGeometry, type BuiltGeometry, type Rect } from './geometry.ts';
import { PPTX_LIMITS } from './limits.ts';
import {
  applyLine,
  applyPaint,
  modifyPaint,
  NO_PAINT,
  readBlipFill,
  readFill,
  readLine,
  shadowFilter,
  styleFill,
  styleReference,
  svg,
  type LineSource,
  type Paint,
} from './paint.ts';
import { relationshipTarget, type RenderContext } from './renderContext.ts';
import { renderTable } from './table.ts';
import { hasText, hyperlink, readBodyProperties, renderTextBody, setBox } from './text.ts';
import { FILL_ELEMENTS, fontRefDefaults, TextStyleResolver, type TextLayers } from './textStyle.ts';
import { EMU_PER_PX, mapIntoGroup, readGroupXfrm, readXfrm, type GroupXfrm, type Xfrm } from './units.ts';
import { attr, child, find, kids, oneOf, relAttr } from './xml.ts';

/** Whose shape tree is drawn: only a slide's own placeholders are drawn (and inherit). */
export type Scope = 'slide' | 'layout' | 'master' | 'drawing';

/** Where members of a shape tree go, and how their frames map into it. */
export interface Container {
  readonly element: HTMLElement;
  readonly scope: Scope;
  readonly depth: number;
  /** Maps a member's frame (EMUs in the tree's coordinate space) to EMUs relative to `element`. */
  readonly place: (frame: Xfrm) => Xfrm;
  /** The enclosing group's fill, for members with grpFill. */
  readonly groupFill?: Paint;
}

interface Inherited {
  readonly layout?: Element;
  readonly master?: Element;
  /** For a slide placeholder, which master text style its text starts from. */
  readonly textKind?: 'title' | 'body' | 'other';
}

export async function renderShapeTree(tree: Element, container: Container, context: RenderContext): Promise<void> {
  for (const member of kids(tree)) {
    if (context.budget.shapes <= 0) {
      context.budget.truncated = true;
      return;
    }
    if (!DRAWN.has(member.localName)) {
      continue;
    }
    if (container.scope !== 'slide' && container.scope !== 'drawing' && placeholderOf(member)) {
      continue;
    }
    context.budget.shapes -= 1;
    try {
      const element = await renderMember(member, container, context);
      if (element) {
        container.element.append(element);
      }
    } catch {
      // One malformed shape must not take the slide down with it; the rest of the slide still renders.
      context.budget.truncated = true;
    }
  }
}

const DRAWN = new Set(['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame', 'contentPart']);

function renderMember(member: Element, container: Container, context: RenderContext): Promise<HTMLElement | undefined> {
  switch (member.localName) {
    case 'sp':
    case 'cxnSp':
      return renderShape(member, container, context);
    case 'pic':
      return renderPicture(member, container, context);
    case 'grpSp':
      return renderGroup(member, container, context);
    case 'graphicFrame':
      return renderGraphicFrame(member, container, context);
    default:
      return Promise.resolve(renderInk(member, container, context));
  }
}

function inheritance(shape: Element, container: Container, slide: SlideInfo): Inherited {
  const ph = container.scope === 'slide' && container.depth === 0 ? placeholderOf(shape) : undefined;
  if (!ph) {
    return {};
  }
  const layout = slide.layout && matchPlaceholder(slide.layout.placeholders, ph, true);
  const layoutPh = layout && placeholderOf(layout);
  const master = slide.master && matchPlaceholder(slide.master.placeholders, layoutPh ?? ph, false);
  const kind = placeholderKind(attr(ph, 'type') ?? 'obj');
  return { layout, master, textKind: kind === 'title' ? 'title' : kind === 'body' ? 'body' : 'other' };
}

/** The shape itself, then its layout and master placeholders. */
function chainOf(shape: Element, inherited: Inherited): Element[] {
  return [shape, inherited.layout, inherited.master].filter((item): item is Element => item !== undefined);
}

function propertiesOf(shape: Element): Element | undefined {
  return oneOf(shape, ['spPr', 'grpSpPr']);
}

function frameOf(chain: readonly Element[]): Xfrm | undefined {
  for (const shape of chain) {
    const xfrm = readXfrm(shape.localName === 'graphicFrame' ? child(shape, 'xfrm') : child(propertiesOf(shape), 'xfrm'));
    if (xfrm) {
      return xfrm;
    }
  }
  return undefined;
}

/** The positioned box of a shape: an <a> when the shape itself is a hyperlink, labelled with the shape's name. */
function shapeBox(member: Element, frame: Xfrm, context: RenderContext): HTMLElement {
  const properties = child(
    kids(member).find((node) => node.localName.startsWith('nv') && node.localName.endsWith('Pr')),
    'cNvPr',
  );
  const element = hyperlink(child(properties, 'hlinkClick'), context) ?? document.createElement('div');
  element.classList.add('pptx-shape');
  const name = attr(properties, 'name');
  if (name) {
    element.dataset['name'] = name.slice(0, 200);
  }
  setBox(element, frame.x / EMU_PER_PX, frame.y / EMU_PER_PX, frame.w / EMU_PER_PX, frame.h / EMU_PER_PX);
  if (frame.rot) {
    element.style.setProperty('transform', `rotate(${round(frame.rot)}deg)`);
  }
  return element;
}

function geometryOf(chain: readonly Element[], frame: Xfrm, connector: boolean): BuiltGeometry {
  for (const shape of chain) {
    const properties = propertiesOf(shape);
    const preset = child(properties, 'prstGeom');
    if (preset) {
      const def = presetGeometry(attr(preset, 'prst') ?? 'rect') ?? presetGeometry(connector ? 'line' : 'rect');
      if (def) {
        return buildGeometry(def, frame.w, frame.h, formulasOf(child(preset, 'avLst')));
      }
    }
    const custom = child(properties, 'custGeom');
    if (custom) {
      return buildGeometry(customGeometry(custom), frame.w, frame.h);
    }
  }
  const fallback = presetGeometry(connector ? 'line' : 'rect');
  return fallback ? buildGeometry(fallback, frame.w, frame.h) : { paths: [], textRect: { x: 0, y: 0, w: frame.w / EMU_PER_PX, h: frame.h / EMU_PER_PX } };
}

async function fillOf(chain: readonly Element[], container: Container, context: RenderContext): Promise<Paint> {
  for (const shape of chain) {
    const fill = oneOf(propertiesOf(shape), FILL_ELEMENTS);
    if (fill?.localName === 'grpFill') {
      return container.groupFill ?? NO_PAINT;
    }
    const paint = await readFill(fill, context.colors, context);
    if (paint) {
      return paint;
    }
  }
  const style = chain.map((shape) => child(shape, 'style')).find((item) => item !== undefined);
  return (await styleFill(child(style, 'fillRef'), context.theme, context.colors, context)) ?? NO_PAINT;
}

function lineSources(chain: readonly Element[], context: RenderContext): LineSource[] {
  const sources: LineSource[] = [];
  for (const shape of chain) {
    const ln = child(propertiesOf(shape), 'ln');
    if (ln) {
      sources.push({ ln, colors: context.colors });
    }
  }
  const style = chain.map((shape) => child(shape, 'style')).find((item) => item !== undefined);
  const reference = styleReference(child(style, 'lnRef'), context.theme.lineStyles, context.colors);
  if (reference) {
    sources.push({ ln: reference.element, colors: reference.colors });
  }
  return sources;
}

function shadowOf(chain: readonly Element[], context: RenderContext): string | undefined {
  for (const shape of chain) {
    const effects = child(propertiesOf(shape), 'effectLst');
    if (effects) {
      return shadowFilter(effects, context.colors);
    }
  }
  const style = chain.map((shape) => child(shape, 'style')).find((item) => item !== undefined);
  const reference = styleReference(child(style, 'effectRef'), context.theme.effectStyles, context.colors);
  return reference ? shadowFilter(child(reference.element, 'effectLst'), reference.colors) : undefined;
}

/** An SVG canvas over the shape box, with geometry flips applied to its content group. */
function geometryCanvas(frame: Xfrm): { root: SVGSVGElement; defs: SVGDefsElement; group: SVGGElement; w: number; h: number } {
  const w = frame.w / EMU_PER_PX;
  const h = frame.h / EMU_PER_PX;
  const root = svg('svg', { class: 'pptx-geometry', width: round(Math.max(w, 1)), height: round(Math.max(h, 1)) });
  const defs = svg('defs');
  const group = svg('g');
  if (frame.flipH || frame.flipV) {
    group.setAttribute('transform', `matrix(${frame.flipH ? -1 : 1} 0 0 ${frame.flipV ? -1 : 1} ${frame.flipH ? round(w) : 0} ${frame.flipV ? round(h) : 0})`);
  }
  root.append(defs, group);
  return { root, defs, group, w, h };
}

async function renderShape(shape: Element, container: Container, context: RenderContext): Promise<HTMLElement | undefined> {
  const inherited = inheritance(shape, container, context.slide);
  const chain = chainOf(shape, inherited);
  const own = frameOf(chain);
  if (!own) {
    return undefined;
  }
  const frame = container.place(own);
  const connector = shape.localName === 'cxnSp';
  const box = shapeBox(shape, frame, context);
  const geometry = geometryOf(chain, frame, connector);
  const fill = connector ? NO_PAINT : await fillOf(chain, container, context);
  const line = await readLine(lineSources(chain, context), context);
  const visibleLine = line !== undefined && line.paint.kind !== 'none';
  if (fill.kind !== 'none' || visibleLine) {
    const canvas = geometryCanvas(frame);
    for (const path of geometry.paths) {
      const element = svg('path', { d: path.d });
      if (path.fill === 'none') {
        element.setAttribute('fill', 'none');
      } else {
        applyPaint(element, 'fill', modifyPaint(fill, path.fill), canvas.defs, canvas);
      }
      applyLine(element, path.stroke ? line : undefined, canvas.defs, canvas, !path.d.includes('Z'));
      canvas.group.append(element);
    }
    const shadow = shadowOf(chain, context);
    if (shadow) {
      canvas.root.style.setProperty('filter', shadow);
    }
    box.append(canvas.root);
  }
  const txBody = child(shape, 'txBody');
  if (txBody && hasText(txBody)) {
    box.append(renderText(shape, txBody, chain, inherited, textRectOf(shape, own, frame, geometry), frame, context));
  }
  return box;
}

/** SmartArt drawings place text in their own rectangle (dsp:txXfrm) rather than the geometry's. */
function textRectOf(shape: Element, own: Xfrm, frame: Xfrm, geometry: BuiltGeometry): Rect {
  const txXfrm = readXfrm(child(shape, 'txXfrm'));
  if (!txXfrm) {
    return geometry.textRect;
  }
  const sx = own.w > 0 ? frame.w / own.w : 1;
  const sy = own.h > 0 ? frame.h / own.h : 1;
  return { x: ((txXfrm.x - own.x) * sx) / EMU_PER_PX, y: ((txXfrm.y - own.y) * sy) / EMU_PER_PX, w: (txXfrm.w * sx) / EMU_PER_PX, h: (txXfrm.h * sy) / EMU_PER_PX };
}

function renderText(shape: Element, txBody: Element, chain: readonly Element[], inherited: Inherited, box: Rect, frame: Xfrm, context: RenderContext): HTMLElement {
  const layers: TextLayers = {
    base: baseTextStyles(context, inherited.textKind),
    shape: [...chain].reverse().map((item) => child(child(item, 'txBody'), 'lstStyle')),
    defaults: fontRefDefaults(find(shape, 'style', 'fontRef')),
  };
  return renderTextBody(txBody, {
    box,
    body: readBodyProperties(chain.map((item) => child(child(item, 'txBody'), 'bodyPr'))),
    styles: new TextStyleResolver(layers),
    context,
    flipV: frame.flipV,
    shapeRotation: frame.rot,
  });
}

/**
 * The list styles text starts from. Placeholders build on the master's title, body or other style; text in other
 * shapes (and tables) follows the presentation's default text style, which PowerPoint ranks above the master's
 * otherStyle for them.
 */
export function baseTextStyles(context: RenderContext, placeholderKind: Inherited['textKind']): (Element | undefined)[] {
  const master = context.slide.master;
  return placeholderKind ? [context.deck.defaultTextStyle, master?.textStyles[placeholderKind]] : [master?.textStyles.other, context.deck.defaultTextStyle];
}

async function renderPicture(picture: Element, container: Container, context: RenderContext): Promise<HTMLElement | undefined> {
  const inherited = inheritance(picture, container, context.slide);
  const chain = chainOf(picture, inherited);
  const own = frameOf(chain);
  if (!own) {
    return undefined;
  }
  const frame = container.place(own);
  const box = shapeBox(picture, frame, context);
  const description = attr(find(picture, 'nvPicPr', 'cNvPr'), 'descr');
  if (description) {
    box.setAttribute('role', 'img');
    box.setAttribute('aria-label', description.slice(0, 500));
  }
  const blipFill = child(picture, 'blipFill');
  const paint = blipFill ? await readBlipFill(blipFill, context) : NO_PAINT;
  const geometry = geometryOf(chain, frame, false);
  const line = await readLine(lineSources(chain, context), context);
  const canvas = geometryCanvas(frame);
  if (paint.kind === 'image' && 'url' in paint.picture) {
    for (const path of geometry.paths.filter((item) => item.fill !== 'none')) {
      const element = svg('path', { d: path.d, stroke: 'none' });
      applyPaint(element, 'fill', paint, canvas.defs, canvas);
      canvas.group.append(element);
    }
    if (paint.grayscale) {
      canvas.root.style.setProperty('filter', 'grayscale(1)');
    }
  } else {
    box.append(placeholderLabel(paint.kind === 'image' && 'unsupported' in paint.picture ? paint.picture.unsupported : 'Picture'));
  }
  for (const path of geometry.paths.filter((item) => item.stroke)) {
    const element = svg('path', { d: path.d, fill: 'none' });
    applyLine(element, line, canvas.defs, canvas, !path.d.includes('Z'));
    canvas.group.append(element);
  }
  const shadow = shadowOf(chain, context);
  if (shadow) {
    canvas.root.style.setProperty('filter', shadow);
  }
  box.prepend(canvas.root);
  const media = mediaKind(picture);
  if (media) {
    const badge = document.createElement('span');
    badge.className = 'pptx-media-badge';
    badge.textContent = media;
    box.append(badge);
  }
  return box;
}

function mediaKind(picture: Element): string | undefined {
  const nvPr = find(picture, 'nvPicPr', 'nvPr');
  if (child(nvPr, 'videoFile') || child(nvPr, 'quickTimeFile')) {
    return 'Video';
  }
  if (child(nvPr, 'audioFile') || child(nvPr, 'audioCd') || child(nvPr, 'wavAudioFile')) {
    return 'Audio';
  }
  const media = kids(child(nvPr, 'extLst'))
    .map((ext) => child(ext, 'media'))
    .find((item) => item !== undefined);
  return media ? 'Media' : undefined;
}

async function renderGroup(group: Element, container: Container, context: RenderContext): Promise<HTMLElement | undefined> {
  if (container.depth + 1 > PPTX_LIMITS.maxGroupDepth) {
    context.budget.truncated = true;
    return undefined;
  }
  const properties = child(group, 'grpSpPr');
  const own = readGroupXfrm(child(properties, 'xfrm'));
  if (!own) {
    return undefined;
  }
  const frame = container.place(own);
  const box = shapeBox(group, frame, context);
  box.classList.add('pptx-group');
  const paint = await readFill(oneOf(properties, FILL_ELEMENTS), context.colors, context);
  const groupXfrm: GroupXfrm = { ...own, x: 0, y: 0, w: frame.w, h: frame.h };
  await renderShapeTree(group, {
    element: box,
    scope: container.scope,
    depth: container.depth + 1,
    place: (member) => mirror(mapIntoGroup(member, { ...groupXfrm, chX: own.chX, chY: own.chY }), frame),
    groupFill: paint ?? container.groupFill,
  }, context);
  return box;
}

/** A flipped group mirrors its members' positions and flips them, without mirroring their text. */
export function mirror(member: Xfrm, group: Xfrm): Xfrm {
  let result = member;
  if (group.flipH) {
    result = { ...result, x: group.w - result.x - result.w, rot: -result.rot, flipH: !result.flipH };
  }
  if (group.flipV) {
    result = { ...result, y: group.h - result.y - result.h, rot: -result.rot, flipV: !result.flipV };
  }
  return result;
}

const GRAPHIC_TABLE = 'http://schemas.openxmlformats.org/drawingml/2006/table';

async function renderGraphicFrame(frameElement: Element, container: Container, context: RenderContext): Promise<HTMLElement | undefined> {
  const inherited = inheritance(frameElement, container, context.slide);
  const own = frameOf(chainOf(frameElement, inherited));
  if (!own) {
    return undefined;
  }
  const frame = container.place(own);
  const data = find(frameElement, 'graphic', 'graphicData');
  const uri = attr(data, 'uri') ?? '';
  const box = shapeBox(frameElement, frame, context);
  if (uri === GRAPHIC_TABLE) {
    const table = child(data, 'tbl');
    if (table) {
      box.classList.add('pptx-table-frame');
      box.append(await renderTable(table, context));
      return box;
    }
  }
  if (uri.endsWith('/chart') || uri.endsWith('/chartex')) {
    box.append(placeholderLabel(await chartLabel(data, context)));
    return box;
  }
  if (uri.endsWith('/diagram')) {
    const drawing = await smartArtDrawing(data, context);
    if (drawing) {
      box.classList.add('pptx-group');
      const scale = { ...frame, x: 0, y: 0, chX: 0, chY: 0, chW: own.w, chH: own.h };
      await renderShapeTree(drawing.tree, { element: box, scope: 'drawing', depth: container.depth + 1, place: (member) => mapIntoGroup(member, scale) }, { ...context, part: drawing.part, rels: await context.deck.pkg.rels(drawing.part) });
      return box;
    }
    box.append(placeholderLabel('SmartArt graphic'));
    return box;
  }
  if (uri.endsWith('/ole')) {
    const preview = findDescendant(data, 'pic', 4);
    if (preview) {
      const blip = await readBlipFill(child(preview, 'blipFill') ?? preview, context);
      if (blip.kind === 'image' && 'url' in blip.picture) {
        const canvas = geometryCanvas(frame);
        const rect = svg('rect', { width: round(canvas.w), height: round(canvas.h) });
        applyPaint(rect, 'fill', blip, canvas.defs, canvas);
        canvas.group.append(rect);
        box.append(canvas.root);
        return box;
      }
    }
    box.append(placeholderLabel('Embedded object'));
    return box;
  }
  box.append(placeholderLabel('Embedded content'));
  return box;
}

function findDescendant(root: Element | undefined, name: string, depth: number): Element | undefined {
  if (!root || depth < 0) {
    return undefined;
  }
  for (const node of kids(root)) {
    if (node.localName === name) {
      return node;
    }
    const found = findDescendant(node, name, depth - 1);
    if (found) {
      return found;
    }
  }
  return undefined;
}

const CHART_NAMES: Readonly<Record<string, string>> = {
  barChart: 'Bar chart',
  bar3DChart: 'Bar chart',
  lineChart: 'Line chart',
  line3DChart: 'Line chart',
  pieChart: 'Pie chart',
  pie3DChart: 'Pie chart',
  doughnutChart: 'Doughnut chart',
  areaChart: 'Area chart',
  area3DChart: 'Area chart',
  scatterChart: 'Scatter chart',
  radarChart: 'Radar chart',
  bubbleChart: 'Bubble chart',
  stockChart: 'Stock chart',
  surfaceChart: 'Surface chart',
  surface3DChart: 'Surface chart',
  ofPieChart: 'Pie chart',
};

async function chartLabel(data: Element | undefined, context: RenderContext): Promise<string> {
  const target = relationshipTarget(context, relAttr(child(data, 'chart'), 'id'));
  if (!target || target.external || !context.deck.pkg.has(target.target)) {
    return 'Chart';
  }
  try {
    const plot = find(await context.deck.pkg.xml(target.target), 'chart', 'plotArea');
    const kind = kids(plot)
      .map((node) => CHART_NAMES[node.localName])
      .find((name) => name !== undefined);
    return kind ?? 'Chart';
  } catch {
    return 'Chart';
  }
}

/** The pre-drawn shapes PowerPoint stores for a SmartArt graphic (diagramDrawing part), when present. */
async function smartArtDrawing(data: Element | undefined, context: RenderContext): Promise<{ tree: Element; part: string } | undefined> {
  const dataModel = relationshipTarget(context, relAttr(child(data, 'relIds'), 'dm'));
  if (!dataModel || dataModel.external || !context.deck.pkg.has(dataModel.target)) {
    return undefined;
  }
  try {
    const model = await context.deck.pkg.xml(dataModel.target);
    const relId = kids(child(model, 'extLst'))
      .map((ext) => attr(child(ext, 'dataModelExt'), 'relId'))
      .find((id) => id !== undefined);
    const drawing = (relId && context.rels.get(relId)) || firstDrawing(context);
    if (!drawing || drawing.external || !context.deck.pkg.has(drawing.target)) {
      return undefined;
    }
    const tree = child(await context.deck.pkg.xml(drawing.target), 'spTree');
    return tree ? { tree, part: drawing.target } : undefined;
  } catch {
    return undefined;
  }
}

function firstDrawing(context: RenderContext) {
  // Only an unambiguous match: with several SmartArt graphics on a slide, the data model must name its drawing.
  const drawings = [...context.rels.values()].filter((rel) => rel.type === 'diagramDrawing');
  return drawings.length === 1 ? drawings[0] : undefined;
}

function renderInk(member: Element, container: Container, context: RenderContext): HTMLElement | undefined {
  const own = readXfrm(child(member, 'xfrm'));
  if (!own) {
    return undefined;
  }
  const box = shapeBox(member, container.place(own), context);
  box.append(placeholderLabel('Ink'));
  return box;
}

/** A labelled stand-in for content the viewer does not draw (charts, embedded objects, unsupported pictures). */
export function placeholderLabel(text: string): HTMLElement {
  const element = document.createElement('div');
  element.className = 'pptx-placeholder';
  const label = document.createElement('span');
  label.textContent = text;
  element.append(label);
  return element;
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}
