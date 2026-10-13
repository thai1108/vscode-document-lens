// One slide as a fixed-size element in CSS pixels: background, then the master's and the layout's own shapes (unless
// the slide hides them), then the slide's shapes.
import { toCss, type ColorContext } from './color.ts';
import type { Deck, PartInfo, SlideInfo } from './deck.ts';
import type { MediaStore } from './media.ts';
import { applyPaint, paintTarget, readFill, styleFill, svg, type Paint } from './paint.ts';
import { slideBudget, type RenderContext } from './renderContext.ts';
import { renderShapeTree, type Scope } from './shapes.ts';
import { setBox } from './text.ts';
import { EMU_PER_PX } from './units.ts';
import { child, find, kids } from './xml.ts';

export interface RenderedSlide {
  readonly element: HTMLElement;
  readonly hidden: boolean;
  /** Something on the slide was left out (a cap was reached or a shape could not be drawn). */
  readonly truncated: boolean;
}

export async function renderSlide(deck: Deck, index: number, media: MediaStore): Promise<RenderedSlide> {
  const slide = await deck.slide(index);
  const width = deck.width / EMU_PER_PX;
  const height = deck.height / EMU_PER_PX;
  const element = document.createElement('div');
  element.className = 'pptx-slide';
  element.style.setProperty('width', `${width}px`);
  element.style.setProperty('height', `${height}px`);

  const colors: ColorContext = { scheme: slide.theme.colors, map: slide.colorMap };
  const base: RenderContext = {
    deck,
    slide,
    theme: slide.theme,
    colors,
    part: slide.part,
    rels: slide.rels,
    media,
    budget: slideBudget(),
    slideNumber: deck.firstSlideNumber + index,
  };
  const forPart = (part: PartInfo): RenderContext => ({ ...base, part: part.part, rels: part.rels });

  try {
    await renderBackground(element, slide, width, height, forPart);
  } catch {
    // A background picture that cannot be read leaves the slide on its default background, not blank.
    base.budget.truncated = true;
  }
  const layers: [PartInfo | undefined, Scope, boolean][] = [
    [slide.master, 'master', slide.showMasterShapes && slide.layout?.showMasterShapes !== false],
    [slide.layout, 'layout', slide.showMasterShapes],
    [slide, 'slide', true],
  ];
  for (const [part, scope, shown] of layers) {
    const tree = part && find(part.root, 'cSld', 'spTree');
    if (shown && tree) {
      await renderShapeTree(tree, { element, scope, depth: 0, place: (frame) => frame }, forPart(part));
    }
  }
  return { element, hidden: slide.hidden, truncated: base.budget.truncated };
}

/** The first of slide, layout and master that defines a background, drawn as an SVG rectangle. */
async function renderBackground(element: HTMLElement, slide: SlideInfo, width: number, height: number, forPart: (part: PartInfo) => RenderContext): Promise<void> {
  for (const part of [slide, slide.layout, slide.master]) {
    const background = part && find(part.root, 'cSld', 'bg');
    if (!part || !background) {
      continue;
    }
    const context = forPart(part);
    const paint = await backgroundPaint(background, context);
    if (!paint) {
      continue;
    }
    if (paint.kind === 'solid') {
      element.style.setProperty('background-color', toCss(paint.color));
      return;
    }
    const canvas = svg('svg', { class: 'pptx-background', width: round(width), height: round(height) });
    const defs = svg('defs');
    const rect = svg('rect', { width: round(width), height: round(height) });
    applyPaint(rect, 'fill', paint, paintTarget(defs, { w: width, h: height }, context.budget));
    canvas.append(defs, rect);
    setBox(canvas, 0, 0, width, height);
    element.prepend(canvas);
    return;
  }
}

async function backgroundPaint(background: Element, context: RenderContext): Promise<Paint | undefined> {
  const properties = child(background, 'bgPr');
  if (properties) {
    return readFill(kids(properties).find((node) => node.localName.endsWith('Fill')), context.colors, context);
  }
  return styleFill(child(background, 'bgRef'), context.theme, context.colors, context);
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}
