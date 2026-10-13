import type { ColorContext } from './color.ts';
import type { Deck, SlideInfo } from './deck.ts';
import { PPTX_LIMITS } from './limits.ts';
import type { MediaStore } from './media.ts';
import type { Relationship, Relationships } from './package.ts';
import type { Theme } from './theme.ts';

/** What is left of the per-slide caps (PPTX_LIMITS); `truncated` records that something was left out. */
export interface RenderBudget {
  shapes: number;
  elements: number;
  runs: number;
  partBytes: number;
  truncated: boolean;
}

export function slideBudget(): RenderBudget {
  return {
    shapes: PPTX_LIMITS.maxShapesPerSlide,
    elements: PPTX_LIMITS.maxElementsPerSlide,
    runs: PPTX_LIMITS.maxRunsPerSlide,
    partBytes: PPTX_LIMITS.maxPartBytesPerSlide,
    truncated: false,
  };
}

/** Takes `cost` elements from the slide's budget; false, with the slide marked truncated, when they are not left. */
export function spend(budget: RenderBudget, cost: number): boolean {
  if (cost > budget.elements) {
    budget.elements = 0;
    budget.truncated = true;
    return false;
  }
  budget.elements -= cost;
  return true;
}

export interface RenderContext {
  readonly deck: Deck;
  readonly slide: SlideInfo;
  readonly theme: Theme;
  /** The slide's colors: theme scheme and the slide's effective color map, also used for master and layout shapes. */
  readonly colors: ColorContext;
  /** The part whose relationships resolve r:ids (the slide, its layout or master, or a SmartArt drawing). */
  readonly part: string;
  readonly rels: Relationships;
  readonly media: MediaStore;
  readonly budget: RenderBudget;
  /** The number shown by slide number fields. */
  readonly slideNumber: number;
}

export function relationshipTarget(context: RenderContext, id: string | undefined): Relationship | undefined {
  return id === undefined ? undefined : context.rels.get(id);
}
