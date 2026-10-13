import type { ColorContext } from './color.ts';
import type { Deck, SlideInfo } from './deck.ts';
import type { MediaStore } from './media.ts';
import type { Relationship, Relationships } from './package.ts';
import type { Theme } from './theme.ts';

/** What is left of the per-slide caps (PPTX_LIMITS); `truncated` records that something was left out. */
export interface RenderBudget {
  shapes: number;
  runs: number;
  truncated: boolean;
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
