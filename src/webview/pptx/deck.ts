// The presentation as the renderer sees it: slide size and order, and the masters and layouts slides inherit from
// (with their themes, color maps, text styles and placeholders), loaded once and shared by every slide.
import { DEFAULT_COLOR_MAP } from './color.ts';
import { PPTX_LIMITS } from './limits.ts';
import { firstOfType, PptxPackage, type Relationships } from './package.ts';
import { DEFAULT_THEME, parseTheme, type Theme } from './theme.ts';
import { attr, bool, child, find, kids, num } from './xml.ts';

export interface Placeholder {
  readonly element: Element;
  /** Normalized kind used to match placeholders between slide, layout and master. */
  readonly kind: string;
  readonly type: string;
  readonly idx?: string;
}

export interface PartInfo {
  readonly part: string;
  readonly root: Element;
  readonly rels: Relationships;
}

export interface MasterInfo extends PartInfo {
  readonly theme: Theme;
  readonly colorMap: Readonly<Record<string, string>>;
  readonly textStyles: { readonly title?: Element; readonly body?: Element; readonly other?: Element };
  readonly placeholders: readonly Placeholder[];
}

export interface LayoutInfo extends PartInfo {
  readonly master: MasterInfo;
  readonly colorMap: Readonly<Record<string, string>>;
  readonly placeholders: readonly Placeholder[];
  readonly showMasterShapes: boolean;
}

export interface SlideInfo extends PartInfo {
  readonly index: number;
  readonly layout?: LayoutInfo;
  readonly master?: MasterInfo;
  readonly theme: Theme;
  readonly colorMap: Readonly<Record<string, string>>;
  readonly hidden: boolean;
  readonly showMasterShapes: boolean;
}

const DEFAULT_SLIDE_SIZE = { cx: 9144000, cy: 6858000 };
const MIN_SLIDE_SIDE = 914400 / 10;
const MAX_SLIDE_SIDE = 51206400;

export class Deck {
  readonly pkg: PptxPackage;
  /** Slide size in EMUs. */
  readonly width: number;
  readonly height: number;
  readonly slideParts: readonly string[];
  readonly defaultTextStyle?: Element;
  readonly firstSlideNumber: number;
  /** Slides beyond PPTX_LIMITS.maxSlides that are not shown. */
  readonly omittedSlides: number;
  private readonly presentationRels: Relationships;
  private readonly masters = new Map<string, Promise<MasterInfo>>();
  private readonly layouts = new Map<string, Promise<LayoutInfo>>();
  private tableStylesRoot?: Promise<Element | undefined>;

  private constructor(pkg: PptxPackage, presentation: Element, rels: Relationships) {
    this.pkg = pkg;
    this.presentationRels = rels;
    const size = child(presentation, 'sldSz');
    this.width = clampSide(num(size, 'cx') ?? DEFAULT_SLIDE_SIZE.cx);
    this.height = clampSide(num(size, 'cy') ?? DEFAULT_SLIDE_SIZE.cy);
    this.defaultTextStyle = child(presentation, 'defaultTextStyle');
    this.firstSlideNumber = num(presentation, 'firstSlideNum') ?? 1;
    const parts: string[] = [];
    for (const slideId of kids(child(presentation, 'sldIdLst'))) {
      const rel = rels.get(attrRelId(slideId) ?? '');
      if (rel && !rel.external && pkg.has(rel.target)) {
        parts.push(rel.target);
      }
    }
    this.slideParts = parts.slice(0, PPTX_LIMITS.maxSlides);
    this.omittedSlides = parts.length - this.slideParts.length;
  }

  static async open(bytes: Uint8Array): Promise<Deck> {
    const pkg = PptxPackage.open(bytes);
    const main = firstOfType(await pkg.rels(''), 'officeDocument')?.target ?? 'ppt/presentation.xml';
    if (!pkg.has(main)) {
      throw new Error('The presentation part is missing');
    }
    const presentation = await pkg.xml(main, true);
    if (presentation.localName !== 'presentation') {
      throw new Error('The main part is not a presentation');
    }
    return new Deck(pkg, presentation, await pkg.rels(main));
  }

  /** ppt/tableStyles.xml, which defines the table styles tables refer to by id. */
  tableStyles(): Promise<Element | undefined> {
    this.tableStylesRoot ??= (async () => {
      const rel = firstOfType(this.presentationRels, 'tableStyles');
      return rel && this.pkg.has(rel.target) ? this.pkg.xml(rel.target, true) : undefined;
    })();
    return this.tableStylesRoot;
  }

  /** Whether a slide is hidden in slide shows (show="0"), read from its root tag without parsing the slide. */
  async isHidden(index: number): Promise<boolean> {
    const part = this.slideParts[index];
    if (part === undefined) {
      return false;
    }
    const head = (await this.pkg.text(part)).slice(0, 4096);
    const root = /<(?:[\w.-]+:)?sld\b[^>]*>/.exec(head)?.[0] ?? '';
    return /\sshow\s*=\s*["'](?:0|false)["']/.test(root);
  }

  async slide(index: number): Promise<SlideInfo> {
    const part = this.slideParts[index];
    if (part === undefined) {
      throw new Error(`No slide ${index + 1}`);
    }
    const [root, rels] = await Promise.all([this.pkg.xml(part), this.pkg.rels(part)]);
    const layoutRel = firstOfType(rels, 'slideLayout');
    const layout = layoutRel && this.pkg.has(layoutRel.target) ? await this.layout(layoutRel.target) : undefined;
    const master = layout?.master;
    return {
      part,
      root,
      rels,
      index,
      layout,
      master,
      theme: master?.theme ?? DEFAULT_THEME,
      colorMap: colorMapOverride(root) ?? layout?.colorMap ?? master?.colorMap ?? DEFAULT_COLOR_MAP,
      hidden: bool(root, 'show') === false,
      showMasterShapes: bool(root, 'showMasterSp') !== false,
    };
  }

  layout(part: string): Promise<LayoutInfo> {
    let loading = this.layouts.get(part);
    if (!loading) {
      loading = this.loadLayout(part);
      this.layouts.set(part, loading);
    }
    return loading;
  }

  private async loadLayout(part: string): Promise<LayoutInfo> {
    const [root, rels] = await Promise.all([this.pkg.xml(part, true), this.pkg.rels(part)]);
    const masterRel = firstOfType(rels, 'slideMaster');
    if (!masterRel || !this.pkg.has(masterRel.target)) {
      throw new Error(`${part} has no slide master`);
    }
    const master = await this.master(masterRel.target);
    return {
      part,
      root,
      rels,
      master,
      colorMap: colorMapOverride(root) ?? master.colorMap,
      placeholders: placeholdersOf(root),
      showMasterShapes: bool(root, 'showMasterSp') !== false,
    };
  }

  master(part: string): Promise<MasterInfo> {
    let loading = this.masters.get(part);
    if (!loading) {
      loading = this.loadMaster(part);
      this.masters.set(part, loading);
    }
    return loading;
  }

  private async loadMaster(part: string): Promise<MasterInfo> {
    const [root, rels] = await Promise.all([this.pkg.xml(part, true), this.pkg.rels(part)]);
    const themeRel = firstOfType(rels, 'theme');
    const theme = themeRel && this.pkg.has(themeRel.target) ? parseTheme(await this.pkg.xml(themeRel.target, true)) : DEFAULT_THEME;
    const styles = child(root, 'txStyles');
    return {
      part,
      root,
      rels,
      theme,
      colorMap: colorMapOf(child(root, 'clrMap')) ?? DEFAULT_COLOR_MAP,
      textStyles: { title: child(styles, 'titleStyle'), body: child(styles, 'bodyStyle'), other: child(styles, 'otherStyle') },
      placeholders: placeholdersOf(root),
    };
  }
}

function clampSide(value: number): number {
  return Math.max(MIN_SLIDE_SIDE, Math.min(MAX_SLIDE_SIDE, value));
}

function attrRelId(element: Element): string | undefined {
  return element.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || element.getAttribute('r:id') || undefined;
}

function colorMapOf(element: Element | undefined): Record<string, string> | undefined {
  if (!element) {
    return undefined;
  }
  const map: Record<string, string> = { ...DEFAULT_COLOR_MAP };
  for (const key of Object.keys(DEFAULT_COLOR_MAP)) {
    const value = attr(element, key);
    if (value) {
      map[key] = value;
    }
  }
  return map;
}

function colorMapOverride(root: Element): Record<string, string> | undefined {
  return colorMapOf(find(root, 'clrMapOvr', 'overrideClrMapping'));
}

/** title, body, dt, ftr, sldNum, hdr: the kinds placeholders inherit through. */
export function placeholderKind(type: string | undefined): string {
  switch (type) {
    case 'title':
    case 'ctrTitle':
      return 'title';
    case 'dt':
    case 'ftr':
    case 'sldNum':
    case 'hdr':
      return type;
    default:
      return 'body';
  }
}

/** The p:ph of a shape (sp, pic, graphicFrame, grpSp, cxnSp), if it is a placeholder. */
export function placeholderOf(shape: Element): Element | undefined {
  for (const nv of kids(shape)) {
    if (nv.localName.startsWith('nv') && nv.localName.endsWith('Pr')) {
      return find(nv, 'nvPr', 'ph');
    }
  }
  return undefined;
}

function placeholdersOf(root: Element): Placeholder[] {
  const result: Placeholder[] = [];
  for (const shape of kids(find(root, 'cSld', 'spTree'))) {
    const ph = placeholderOf(shape);
    if (ph) {
      const type = attr(ph, 'type') ?? 'obj';
      result.push({ element: shape, type, kind: placeholderKind(type), idx: attr(ph, 'idx') });
    }
  }
  return result;
}

/**
 * The layout or master placeholder a placeholder inherits from: on a layout, the one with the same index (falling back
 * to the same kind); on a master, the one of the same kind.
 */
export function matchPlaceholder(candidates: readonly Placeholder[], ph: Element, byIndex: boolean): Element | undefined {
  const type = attr(ph, 'type') ?? 'obj';
  const kind = placeholderKind(type);
  const idx = attr(ph, 'idx');
  if (byIndex) {
    const sameIdx = candidates.filter((candidate) => (candidate.idx ?? '0') === (idx ?? '0'));
    const exact = sameIdx.find((candidate) => candidate.kind === kind) ?? (idx !== undefined ? sameIdx[0] : undefined);
    if (exact) {
      return exact.element;
    }
  }
  return (candidates.find((candidate) => candidate.type === type) ?? candidates.find((candidate) => candidate.kind === kind))?.element;
}
