// The presentation viewer inside the shared shell, with a fake VS Code host (test/webview/dom/env.ts).
import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import type { HostMessage } from '../../../src/protocol.ts';
import { PPTX_LIMITS } from '../../../src/webview/pptx/limits.ts';
import { createPptxViewer } from '../../../src/webview/pptx/viewer.ts';
import { sampleDeck } from '../../fixtures/pptx.ts';
import { buildPptx, paragraph, run, shape } from '../../support/pptxBuilder.ts';
import { writeZip } from '../../support/zipWriter.ts';
import { bootShell, type FakeHost } from '../dom/env.ts';
import { conformingParser } from './dom.ts';

let host: FakeHost | undefined;
afterEach(async () => {
  await host?.close();
  host = undefined;
  delete (globalThis as { IntersectionObserver?: unknown }).IntersectionObserver;
});

const pptx = (data: Uint8Array, fileName = 'deck.pptx'): HostMessage => ({ type: 'document', format: 'pptx', fileName, data: new Uint8Array(data) });

/** Reports every observed thumbnail as visible at once, as a tall window would. */
class VisibleObserver {
  private readonly callback: (entries: { target: Element; isIntersecting: boolean }[]) => void;
  constructor(callback: (entries: { target: Element; isIntersecting: boolean }[]) => void) {
    this.callback = callback;
  }
  observe(target: Element) {
    queueMicrotask(() => this.callback([{ target, isIntersecting: true }]));
  }
  unobserve() {}
  disconnect() {}
}

type ObserverCallback = (entries: { target: Element; isIntersecting: boolean }[]) => void;

/** An IntersectionObserver the test drives: report() says which thumbnails enter or leave its area. */
class ManualObserver {
  static readonly created: ManualObserver[] = [];
  readonly margin: string | undefined;
  private readonly callback: ObserverCallback;
  private readonly targets: Element[] = [];
  constructor(callback: ObserverCallback, options: { rootMargin?: string } = {}) {
    this.callback = callback;
    this.margin = options.rootMargin;
    ManualObserver.created.push(this);
  }
  observe(target: Element) {
    this.targets.push(target);
  }
  unobserve() {}
  disconnect() {}
  report(indices: number[], isIntersecting: boolean) {
    this.callback(this.targets.filter((target) => indices.includes(Number(target.getAttribute('data-index')))).map((target) => ({ target, isIntersecting })));
  }
}

async function open(data: Uint8Array, options: Parameters<typeof bootShell>[1] = {}): Promise<FakeHost> {
  host = await bootShell(createPptxViewer(), options);
  (globalThis as { DOMParser?: unknown }).DOMParser = conformingParser(host.window);
  host.deliver(pptx(data));
  const document = host.window.document;
  await host.until(() => document.getElementById('pptx-current-slide') !== null || document.querySelector('.pptx-empty') !== null || host?.$('#error').hidden === false, 'first slide');
  return host;
}

const counter = () => host?.$('#pptx-counter').textContent;
const currentText = () => host?.$('#pptx-current-slide').textContent ?? '';

describe('presentation viewer', () => {
  it('lists the slides and shows the first one with its number in the toolbar', async () => {
    host = await open(sampleDeck());
    assert.equal(counter(), '1 / 4');
    const thumbs = [...host.window.document.querySelectorAll('.pptx-thumb')];
    assert.deepEqual(thumbs.map((thumb) => thumb.textContent), ['1', '2', '3', '4']);
    assert.equal(thumbs[0]?.getAttribute('aria-current'), 'true');
    assert.ok(currentText().includes('Document Lens サンプル'));
    assert.equal(host.$<HTMLButtonElement>('#pptx-previous').disabled, true);
    assert.equal(host.$<HTMLButtonElement>('#pptx-next').disabled, false);
    assert.equal(host.$('#notice').hidden, true);
    assert.equal(host.$('#fit-width').hidden, false, 'the shell offers fitting the slide');
  });

  it('moves between slides with the toolbar, the list and the keyboard, and remembers the slide', async () => {
    host = await open(sampleDeck());
    host.click(host.$('#pptx-next'));
    await host.until(() => counter() === '2 / 4', 'next');
    await host.until(() => currentText().includes('Shapes and colors'), 'slide 2');
    host.key('ArrowRight');
    await host.until(() => counter() === '3 / 4', 'arrow right');
    host.key('End');
    await host.until(() => counter() === '4 / 4', 'end');
    await host.until(() => host?.$('#pptx-hidden-tag').hidden === false, 'hidden tag');
    host.key('Home');
    await host.until(() => counter() === '1 / 4', 'home');
    host.key('ArrowLeft');
    host.key('PageDown');
    await host.until(() => counter() === '2 / 4', 'page down');
    host.key('PageUp');
    await host.until(() => counter() === '1 / 4', 'page up');
    host.key('ArrowRight', { metaKey: true });
    host.click(host.window.document.querySelectorAll('.pptx-thumb')[2] as unknown as Element);
    await host.until(() => counter() === '3 / 4', 'thumbnail click');
    // happy-dom has no layout, so the slide could not be fitted yet; it stays in fit mode until it can.
    assert.deepEqual(host.state?.viewer, { slide: 2, fit: true });
    host.click(host.$('#zoom-in'));
    assert.deepEqual(host.state?.viewer, { slide: 2, fit: false }, 'zooming leaves fit mode');
  });

  it('shows the slide chosen last when several are still rendering', async () => {
    host = await open(sampleDeck());
    host.key('ArrowRight');
    host.key('ArrowRight');
    host.key('ArrowRight');
    await host.until(() => currentText().includes('Hidden slide'), 'last slide');
    await new Promise((resolve) => setTimeout(resolve, 50));
    assert.equal(counter(), '4 / 4');
    assert.ok(currentText().includes('Hidden slide'));
    assert.equal(host.$('#pptx-hidden-tag').hidden, false, 'an earlier slide finishing late does not clear the hidden tag');
  });

  it('restores the saved slide and keeps it when the file changes on disk, releasing old pictures', async () => {
    host = await open(sampleDeck(), { state: { zoom: 1.5, viewer: { slide: 2, fit: false } } });
    assert.equal(counter(), '3 / 4');
    assert.equal(host.$('#zoom-value').textContent, '150%');
    const revoked: string[] = [];
    const revoke = URL.revokeObjectURL;
    URL.revokeObjectURL = (url: string) => revoked.push(url);
    try {
      const before = host.$('#pptx-current-slide');
      host.deliver(pptx(sampleDeck()));
      await host.until(() => (host?.window.document.getElementById('pptx-current-slide') as unknown) !== before, 'reloaded slide');
      await host.until(() => revoked.length > 0, 'old picture URLs revoked');
    } finally {
      URL.revokeObjectURL = revoke;
    }
    assert.equal(counter(), '3 / 4');
    assert.equal(host.$('#zoom-value').textContent, '150%');
  });

  it('opens web links through the host and follows links to other slides', async () => {
    const jump = (rId: string, action: string, text: string) => run(text, `<a:rPr lang="en-US"><a:hlinkClick r:id="${rId}" action="${action}"/></a:rPr>`);
    const bytes = buildPptx({
      slides: [
        {
          layout: 2,
          shapes: shape({ id: 2, name: 'Links', xfrm: [0, 0, 5000000, 1000000], geometry: 'rect', paragraphs: paragraph(jump('rId2', 'ppaction://hlinksldjump', 'third') + jump('', 'ppaction://hlinkshowjump?jump=nextslide', 'next')) }),
          rels: [{ id: 'rId2', type: 'slide', target: 'slide3.xml' }],
        },
        { layout: 2, shapes: '' },
        { layout: 2, shapes: shape({ id: 2, name: 'Web', xfrm: [0, 0, 5000000, 1000000], geometry: 'rect', paragraphs: paragraph(run('site', '<a:rPr lang="en-US"><a:hlinkClick r:id="rId2"/></a:rPr>')) }), rels: [{ id: 'rId2', type: 'hyperlink', target: 'https://example.com/docs', external: true }] },
      ],
    });
    host = await open(bytes);
    const links = [...host.window.document.querySelectorAll('#pptx-current-slide a')];
    assert.deepEqual(links.map((link) => [link.getAttribute('href'), link.getAttribute('data-slide-index'), link.getAttribute('data-slide-jump')]), [
      [null, '2', null],
      [null, null, 'nextslide'],
    ]);
    host.click(links[1]?.querySelector('span') as unknown as Element);
    await host.until(() => counter() === '2 / 3', 'next slide link');
    host.key('ArrowLeft');
    await host.until(() => counter() === '1 / 3', 'back');
    await host.until(() => host?.window.document.querySelector('#pptx-current-slide a') !== null, 'links again');
    host.click(host.window.document.querySelector('#pptx-current-slide a span') as unknown as Element);
    await host.until(() => counter() === '3 / 3', 'slide jump');
    await host.until(() => currentText().includes('site'), 'third slide');
    host.click(host.$('#pptx-current-slide a span'));
    assert.deepEqual(host.posted.at(-1), { type: 'openLink', url: 'https://example.com/docs' });
  });

  it('draws thumbnails as they come into view and marks hidden slides', async () => {
    (globalThis as { IntersectionObserver?: unknown }).IntersectionObserver = VisibleObserver;
    host = await open(sampleDeck());
    await host.until(() => host?.window.document.querySelectorAll('.pptx-thumb-frame .pptx-slide').length === 4, 'thumbnails');
    const thumbs = [...host.window.document.querySelectorAll('.pptx-thumb')];
    assert.deepEqual(thumbs.map((thumb) => thumb.classList.contains('pptx-thumb-hidden')), [false, false, false, true]);
    assert.equal(thumbs[3]?.getAttribute('aria-label'), 'Slide 4 (hidden)');
    assert.equal(thumbs[0]?.querySelector('.pptx-thumb-frame')?.hasAttribute('inert'), true, 'thumbnail links cannot be followed');
    assert.equal(host.window.document.querySelectorAll('[id="pptx-current-slide"]').length, 1);
  });

  it('releases thumbnails far out of view and keeps none for slides that hit a limit', async () => {
    (globalThis as { IntersectionObserver?: unknown }).IntersectionObserver = ManualObserver;
    // Slide 2 of the sample has more shapes than this; the others fit.
    PPTX_LIMITS.maxShapesPerSlide = 6;
    try {
      host = await open(sampleDeck());
      const [draw, keep] = ManualObserver.created.slice(-2);
      assert.ok(draw && keep);
      assert.deepEqual([draw.margin, keep.margin], ['300px', '1500px']);
      const frames = () => [...(host?.window.document.querySelectorAll('.pptx-thumb-frame') ?? [])];
      draw.report([0, 1, 2, 3], true);
      await host.until(() => frames().every((frame) => frame.childElementCount > 0), 'thumbnails');
      assert.deepEqual(frames().map((frame) => (frame.querySelector('.pptx-slide') ? 'slide' : frame.textContent)), ['slide', 'Too large to preview', 'slide', 'slide']);

      keep.report([2], false);
      assert.equal(frames()[2]?.childElementCount, 0, 'released');
      draw.report([2], true);
      await host.until(() => frames()[2]?.querySelector('.pptx-slide') !== null, 'drawn again');

      host.key('ArrowRight');
      await host.until(() => host?.$('#notice').textContent === 'Some content on this slide could not be shown.', 'notice');
    } finally {
      PPTX_LIMITS.maxShapesPerSlide = 5_000;
    }
  });

  it('reports a file that is not a presentation in the error banner', async () => {
    host = await open(new Uint8Array(writeZip([{ name: '[Content_Types].xml', data: '<Types/>' }])));
    assert.equal(host.$('#error').textContent, 'Could not display deck.pptx: The presentation part is missing');
  });

  it('says so when there are no slides, and when a slide cannot be shown', async () => {
    host = await open(buildPptx({ slides: [] }));
    assert.equal(counter(), '0 / 0');
    assert.equal(host.$('#pptx-frame').textContent, 'This presentation has no slides.');
    await host.close();

    const broken = buildPptx({
      slides: [
        { layout: 2, shapes: '' },
        { layout: 2, shapes: '', xml: '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>' },
      ],
    });
    host = await open(broken);
    host.key('ArrowRight');
    await host.until(() => counter() === '2 / 2', 'broken slide');
    await host.until(() => currentText().startsWith('Slide 2 could not be displayed:'), 'error card');
    assert.match(currentText(), /ppt\/slides\/slide2\.xml is not well-formed XML/);
    assert.equal(host.$('#error').hidden, true, 'the rest of the deck still works');
  });
});
