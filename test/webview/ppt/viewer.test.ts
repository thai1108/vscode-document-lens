import assert from 'node:assert/strict';
import { afterEach, describe, it, mock } from 'node:test';
import { pptViewer } from '../../../src/webview/ppt/viewer.ts';
import { SAMPLE_PPT } from '../../fixtures/ppt.ts';
import { rgbPng } from '../../support/png.ts';
import { atom, container, pptFile, pptStreams, slidePersistId, u32, writePpt, type PptSpec } from '../../support/pptWriter.ts';
import { bootShell, type FakeHost } from '../dom/env.ts';

let host: FakeHost | undefined;
afterEach(async () => {
  mock.restoreAll();
  await host?.close();
  host = undefined;
});

function deliver(target: FakeHost, data: Uint8Array, fileName = 'deck.ppt'): void {
  target.deliver({ type: 'document', format: 'ppt', fileName, data: new Uint8Array(data) });
}

async function render(spec: PptSpec | Uint8Array): Promise<FakeHost> {
  const target = await bootShell(pptViewer());
  deliver(target, spec instanceof Uint8Array ? spec : writePpt(spec));
  await target.until(() => target.window.document.querySelector('.ppt-deck, #error:not([hidden])') !== null, 'slides or an error');
  return target;
}

const all = (target: FakeHost, selector: string) => [...target.window.document.querySelectorAll(selector)] as unknown as HTMLElement[];
const textOf = (element: Element | undefined) => element?.textContent ?? '';

describe('ppt viewer', () => {
  it('renders one card per slide with its title, text and pictures', async () => {
    host = await render(SAMPLE_PPT);
    const slides = all(host, '.ppt-slide');
    assert.deepEqual(
      slides.map((slide) => slide.id),
      ['slide-1', 'slide-2', 'slide-3'],
    );
    assert.deepEqual(
      slides.map((slide) => textOf(slide.querySelector('.ppt-slide-number') ?? undefined)),
      ['Slide 1', 'Slide 2', 'Slide 3'],
    );
    assert.equal(textOf(host.$('#slide-1 h2.ppt-title')), 'Quarterly review');
    assert.deepEqual(
      all(host, '#slide-1 .ppt-text p').map((paragraph) => paragraph.textContent),
      ['四半期レビュー', 'Document Lens sample'],
    );

    const bullets = all(host, '#slide-2 .ppt-text-body ul > li');
    assert.deepEqual(
      bullets.map((item) => [item.textContent, item.className]),
      [
        ['Revenue grew 12%', 'ppt-level-0'],
        ['New customers in Osaka', 'ppt-level-1'],
        ['Roadmap\nsecond line', 'ppt-level-0'],
        ['東京オフィス開設', 'ppt-level-1'],
      ],
    );
    assert.equal(textOf(host.$('#slide-2 .ppt-text-other p')), '注: 速報値です');
    assert.equal(host.window.document.body.textContent?.includes('Confidential'), false, 'footer placeholders are left out');

    const image = host.$<HTMLImageElement>('#slide-2 .ppt-picture img');
    assert.match(image.getAttribute('src') ?? '', /^blob:/);
    assert.equal(image.getAttribute('alt'), 'Picture 1 on Slide 2');
    assert.deepEqual([image.width, image.height], [120, 90]);
    assert.equal(host.$('#ppt-deck').style.getPropertyValue('--ppt-aspect'), '5760 / 3240');
  });

  it('marks hidden slides', async () => {
    host = await render(SAMPLE_PPT);
    const hidden = host.$('#slide-3');
    assert.ok(hidden.classList.contains('ppt-slide-hidden'));
    assert.equal(hidden.getAttribute('aria-label'), 'Slide 3 (hidden)');
    assert.equal(textOf(hidden.querySelector('.ppt-badge') ?? undefined), 'Hidden');
    assert.equal(host.$('#slide-2').querySelector('.ppt-badge'), null);
    assert.equal(textOf(host.$('#ppt-slide-count')), '3 slides · 1 hidden');
  });

  it('shows the simplified preview notice', async () => {
    host = await render(SAMPLE_PPT);
    assert.equal(host.$('#notice').hidden, false);
    assert.equal(textOf(host.$('#notice')), 'Simplified preview: only the text and pictures of each slide are shown, without layout, shapes or formatting.');
  });

  it('adds what was left out to the notice and labels pictures it cannot show', async () => {
    let nested: Buffer = Buffer.alloc(0);
    for (let i = 0; i < 60; i++) {
      nested = container(0xf003, [container(0xf004, [atom(0xf00a, u32(0, 0x201), 0, 2)]), nested]);
    }
    host = await render({
      pictures: [{ format: 'emf', data: Buffer.alloc(40) }],
      slides: [{ pictures: [1], extraShapes: [nested] }, {}],
    });
    assert.match(textOf(host.$('#notice')), /formatting\. Some deeply nested shapes are left out\.$/);
    const placeholder = host.$('#slide-1 .ppt-picture');
    assert.ok(placeholder.classList.contains('ppt-picture-placeholder'));
    assert.equal(textOf(placeholder), 'EMF picture (not shown)');
    assert.equal(textOf(host.$('#slide-2 .ppt-note')), 'No text or pictures on this slide.');
    assert.equal(textOf(host.$('#ppt-slide-count')), '2 slides');
  });

  it('says when a presentation has no slides', async () => {
    host = await render({ slides: [] });
    assert.equal(textOf(host.$('.ppt-no-slides')), 'This presentation has no slides.');
    assert.equal(textOf(host.$('#ppt-slide-count')), '0 slides');
  });

  it('says when part of a slide could not be read', async () => {
    const streams = pptStreams(SAMPLE_PPT);
    // Point slide 1's persist entry at the DocumentContainer (offset 0), which is not a slide.
    const directory = streams.document.readUInt32LE((streams.edits[0] ?? 0) + 8 + 12);
    streams.document.writeUInt32LE(0, directory + 8 + 4 + (slidePersistId(0) - 1) * 4);
    host = await render(new Uint8Array(pptFile({ 'PowerPoint Document': streams.document, 'Current User': streams.currentUser, Pictures: streams.pictures })));
    assert.equal(textOf(host.$('#slide-1 .ppt-note')), 'Part of this slide could not be read.');
    assert.equal(textOf(host.$('#slide-1 h2')), 'Quarterly review');
    assert.equal(host.$('#slide-2').querySelector('.ppt-note'), null);
  });

  it('puts file text into the page as text, never as markup', async () => {
    host = await render({ slides: [{ placeholders: [{ type: 'title', text: '<img src=x onerror=alert(1)>' }, { type: 'body', text: '<b>bold</b>' }] }] });
    assert.equal(host.window.document.querySelectorAll('#content img, #content b').length, 0);
    assert.equal(textOf(host.$('#slide-1 h2')), '<img src=x onerror=alert(1)>');
  });

  it('shows the error banner for password-protected and damaged files', async () => {
    host = await render({ ...SAMPLE_PPT, encrypted: { token: true } });
    assert.equal(host.$('#error').hidden, false);
    assert.equal(textOf(host.$('#error')), 'Could not display deck.ppt: This .ppt file is password-protected; Document Lens cannot open encrypted files');
    await host.close();

    const broken = new Uint8Array(writePpt(SAMPLE_PPT));
    broken.fill(0xff, 600, broken.length);
    host = await render(broken);
    assert.match(textOf(host.$('#error')), /^Could not display deck\.ppt: (This \.ppt file is damaged|Not a \.ppt file)/);
  });

  it('re-renders on reload, keeps the scroll position and releases the old picture URLs', async () => {
    const revoked: string[] = [];
    mock.method(URL, 'revokeObjectURL', (url: string) => revoked.push(url));
    host = await render(SAMPLE_PPT);
    const firstUrl = host.$('#slide-2 img').getAttribute('src');
    host.$('#content').scrollTop = 120;

    const updated: PptSpec = { ...SAMPLE_PPT, slides: [...SAMPLE_PPT.slides, { placeholders: [{ type: 'title', text: 'Added' }] }] };
    deliver(host, writePpt(updated));
    await host.until(() => host?.window.document.getElementById('slide-4') !== null, 'the added slide');
    assert.deepEqual(revoked, [firstUrl]);
    assert.notEqual(host.$('#slide-2 img').getAttribute('src'), firstUrl);
    assert.equal(host.$('#content').scrollTop, 120);
    assert.equal(textOf(host.$('#ppt-slide-count')), '4 slides · 1 hidden');
  });

  it('keeps the previous slides when a reload fails', async () => {
    host = await render(SAMPLE_PPT);
    deliver(host, writePpt({ ...SAMPLE_PPT, encrypted: { token: true } }));
    await host.until(() => host?.$('#error').hidden === false, 'the error');
    assert.equal(all(host, '.ppt-slide').length, 3);
  });

  it('creates one URL per picture even when several slides show it', async () => {
    const created: unknown[] = [];
    const original = URL.createObjectURL.bind(URL);
    mock.method(URL, 'createObjectURL', (blob: Blob) => {
      created.push(blob);
      return original(blob);
    });
    host = await render({ pictures: [{ format: 'png', data: rgbPng(2, 2, () => [0, 0, 0]) }], slides: [{ pictures: [1] }, { pictures: [1, 1] }] });
    assert.equal(all(host, '.ppt-picture img').length, 3);
    assert.equal(created.length, 1);
    assert.equal((created[0] as Blob).type, 'image/png');
  });
});
