import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import { flattenOutline, renderOutline, type OutlineNode } from '../../../../src/webview/pdf/outline.ts';
import { bootShell, type FakeHost } from '../../dom/env.ts';

function node(title: string, items: OutlineNode[] = [], target: Partial<OutlineNode> = {}): OutlineNode {
  return { title, dest: null, url: null, items, ...target };
}

describe('flattenOutline', () => {
  it('lists entries depth first with their depth', () => {
    const outline = [node('1', [node('1.1', [node('1.1.1')]), node('1.2')]), node('2')];
    assert.deepEqual(
      flattenOutline(outline).map(({ node: { title }, depth }) => `${depth}:${title}`),
      ['0:1', '1:1.1', '2:1.1.1', '1:1.2', '0:2'],
    );
  });

  it('stops at the depth and entry limits of a hostile outline', () => {
    let deep = node('leaf');
    for (let i = 0; i < 10_000; i++) {
      deep = node(`level ${i}`, [deep]);
    }
    assert.equal(flattenOutline([deep], { maxEntries: 5000, maxDepth: 12 }).length, 12);
    const wide = Array.from({ length: 100 }, (_, i) => node(`${i}`, Array.from({ length: 100 }, (__, k) => node(`${i}.${k}`))));
    const entries = flattenOutline(wide, { maxEntries: 5000, maxDepth: 12 });
    assert.equal(entries.length, 5000);
    assert.equal(entries[101]?.node.title, '1');
  });

  it('tolerates nodes without a child list', () => {
    const broken = { title: 'x', dest: null, url: null } as unknown as OutlineNode;
    assert.equal(flattenOutline([broken]).length, 1);
  });
});

describe('renderOutline', () => {
  let host: FakeHost | undefined;
  afterEach(() => host?.close());

  it('renders indented buttons whose titles stay text and reports the picked entry', async () => {
    host = await bootShell({ render() {} });
    const panel = host.window.document.createElement('nav') as unknown as HTMLElement;
    const picked: OutlineNode[] = [];
    const chapter = node('<img src=x onerror=alert(1)>', [node('  ', [], { dest: [{ num: 4, gen: 0 }, { name: 'Fit' }] })], { dest: 'chapter-1' });
    renderOutline(panel, flattenOutline([chapter]), (entry) => picked.push(entry));

    const items = [...panel.querySelectorAll<HTMLButtonElement>('.pdf-outline-item')];
    assert.deepEqual(
      items.map((item) => [item.textContent, item.getAttribute('aria-level'), item.style.getPropertyValue('--depth')]),
      [
        ['<img src=x onerror=alert(1)>', '1', '0'],
        ['Untitled', '2', '1'],
      ],
    );
    assert.equal(panel.querySelector('img'), null);
    items[1]?.click();
    assert.deepEqual(picked, [chapter.items[0]]);
  });
});
