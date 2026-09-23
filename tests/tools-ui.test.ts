// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { renderToStaticMarkup } from 'react-dom/server';
import { ToolsPanel } from '../src/components/ToolsPanel';
import { ChangeCard } from '../src/components/ChangeCard';
import type { Snapshot } from '../shared/contracts';

const snapshot: Snapshot = {
  title: 'test',
  documentId: 'doc',
  revision: 'rev',
  selection: '',
  selectionKey: '',
  paragraphs: [{ id: 'p0', text: 'old old', style: 'Normal', editable: true }],
};
let element: HTMLDivElement;
let root: Root;
let props: Parameters<typeof ToolsPanel>[0];
beforeEach(async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  vi.stubGlobal('fetch', vi.fn());
  element = document.createElement('div');
  document.body.append(element);
  root = createRoot(element);
  props = {
    snapshot,
    busy: false,
    onRefresh: vi.fn(),
    onFocus: vi.fn(),
    onStage: vi.fn(),
    onRecipe: vi.fn(),
  };
  await act(async () => {
    root.render(createElement(ToolsPanel, props));
  });
});
afterEach(async () => {
  await act(async () => {
    root.unmount();
  });
  element.remove();
  vi.unstubAllGlobals();
});
async function fill(id: string, value: string) {
  const input = element.querySelector<HTMLInputElement>(id)!;
  await act(async () => {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
}
describe('document tools UI', () => {
  it('searches and stages a preview locally without model requests or document writes', async () => {
    await fill('#tool-find', 'old');
    await fill('#tool-replacement', 'new');
    expect(element.textContent).toContain('找到 2 处匹配');
    expect(element.querySelector<HTMLButtonElement>('button[type="submit"]')!.disabled).toBe(false);
    await act(async () => {
      element
        .querySelector('form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    });
    expect(props.onStage).toHaveBeenCalledWith(
      expect.objectContaining({
        operations: [
          expect.objectContaining({ type: 'replace_text', expectedMatches: 2, text: 'new' }),
        ],
      }),
      snapshot,
    );
    expect(snapshot.paragraphs[0].text).toBe('old old');
    expect(fetch).not.toHaveBeenCalled();
  });
  it('navigates to a search result without requesting AI', async () => {
    await fill('#tool-find', 'old');
    await act(async () => {
      element.querySelector<HTMLButtonElement>('.tool-hit')!.click();
    });
    expect(props.onFocus).toHaveBeenCalledWith('p0');
    expect(fetch).not.toHaveBeenCalled();
  });
  it('prepares a translation prompt but does not send it', async () => {
    const button = [...element.querySelectorAll<HTMLButtonElement>('.tool-recipes button')].find(
      (b) => b.textContent?.includes('中译英'),
    )!;
    await act(async () => {
      button.click();
    });
    expect(props.onRecipe).toHaveBeenCalledWith(
      expect.objectContaining({ id: 'translate', mode: 'agent' }),
    );
    expect(fetch).not.toHaveBeenCalled();
  });
  it('keeps preview actions disabled while another action is running', async () => {
    await fill('#tool-find', 'old');
    await fill('#tool-replacement', 'new');
    await act(async () => {
      root.render(createElement(ToolsPanel, { ...props, busy: true }));
    });
    expect(element.querySelector<HTMLButtonElement>('button[type="submit"]')!.disabled).toBe(true);
    expect(
      [...element.querySelectorAll<HTMLButtonElement>('.tool-recipes button')].every(
        (b) => b.disabled,
      ),
    ).toBe(true);
  });
  it('reports protected matches instead of silently skipping them', async () => {
    await act(async () => {
      root.render(
        createElement(ToolsPanel, {
          ...props,
          snapshot: { ...snapshot, paragraphs: [{ ...snapshot.paragraphs[0], editable: false }] },
        }),
      );
    });
    await fill('#tool-find', 'old');
    await fill('#tool-replacement', 'new');
    await act(async () => {
      element
        .querySelector('form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    });
    expect(element.querySelector('[role="alert"]')?.textContent).toContain('受保护');
    expect(props.onStage).not.toHaveBeenCalled();
  });
  it('renders the complete expected before/after text for a phrase replacement preview', () => {
    const html = renderToStaticMarkup(
      createElement(ChangeCard, {
        change: {
          id: 'change',
          created: 0,
          status: 'pending',
          base: snapshot,
          plan: {
            summary: 'replace',
            operations: [
              {
                type: 'replace_text',
                paragraphId: 'p0',
                expectedText: 'old old',
                find: 'old',
                text: 'new',
                expectedMatches: 2,
                occurrence: 2,
              },
            ],
          },
        },
        busy: false,
        canUndo: false,
        onApply: vi.fn(),
        onReject: vi.fn(),
        onUndo: vi.fn(),
      }),
    );
    expect(html).toContain('精确替换文字');
    expect(html).toContain('<span>old </span>');
    expect(html).toContain('<ins>new</ins>');
    expect(html).toContain('应用更改');
  });
});
