import { beforeEach, afterEach, describe, it, expect, vi } from 'vitest';
import { WordAdapter } from '../src/document/word';

function officeFixture() {
  let items: MockParagraph[] = [];
  let selectionIndex = 0;
  const actions: (() => void)[] = [];
  const queue = (fn: () => void) => actions.push(fn);
  class MockRange {
    text = '';
    tracked = false;
    constructor(public paragraph: MockParagraph) {}
    load() {
      this.text = this.paragraph.text;
      return this;
    }
    track() {
      this.tracked = true;
      return this;
    }
    untrack() {
      this.tracked = false;
      return this;
    }
    insertText(text: string) {
      queue(() => {
        this.paragraph.text = text;
      });
    }
    clear() {
      queue(() => {
        this.paragraph.text = '';
      });
    }
  }
  class MockParagraph {
    styleBuiltIn = 'Normal';
    alignment = 'Left';
    font = { bold: false, italic: false, size: 12, name: '等线' };
    parentTableOrNullObject = { isNullObject: true, load() {} };
    constructor(public text: string) {}
    getRange() {
      return new MockRange(this);
    }
    delete() {
      queue(() => {
        items.splice(items.indexOf(this), 1);
      });
    }
    select() {
      selectionIndex = items.indexOf(this);
    }
    insertParagraph(text: string, position: string) {
      const p = new MockParagraph(text);
      queue(() => {
        items.splice(items.indexOf(this) + (position === 'After' ? 1 : 0), 0, p);
      });
      return p;
    }
  }
  items = [
    new MockParagraph('同名文字'),
    new MockParagraph('同名文字'),
    new MockParagraph('第三段'),
  ];
  const serialize = () =>
    JSON.stringify(
      items.map((p) => ({
        text: p.text,
        style: p.styleBuiltIn,
        alignment: p.alignment,
        font: p.font,
      })),
    );
  const body = {
    get paragraphs() {
      return { items: [...items], load() {} };
    },
    getOoxml() {
      const result = { value: '' };
      queue(() => {
        result.value = serialize();
      });
      return result;
    },
    insertOoxml(xml: string) {
      queue(() => {
        items = JSON.parse(xml).map(
          (v: { text: string; style: string; alignment: string; font: MockParagraph['font'] }) =>
            Object.assign(new MockParagraph(v.text), {
              styleBuiltIn: v.style,
              alignment: v.alignment,
              font: v.font,
            }),
        );
      });
    },
  };
  const context = {
    document: { body, getSelection: () => new MockRange(items[selectionIndex]) },
    sync: async () => {
      while (actions.length) actions.shift()!();
    },
  };
  vi.stubGlobal('Word', {
    run: async (...args: unknown[]) =>
      (args.at(-1) as (ctx: typeof context) => Promise<unknown>)(context),
  });
  vi.stubGlobal('Office', {
    context: {
      document: {
        url: 'https://local/document.docx',
        settings: { get: () => 'doc', set() {}, saveAsync() {} },
      },
    },
  });
  return {
    items: () => items,
    select: (index: number) => {
      selectionIndex = index;
    },
  };
}
let fixture: ReturnType<typeof officeFixture>;
beforeEach(() => {
  fixture = officeFixture();
});
afterEach(() => vi.unstubAllGlobals());
describe('Word adapter contract with mocked Office host', () => {
  it('captures Word paragraph metadata and OOXML revision', async () => {
    const adapter = new WordAdapter();
    const snapshot = await adapter.snapshot();
    expect(snapshot.title).toBe('document.docx');
    expect(snapshot.paragraphs).toHaveLength(3);
    expect(snapshot.paragraphs[0].editable).toBe(true);
  });
  it('uses the captured selection, not a different selection with identical text', async () => {
    const adapter = new WordAdapter();
    const base = await adapter.capture();
    fixture.select(1);
    const checkpoint = await adapter.apply(
      {
        summary: '选区',
        operations: [{ type: 'replace_selection', expectedText: '同名文字', text: '原始选区已改' }],
      },
      base,
    );
    expect(fixture.items().map((p) => p.text)).toEqual(['原始选区已改', '同名文字', '第三段']);
    await adapter.undo(checkpoint);
    expect(fixture.items()[0].text).toBe('同名文字');
  });
  it('refuses a reloaded Word selection without its tracked Range', async () => {
    const adapter = new WordAdapter();
    const base = await adapter.snapshot();
    await expect(
      adapter.apply(
        {
          summary: '选区',
          operations: [{ type: 'replace_selection', expectedText: '同名文字', text: 'new' }],
        },
        base,
      ),
    ).rejects.toThrow('定位已失效');
  });
  it('does not overwrite Word manual edits', async () => {
    const adapter = new WordAdapter();
    const base = await adapter.capture();
    fixture.items()[2].text = '用户更新';
    await expect(
      adapter.apply(
        {
          summary: '替换',
          operations: [
            { type: 'replace', paragraphId: 'p0', expectedText: '同名文字', text: '新' },
          ],
        },
        base,
      ),
    ).rejects.toThrow('发生了变化');
  });
  it('inserts multiple paragraphs in the correct order before a target', async () => {
    const adapter = new WordAdapter();
    const base = await adapter.capture();
    await adapter.apply(
      {
        summary: '插入',
        operations: [
          {
            type: 'insert',
            paragraphId: 'p2',
            expectedText: '第三段',
            position: 'before',
            text: 'A\nB',
          },
        ],
      },
      base,
    );
    expect(fixture.items().map((p) => p.text)).toEqual([
      '同名文字',
      '同名文字',
      'A',
      'B',
      '第三段',
    ]);
  });
  it('protects later formatting during Word rollback', async () => {
    const adapter = new WordAdapter();
    const base = await adapter.capture();
    const checkpoint = await adapter.apply(
      {
        summary: '替换',
        operations: [{ type: 'replace', paragraphId: 'p0', expectedText: '同名文字', text: '新' }],
      },
      base,
    );
    fixture.items()[2].font.bold = true;
    await expect(adapter.undo(checkpoint)).rejects.toThrow('不能直接撤回');
  });
});
