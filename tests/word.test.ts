import { beforeEach, afterEach, describe, it, expect, vi } from 'vitest';
import { literalOffsets } from '../shared/text-edits';
import { WordAdapter } from '../src/document/word';

function officeFixture() {
  let items: MockParagraph[] = [];
  let selectionIndex = 0;
  let badSearch = false;
  let packageSuffix = '';
  const searches: { query: string; options: unknown }[] = [];
  const actions: (() => void)[] = [];
  const queue = (fn: () => void) => actions.push(fn);
  class MockRange {
    text = '';
    tracked = false;
    constructor(
      public paragraph: MockParagraph,
      public from = 0,
      public to?: number,
    ) {}
    load() {
      this.text = this.paragraph.text.slice(this.from, this.to);
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
        this.paragraph.text =
          this.to === undefined
            ? text
            : this.paragraph.text.slice(0, this.from) + text + this.paragraph.text.slice(this.to);
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
    search(query: string, options: unknown) {
      searches.push({ query, options });
      const collection = {
        items: [] as MockRange[],
        load: () => {
          queue(() => {
            const find = query.replace(/\^\^/g, '^');
            collection.items = badSearch
              ? []
              : literalOffsets(this.text, find).map((from) => {
                  const range = new MockRange(this, from, from + find.length);
                  range.load();
                  return range;
                });
          });
        },
      };
      return collection;
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
        result.value = serialize() + packageSuffix;
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
    searches,
    mismatch: () => {
      badSearch = true;
    },
    changePackage: () => {
      packageSuffix = ' ';
    },
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
  it('replaces exact Word ranges rather than the whole paragraph and supports undo', async () => {
    fixture.items()[0].text = 'old ^p old';
    const adapter = new WordAdapter();
    const base = await adapter.capture();
    const checkpoint = await adapter.apply(
      {
        summary: '短语',
        operations: [
          {
            type: 'replace_text',
            paragraphId: 'p0',
            expectedText: 'old ^p old',
            find: 'old',
            text: 'new longer',
            expectedMatches: 2,
            occurrence: 2,
          },
        ],
      },
      base,
    );
    expect(fixture.items()[0].text).toBe('old ^p new longer');
    expect(fixture.searches[0].options).toMatchObject({
      matchWildcards: false,
      matchCase: true,
      ignorePunct: false,
      ignoreSpace: false,
    });
    await adapter.undo(checkpoint);
    expect(fixture.items()[0].text).toBe('old ^p old');
  });
  it('escapes caret codes and replaces all matches from bottom to top', async () => {
    fixture.items()[0].text = '^p ^p';
    const adapter = new WordAdapter();
    const base = await adapter.capture();
    await adapter.apply(
      {
        summary: '字面查找',
        operations: [
          {
            type: 'replace_text',
            paragraphId: 'p0',
            expectedText: '^p ^p',
            find: '^p',
            text: 'X',
            expectedMatches: 2,
          },
        ],
      },
      base,
    );
    expect(fixture.searches[0].query).toBe('^^p');
    expect(fixture.items()[0].text).toBe('X X');
  });
  it('refuses the entire batch before writes if Word search disagrees', async () => {
    const adapter = new WordAdapter();
    const base = await adapter.capture();
    fixture.mismatch();
    await expect(
      adapter.apply(
        {
          summary: '错误搜索',
          operations: [
            {
              type: 'replace_text',
              paragraphId: 'p0',
              expectedText: '同名文字',
              find: '同名',
              text: '新',
              expectedMatches: 1,
            },
            { type: 'delete', paragraphId: 'p2', expectedText: '第三段' },
          ],
        },
        base,
      ),
    ).rejects.toThrow('匹配结果');
    expect(fixture.items().map((p) => p.text)).toEqual(['同名文字', '同名文字', '第三段']);
  });
  it('rejects an OOXML-only change even when paragraph metadata is identical', async () => {
    const adapter = new WordAdapter();
    const base = await adapter.capture();
    fixture.changePackage();
    expect((await adapter.snapshot()).paragraphs).toEqual(base.paragraphs);
    await expect(
      adapter.apply(
        {
          summary: '旧方案',
          operations: [{ type: 'delete', paragraphId: 'p2', expectedText: '第三段' }],
        },
        base,
      ),
    ).rejects.toThrow('发生了变化');
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
