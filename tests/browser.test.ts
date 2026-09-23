// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { Editor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { TableKit } from '@tiptap/extension-table';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyleKit } from '@tiptap/extension-text-style';
import { BrowserAdapter } from '../src/document/browser';

let editor: Editor;
let adapter: BrowserAdapter;
beforeEach(() => {
  editor = new Editor({
    element: document.createElement('div'),
    extensions: [
      StarterKit,
      TableKit,
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      TextStyleKit,
    ],
    content: '<h1>标题</h1><p>第一段原文</p><p>第二段原文</p>',
  });
  adapter = new BrowserAdapter(editor, 'test.docx', 'doc');
});
afterEach(() => editor.destroy());
describe('real ProseMirror document adapter', () => {
  it('applies multiple edits against original paragraph positions atomically', async () => {
    const base = await adapter.snapshot();
    const checkpoint = await adapter.apply(
      {
        summary: '修改并插入',
        operations: [
          { type: 'replace', paragraphId: 'p1', expectedText: '第一段原文', text: '第一段已修改' },
          {
            type: 'insert',
            paragraphId: 'p2',
            expectedText: '第二段原文',
            position: 'before',
            text: '插入 A\n插入 B',
          },
        ],
      },
      base,
    );
    expect((await adapter.snapshot()).paragraphs.map((p) => p.text)).toEqual([
      '标题',
      '第一段已修改',
      '插入 A',
      '插入 B',
      '第二段原文',
    ]);
    await adapter.undo(checkpoint);
    expect((await adapter.snapshot()).revision).toBe(base.revision);
  });
  it('preserves untouched rich text and replaces literal HTML as text', async () => {
    editor.commands.setContent('<p><strong>保留格式</strong></p><p>修改目标</p>');
    const base = await adapter.snapshot();
    await adapter.apply(
      {
        summary: '文字',
        operations: [
          {
            type: 'replace',
            paragraphId: 'p1',
            expectedText: '修改目标',
            text: '<img src=x onerror=evil()>',
          },
        ],
      },
      base,
    );
    expect(editor.getHTML()).toContain('<strong>保留格式</strong>');
    expect(editor.getHTML()).toContain('&lt;img');
    expect(editor.getHTML()).not.toContain('<img');
  });
  it('rejects a stale proposal without changing the current document', async () => {
    const base = await adapter.snapshot();
    editor.commands.insertContent('手动输入');
    const current = editor.getJSON();
    await expect(
      adapter.apply(
        {
          summary: '过时',
          operations: [{ type: 'delete', paragraphId: 'p1', expectedText: '第一段原文' }],
        },
        base,
      ),
    ).rejects.toThrow('发生了变化');
    expect(editor.getJSON()).toEqual(current);
  });
  it('protects manual formatting from an old proposal', async () => {
    const base = await adapter.snapshot();
    editor.commands.selectAll();
    editor.commands.toggleBold();
    await expect(
      adapter.apply(
        {
          summary: '过时',
          operations: [{ type: 'delete', paragraphId: 'p1', expectedText: '第一段原文' }],
        },
        base,
      ),
    ).rejects.toThrow('发生了变化');
  });
  it('refuses rollback over later user changes', async () => {
    const base = await adapter.snapshot();
    const checkpoint = await adapter.apply(
      {
        summary: '改写',
        operations: [
          { type: 'replace', paragraphId: 'p1', expectedText: '第一段原文', text: '改写内容' },
        ],
      },
      base,
    );
    editor.commands.insertContent('后续手动修改');
    await expect(adapter.undo(checkpoint)).rejects.toThrow('不能直接撤回');
  });
  it('keeps paragraph identity and style for a replacement', async () => {
    const base = await adapter.snapshot();
    await adapter.apply(
      {
        summary: '标题',
        operations: [{ type: 'replace', paragraphId: 'p0', expectedText: '标题', text: '新标题' }],
      },
      base,
    );
    expect(editor.getJSON().content?.[0].type).toBe('heading');
  });
  it('applies heading, alignment, font size and inline marks', async () => {
    await adapter.apply(
      {
        summary: '格式',
        operations: [
          {
            type: 'format',
            paragraphId: 'p1',
            expectedText: '第一段原文',
            style: 'Heading2',
            bold: true,
            italic: true,
            fontSize: 18,
            alignment: 'center',
          },
        ],
      },
      await adapter.snapshot(),
    );
    const node = editor.getJSON().content![1];
    expect(node.type).toBe('heading');
    expect(node.attrs?.level).toBe(2);
    expect(node.attrs?.textAlign).toBe('center');
    expect(node.content![0].marks?.map((m) => m.type)).toEqual(
      expect.arrayContaining(['bold', 'italic', 'textStyle']),
    );
  });
  it('inserts a rectangular table and protects its cells', async () => {
    await adapter.apply(
      {
        summary: '表格',
        operations: [
          {
            type: 'table',
            paragraphId: 'p2',
            expectedText: '第二段原文',
            rows: [
              ['名称', '说明'],
              ['A', '测试'],
            ],
          },
        ],
      },
      await adapter.snapshot(),
    );
    expect(editor.getJSON().content?.some((n) => n.type === 'table')).toBe(true);
    expect((await adapter.snapshot()).paragraphs.find((p) => p.text === '名称')?.editable).toBe(
      false,
    );
  });
  it('edits the originally selected range even when browser focus moves', async () => {
    editor.commands.setTextSelection({ from: 5, to: 8 });
    const base = await adapter.snapshot();
    editor.commands.setTextSelection(1);
    await adapter.apply(
      {
        summary: '选区',
        operations: [{ type: 'replace_selection', expectedText: base.selection, text: '新' }],
      },
      base,
    );
    expect(editor.getText()).toContain('新');
  });
  it('rejects cross-document edits and rollback', async () => {
    const base = await adapter.snapshot();
    adapter.documentId = 'another';
    await expect(
      adapter.apply(
        {
          summary: '删除',
          operations: [{ type: 'delete', paragraphId: 'p0', expectedText: '标题' }],
        },
        base,
      ),
    ).rejects.toThrow();
  });
  it('preserves a valid empty document after deleting all paragraphs', async () => {
    const base = await adapter.snapshot();
    await adapter.apply(
      {
        summary: '清空',
        operations: base.paragraphs.map((p) => ({
          type: 'delete',
          paragraphId: p.id,
          expectedText: p.text,
        })),
      },
      base,
    );
    expect(editor.getText()).toBe('');
    expect(editor.state.doc.childCount).toBe(1);
  });
});
