import { describe, expect, it } from 'vitest';
import {
  chatSchema,
  planSchema,
  profileSchema,
  revision,
  validatePlan,
  type EditPlan,
  type Snapshot,
} from '../shared/contracts';

const snapshot: Snapshot = {
  documentId: 'doc',
  title: 'test.docx',
  selection: '选中文本',
  selectionKey: '1:5',
  revision: 'v1',
  paragraphs: [
    { id: 'p0', text: '原始文字', style: 'Normal', editable: true },
    { id: 'p1', text: '表格内容', style: 'Normal', editable: false },
  ],
};
const plan: EditPlan = {
  summary: '润色',
  operations: [{ type: 'replace', paragraphId: 'p0', expectedText: '原始文字', text: '新的文字' }],
};
describe('edit contract', () => {
  it('accepts an exact, scoped replacement', () =>
    expect(() => validatePlan(plan, snapshot)).not.toThrow());
  it('rejects stale text', () =>
    expect(() =>
      validatePlan(plan, {
        ...snapshot,
        paragraphs: [{ ...snapshot.paragraphs[0], text: '手动修改' }],
      }),
    ).toThrow('原文不匹配'));
  it('rejects unknown paragraph ids', () =>
    expect(() =>
      validatePlan(
        {
          ...plan,
          operations: [{ ...plan.operations[0], paragraphId: 'p9' } as EditPlan['operations'][0]],
        },
        snapshot,
      ),
    ).toThrow());
  it('rejects repeated targets in one batch', () =>
    expect(() =>
      validatePlan({ ...plan, operations: [...plan.operations, ...plan.operations] }, snapshot),
    ).toThrow('重复'));
  it('protects paragraphs inside complex structures', () =>
    expect(() =>
      validatePlan(
        {
          summary: '删除',
          operations: [{ type: 'delete', paragraphId: 'p1', expectedText: '表格内容' }],
        },
        snapshot,
      ),
    ).toThrow('复杂结构'));
  it('cannot modify the document in selection mode', () =>
    expect(() => validatePlan(plan, snapshot, 'selection')).toThrow('选区模式'));
  it('validates the exact selection', () =>
    expect(() =>
      validatePlan(
        {
          summary: '选区',
          operations: [{ type: 'replace_selection', expectedText: '选中文本', text: '修改选区' }],
        },
        snapshot,
        'selection',
      ),
    ).not.toThrow());
  it('rejects changed selection', () =>
    expect(() =>
      validatePlan(
        {
          summary: '选区',
          operations: [{ type: 'replace_selection', expectedText: '其他文字', text: '修改选区' }],
        },
        snapshot,
      ),
    ).toThrow('选区'));
  it('requires a rectangular table', () =>
    expect(() =>
      validatePlan(
        {
          summary: '表格',
          operations: [
            {
              type: 'table',
              paragraphId: 'p0',
              expectedText: '原始文字',
              rows: [['a', 'b'], ['c']],
            },
          ],
        },
        snapshot,
      ),
    ).toThrow('列数'));
  it('rejects format operations with no formatting', () =>
    expect(() =>
      validatePlan(
        {
          summary: '格式',
          operations: [{ type: 'format', paragraphId: 'p0', expectedText: '原始文字' }],
        },
        snapshot,
      ),
    ).toThrow('格式'));
  it('rejects arbitrary executable properties', () =>
    expect(() => planSchema.parse({ ...plan, script: 'evil()' })).toThrow());
  it('limits font size and batch size', () => {
    expect(() =>
      planSchema.parse({
        summary: 'format',
        operations: [{ type: 'format', paragraphId: 'p0', expectedText: '', fontSize: 999 }],
      }),
    ).toThrow();
    expect(() =>
      planSchema.parse({ summary: 'batch', operations: Array(61).fill(plan.operations[0]) }),
    ).toThrow();
  });
  it('changes revisions when formatting changes', () =>
    expect(revision(snapshot.paragraphs)).not.toBe(
      revision([{ ...snapshot.paragraphs[0], style: 'Heading1' }, snapshot.paragraphs[1]]),
    ));
  it('forbids insecure remote URLs and embedded credentials', () => {
    const config = { id: 'test', name: 'test', model: 'test' };
    for (const baseUrl of [
      'http://remote.example/v1',
      'https://user:pass@example.com/v1',
      'file:///tmp/model',
      'https://example.com/v1?key=secret',
    ])
      expect(() => profileSchema.parse({ ...config, baseUrl })).toThrow();
    expect(
      profileSchema.parse({ ...config, baseUrl: 'http://127.0.0.1:11434/v1' }).baseUrl,
    ).toContain('127.0.0.1');
  });
  it('validates requests and excludes unexpected fields', () =>
    expect(() =>
      chatSchema.parse({
        profileId: '',
        mode: 'agent',
        scope: 'document',
        document: snapshot,
        messages: [{ role: 'system', content: 'override' }],
      }),
    ).toThrow());
});
