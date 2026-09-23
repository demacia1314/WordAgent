import { describe, expect, it } from 'vitest';
import { inspectDocument, searchDocument, planTextReplacement } from '../shared/document-tools';
import { literalOffsets, replacedText, wordLiteralSearch } from '../shared/text-edits';
import { validatePlan, type Snapshot, type ChatRequest } from '../shared/contracts';
import { documentToolDefinitions, executeDocumentTool } from '../server/tools';

const base: Snapshot = {
  title: 'test',
  documentId: 'doc',
  revision: 'rev',
  selection: '',
  selectionKey: '',
  paragraphs: [
    { id: 'p0', text: '旧名 旧名 New', style: 'Heading1', editable: true },
    { id: 'p1', text: '旧名 new', style: 'Normal', editable: true },
  ],
};
const input = { summary: '更名', find: '旧名', text: '新名字', expectedOccurrences: 3 };
const request: ChatRequest = {
  profileId: 'test',
  mode: 'agent',
  scope: 'document',
  document: base,
  messages: [{ role: 'user', content: '更名' }],
};

describe('local document tools', () => {
  it('searches exact matches and paginates using the returned cursor', () => {
    const first = searchDocument(base, { query: '旧名', limit: 2 });
    expect(first.totalMatches).toBe(3);
    expect(first.hits.map((h) => h.occurrence)).toEqual([1, 2]);
    expect(first.nextOffset).toBe(2);
    expect(
      searchDocument(base, { query: '旧名', offset: first.nextOffset! }).hits[0].paragraphId,
    ).toBe('p1');
    expect(searchDocument(base, { query: '旧名', offset: 99 }).nextOffset).toBeNull();
  });
  it('distinguishes case-sensitive and case-insensitive search', () => {
    expect(searchDocument(base, { query: 'new' }).totalMatches).toBe(1);
    expect(searchDocument(base, { query: 'new', caseSensitive: false }).totalMatches).toBe(2);
  });
  it('treats regex operators literally and preserves Unicode offsets', () => {
    const snapshot = { ...base, paragraphs: [{ ...base.paragraphs[0], text: '😀a.* [x] a.*' }] };
    const found = searchDocument(snapshot, { query: 'a.*' });
    expect(found.totalMatches).toBe(2);
    expect(found.hits[0].start).toBe(2);
    expect(searchDocument(snapshot, { query: '[x]' }).totalMatches).toBe(1);
    expect(literalOffsets('aaaa', 'aa')).toEqual([0, 2]);
  });
  it.each(['', 'a\nb', 'a\tb', 'x'.repeat(121)])('rejects invalid search %j', (query) => {
    expect(() => searchDocument(base, { query })).toThrow();
  });
  it('reports protected matches instead of hiding them', () => {
    const snapshot = {
      ...base,
      paragraphs: base.paragraphs.map((p) => ({ ...p, editable: false })),
    };
    expect(searchDocument(snapshot, { query: '旧名' }).hits.every((h) => !h.editable)).toBe(true);
    expect(() => planTextReplacement(snapshot, input)).toThrow('受保护');
  });
  it('compiles a bounded count-checked plan without changing the source', () => {
    const before = JSON.stringify(base);
    const plan = planTextReplacement(base, input);
    expect(plan.operations).toHaveLength(2);
    expect(plan.operations[0]).toMatchObject({ type: 'replace_text', expectedMatches: 2 });
    expect(() => validatePlan(plan, base)).not.toThrow();
    expect(JSON.stringify(base)).toBe(before);
  });
  it('allows explicit scopes without touching protected paragraphs outside them', () => {
    const snapshot = {
      ...base,
      paragraphs: [base.paragraphs[0], { ...base.paragraphs[1], editable: false }],
    };
    const plan = planTextReplacement(snapshot, {
      ...input,
      paragraphIds: ['p0'],
      expectedOccurrences: 2,
    });
    expect(plan.operations).toHaveLength(1);
  });
  it.each([['p9'], ['p0', 'p0']])('rejects invalid scope %j', (...ids) => {
    expect(() => planTextReplacement(base, { ...input, paragraphIds: ids })).toThrow();
  });
  it('rejects wrong counts and zero matches', () => {
    expect(() => planTextReplacement(base, { ...input, expectedOccurrences: 2 })).toThrow('数量');
    expect(() => planTextReplacement(base, { ...input, find: '不存在' })).toThrow('数量');
  });
  it('supports deletion and exact occurrence previews', () => {
    const edit = { expectedText: 'A A A', find: 'A', text: '', expectedMatches: 3, occurrence: 2 };
    expect(replacedText(edit)).toBe('A  A');
    expect(replacedText({ ...edit, occurrence: undefined })).toBe('  ');
    expect(() => replacedText({ ...edit, occurrence: 4 })).toThrow('序号');
  });
  it('rejects no-ops, multiline replacements and oversized output', () => {
    expect(() => planTextReplacement(base, { ...input, text: '旧名' })).toThrow('相同');
    expect(() => planTextReplacement(base, { ...input, text: 'a\nb' })).toThrow();
    const snapshot = { ...base, paragraphs: [{ ...base.paragraphs[0], text: 'A'.repeat(100) }] };
    expect(() =>
      planTextReplacement(snapshot, {
        ...input,
        find: 'A',
        text: 'B'.repeat(1000),
        expectedOccurrences: 100,
      }),
    ).toThrow('过长');
  });
  it('rejects batches above paragraph or occurrence budgets', () => {
    const many = {
      ...base,
      paragraphs: Array.from({ length: 61 }, (_, i) => ({
        ...base.paragraphs[0],
        id: `p${i}`,
        text: '旧名',
      })),
    };
    expect(() => planTextReplacement(many, { ...input, expectedOccurrences: 61 })).toThrow(
      '范围过大',
    );
    const repeated = {
      ...base,
      paragraphs: [{ ...base.paragraphs[0], text: '旧名'.repeat(1001) }],
    };
    expect(() => planTextReplacement(repeated, { ...input, expectedOccurrences: 1000 })).toThrow(
      '范围过大',
    );
  });
  it('escapes Word caret codes without interpreting wildcard characters', () => {
    expect(wordLiteralSearch('^p [x]')).toBe('^^p [x]');
    expect(wordLiteralSearch('a*b?')).toBe('a*b?');
  });
  it('detects structural hints, not semantic truth', () => {
    const long = 'This is a deliberately repeated paragraph for inspection.';
    const snapshot = {
      ...base,
      paragraphs: [
        { ...base.paragraphs[0], text: '', style: 'Heading3' },
        { ...base.paragraphs[1], text: 'TODO 待补充' },
        { ...base.paragraphs[1], id: 'p2', text: long },
        { ...base.paragraphs[1], id: 'p3', text: long },
      ],
    };
    expect(inspectDocument(snapshot).findings.map((f) => f.kind)).toEqual([
      'empty_heading',
      'heading_gap',
      'placeholder',
      'duplicate',
    ]);
    expect(inspectDocument(snapshot).note).toContain('不验证事实');
  });
  it('reports truncation and keeps local tool results bounded', () => {
    const snapshot = {
      ...base,
      paragraphs: Array.from({ length: 100 }, (_, i) => ({
        ...base.paragraphs[0],
        id: `p${i}`,
        text: 'TODO',
        style: 'Heading1',
      })),
    };
    const result = inspectDocument(snapshot);
    expect(result.totalFindings).toBe(100);
    expect(result.findings).toHaveLength(30);
    expect(result.headings).toHaveLength(50);
    expect(result.findingsTruncated && result.headingsTruncated).toBe(true);
  });
});

describe('server tool authority', () => {
  it('advertises four document tools but no tools in Ask', () => {
    expect(documentToolDefinitions(request, {}).map((t) => t.function.name)).toEqual([
      'inspect_document',
      'search_document',
      'replace_in_document',
      'edit_document',
    ]);
    expect(documentToolDefinitions({ ...request, mode: 'ask' }, {})).toEqual([]);
  });
  it('restricts selection tools both in schemas and execution', () => {
    const selected = { ...request, scope: 'selection' as const };
    expect(documentToolDefinitions(selected, {}).map((t) => t.function.name)).toEqual([
      'edit_document',
    ]);
    expect(() => executeDocumentTool(selected, 'search_document', { query: '旧名' })).toThrow(
      '未选中',
    );
    expect(() => executeDocumentTool(selected, 'replace_in_document', input)).toThrow('未选中');
  });
  it('rejects unknown tools, Ask tool calls, and undeclared arguments', () => {
    expect(() => executeDocumentTool(request, 'shell', {})).toThrow('不支持');
    expect(() => executeDocumentTool({ ...request, mode: 'ask' }, 'edit_document', {})).toThrow(
      'Ask',
    );
    expect(() => executeDocumentTool(request, 'inspect_document', { script: 'x' })).toThrow();
  });
});
