import { z } from 'zod';
import { planSchema, validatePlan, type ChatRequest, type EditPlan } from '../shared/contracts';
import {
  inspectDocument,
  searchDocument,
  searchInputSchema,
  planTextReplacement,
  replaceInputSchema,
} from '../shared/document-tools';

export const MAX_READ_ROUNDS = 4;
export const toolLabels: Record<string, string> = {
  inspect_document: '检查文档结构',
  search_document: '查找文档文字',
  replace_in_document: '生成精确替换方案',
  edit_document: '校验编辑方案',
};
const object = (properties: Record<string, unknown>, required: string[] = []) => ({
  type: 'object',
  additionalProperties: false,
  properties,
  required,
});
const string = (maxLength: number, minLength = 0) => ({ type: 'string', minLength, maxLength });
const literal = {
  ...string(120, 1),
  description: 'Literal text, not regex. No line breaks or control characters.',
};

/** Ask retains text-only provider compatibility; selection never exposes body tools. */
export function documentToolDefinitions(request: ChatRequest, editParameters: object) {
  if (request.mode === 'ask') return [];
  const edit = {
    type: 'function',
    function: {
      name: 'edit_document',
      description:
        'Propose ONE validated batch only for requested changes. Nothing is applied here. Prefer replace_text for phrase corrections, preserving surrounding formatting. expectedMatches counts ALL exact non-overlapping matches in the original paragraph; occurrence optionally selects one (1-based).',
      parameters:
        request.scope === 'selection'
          ? object(
              {
                summary: string(300, 1),
                operations: {
                  type: 'array',
                  minItems: 1,
                  maxItems: 1,
                  items: object(
                    {
                      type: { const: 'replace_selection' },
                      expectedText: string(30000, 1),
                      text: string(30000),
                    },
                    ['type', 'expectedText', 'text'],
                  ),
                },
              },
              ['summary', 'operations'],
            )
          : editParameters,
    },
  };
  if (request.scope === 'selection') return [edit];
  return [
    {
      type: 'function',
      function: {
        name: 'inspect_document',
        description:
          'Read-only local body overview: counts, headings and heuristic flags for placeholders, skipped heading levels and duplicate paragraphs. Not a semantic, factual or citation check. Explicit truncation flags. No network access.',
        parameters: object({}),
      },
    },
    {
      type: 'function',
      function: {
        name: 'search_document',
        description:
          'Read-only literal body search. Returns exact total count, paragraph IDs, occurrence numbers, snippets and nextOffset. Default caseSensitive=true; use it before replacement. Paginate using nextOffset. Does not search the web or edit.',
        parameters: object(
          {
            query: literal,
            caseSensitive: { type: 'boolean', default: true },
            offset: { type: 'integer', minimum: 0, maximum: 3000000, default: 0 },
            limit: { type: 'integer', minimum: 1, maximum: 50, default: 20 },
          },
          ['query'],
        ),
      },
    },
    {
      type: 'function',
      function: {
        name: 'replace_in_document',
        description:
          'Propose ALL exact case-sensitive substring replacements in paragraphIds, or body if omitted. expectedOccurrences MUST equal the total in that scope, confirmed by search. Not whole-word search. Refuses protected matches, >60 paragraphs or >1000 occurrences. Final proposal only, never writes the document.',
        parameters: object(
          {
            summary: string(300, 1),
            find: literal,
            text: string(2000),
            expectedOccurrences: { type: 'integer', minimum: 1, maximum: 1000 },
            paragraphIds: {
              type: 'array',
              minItems: 1,
              maxItems: 60,
              uniqueItems: true,
              items: { type: 'string', pattern: '^p\\d+$' },
            },
          },
          ['summary', 'find', 'text', 'expectedOccurrences'],
        ),
      },
    },
    edit,
  ];
}

export function executeDocumentTool(
  request: ChatRequest,
  name: string,
  input: unknown,
): { kind: 'read'; result: unknown } | { kind: 'proposal'; plan: EditPlan } {
  if (request.mode !== 'agent') throw new Error('Ask 模式不允许调用编辑工具');
  if (name === 'edit_document') {
    const plan = planSchema.parse(input);
    validatePlan(plan, request.document, request.scope);
    return { kind: 'proposal', plan };
  }
  if (request.scope !== 'document') throw new Error('选区模式不允许读取或修改未选中的正文');
  if (name === 'inspect_document') {
    z.object({}).strict().parse(input);
    return { kind: 'read', result: inspectDocument(request.document) };
  }
  if (name === 'search_document')
    return {
      kind: 'read',
      result: searchDocument(request.document, searchInputSchema.parse(input)),
    };
  if (name === 'replace_in_document')
    return {
      kind: 'proposal',
      plan: planTextReplacement(request.document, replaceInputSchema.parse(input)),
    };
  throw new Error('模型返回了不支持的工具');
}
