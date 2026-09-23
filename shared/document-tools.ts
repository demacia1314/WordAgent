import { z } from 'zod';
import { validatePlan, type EditPlan, type Snapshot } from './contracts';
import { literalOffsets, literalTextSchema, replacementTextSchema } from './text-edits';

export const searchInputSchema = z
  .object({
    query: literalTextSchema,
    caseSensitive: z.boolean().default(true),
    offset: z.number().int().min(0).max(3000000).default(0),
    limit: z.number().int().min(1).max(50).default(20),
  })
  .strict();
export type SearchInput = z.input<typeof searchInputSchema>;
export type SearchHit = {
  paragraphId: string;
  occurrence: number;
  start: number;
  end: number;
  text: string;
  excerpt: string;
  editable: boolean;
};

/** Local literal search. Pagination counts matches, not paragraphs. No user regex. */
export function searchDocument(snapshot: Snapshot, input: SearchInput) {
  const args = searchInputSchema.parse(input);
  const escaped = args.query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const expression = new RegExp(escaped, args.caseSensitive ? 'gu' : 'giu');
  const hits: SearchHit[] = [];
  let totalMatches = 0;
  for (const paragraph of snapshot.paragraphs) {
    expression.lastIndex = 0;
    let match: RegExpExecArray | null;
    let occurrence = 0;
    while ((match = expression.exec(paragraph.text))) {
      occurrence++;
      if (totalMatches >= args.offset && hits.length < args.limit) {
        const start = match.index;
        const end = start + match[0].length;
        hits.push({
          paragraphId: paragraph.id,
          occurrence,
          start,
          end,
          text: match[0],
          excerpt: paragraph.text.slice(
            Math.max(0, start - 60),
            Math.min(paragraph.text.length, end + 60),
          ),
          editable: paragraph.editable,
        });
      }
      totalMatches++;
    }
  }
  const nextOffset = args.offset + hits.length < totalMatches ? args.offset + hits.length : null;
  return {
    query: args.query,
    caseSensitive: args.caseSensitive,
    totalMatches,
    hits,
    nextOffset,
    note: '仅搜索当前快照的正文；序号在每个段落内从 1 开始。替换工具只支持区分大小写的精确匹配。',
  };
}

export type DocumentFinding = {
  paragraphId: string;
  kind: 'heading_gap' | 'empty_heading' | 'placeholder' | 'duplicate';
  message: string;
  excerpt: string;
};

/** Heuristics, not a semantic review or a claim that the document is error-free. */
export function inspectDocument(snapshot: Snapshot) {
  const headings: { paragraphId: string; level: number; text: string }[] = [];
  const findings: DocumentFinding[] = [];
  const seen = new Map<string, string>();
  let previousLevel = 0;
  let characters = 0;
  let nonWhitespaceCharacters = 0;
  let protectedParagraphs = 0;
  for (const paragraph of snapshot.paragraphs) {
    characters += paragraph.text.length;
    nonWhitespaceCharacters += Array.from(paragraph.text.replace(/\s/gu, '')).length;
    if (!paragraph.editable) protectedParagraphs++;
    const heading = /^Heading([1-6])$/.exec(paragraph.style);
    const add = (kind: DocumentFinding['kind'], message: string) =>
      findings.push({
        paragraphId: paragraph.id,
        kind,
        message,
        excerpt: paragraph.text.slice(0, 120),
      });
    if (heading || paragraph.style === 'Title') {
      const level = heading ? Number(heading[1]) : 0;
      headings.push({ paragraphId: paragraph.id, level, text: paragraph.text.slice(0, 120) });
      if (!paragraph.text.trim()) add('empty_heading', '标题为空');
      if (heading) {
        if (level > previousLevel + 1)
          add('heading_gap', `标题层级从 ${previousLevel || '正文'} 跳到 ${level}，请核对结构`);
        previousLevel = level;
      }
    }
    if (/\b(?:TODO|TBD|FIXME)\b|待补充|待完善|此处插入/iu.test(paragraph.text))
      add('placeholder', '可能存在尚未完成的占位内容');
    const normalized = paragraph.text.trim().replace(/\s+/gu, ' ');
    if (normalized.length >= 20) {
      const earlier = seen.get(normalized);
      if (earlier)
        add('duplicate', `与段落 ${Number(earlier.slice(1)) + 1} 内容重复，请确认是否有意保留`);
      else seen.set(normalized, paragraph.id);
    }
  }
  return {
    statistics: {
      paragraphs: snapshot.paragraphs.length,
      characters,
      nonWhitespaceCharacters,
      headings: headings.length,
      protectedParagraphs,
    },
    headings: headings.slice(0, 50),
    findings: findings.slice(0, 30),
    totalFindings: findings.length,
    headingsTruncated: headings.length > 50,
    findingsTruncated: findings.length > 30,
    note: '本机规则检查，仅覆盖正文快照；不验证事实、引文或语义。字符数不是 Word 的字数统计。',
  };
}

export const replaceInputSchema = z
  .object({
    summary: z.string().min(1).max(300),
    find: literalTextSchema,
    text: replacementTextSchema,
    expectedOccurrences: z.number().int().min(1).max(1000),
    paragraphIds: z
      .array(z.string().regex(/^p\d+$/))
      .min(1)
      .max(60)
      .optional(),
  })
  .strict();

/** Compile exact replacements to the existing guarded proposal/apply/undo pipeline. */
export function planTextReplacement(
  snapshot: Snapshot,
  input: z.input<typeof replaceInputSchema>,
): EditPlan {
  const args = replaceInputSchema.parse(input);
  const ids = args.paragraphIds ? new Set(args.paragraphIds) : undefined;
  if (
    ids &&
    (ids.size !== args.paragraphIds!.length ||
      [...ids].some((id) => !snapshot.paragraphs.some((p) => p.id === id)))
  )
    throw new Error('段落范围重复或不存在，请重新查找');
  const operations: EditPlan['operations'] = [];
  let total = 0;
  for (const paragraph of snapshot.paragraphs) {
    if (ids && !ids.has(paragraph.id)) continue;
    const count = literalOffsets(paragraph.text, args.find).length;
    if (!count) continue;
    if (!paragraph.editable)
      throw new Error('匹配包含受保护的表格或复杂结构；请明确限定可编辑段落范围');
    total += count;
    if (total > 1000 || operations.length >= 60)
      throw new Error('替换范围过大，请限定段落后分批处理');
    operations.push({
      type: 'replace_text',
      paragraphId: paragraph.id,
      expectedText: paragraph.text,
      find: args.find,
      text: args.text,
      expectedMatches: count,
    });
  }
  if (total !== args.expectedOccurrences || total === 0)
    throw new Error('匹配数量与预期不符，请使用区分大小写的查找确认范围和数量');
  const plan = { summary: args.summary, operations };
  validatePlan(plan, snapshot);
  return plan;
}
