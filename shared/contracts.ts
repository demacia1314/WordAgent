import { z } from 'zod';

export const styleSchema = z.enum(['Normal', 'Title', 'Heading1', 'Heading2', 'Heading3', 'Quote']);
export type ParagraphStyle = z.infer<typeof styleSchema>;
const text = z.string().max(30000);
const target = { paragraphId: z.string().regex(/^p\d+$/), expectedText: text };
export const operationSchema = z.discriminatedUnion('type', [
  z.object({ type: z.literal('replace'), ...target, text }).strict(),
  z
    .object({
      type: z.literal('insert'),
      ...target,
      position: z.enum(['before', 'after']),
      text: text.min(1),
      style: styleSchema.optional(),
    })
    .strict(),
  z.object({ type: z.literal('delete'), ...target }).strict(),
  z
    .object({
      type: z.literal('format'),
      ...target,
      style: styleSchema.optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      fontSize: z.number().min(8).max(72).optional(),
      fontFamily: z.string().trim().min(1).max(100).optional(),
      alignment: z.enum(['left', 'center', 'right', 'justify']).optional(),
    })
    .strict(),
  z
    .object({
      type: z.literal('table'),
      ...target,
      rows: z
        .array(z.array(z.string().max(2000)).min(1).max(12))
        .min(1)
        .max(50),
    })
    .strict(),
  z.object({ type: z.literal('replace_selection'), expectedText: text.min(1), text }).strict(),
]);
export type Operation = z.infer<typeof operationSchema>;
export const planSchema = z
  .object({
    summary: z.string().min(1).max(300),
    operations: z.array(operationSchema).min(1).max(60),
  })
  .strict();
export type EditPlan = z.infer<typeof planSchema>;
export const paragraphSchema = z.object({
  id: z.string().regex(/^p\d+$/),
  text,
  style: z.string().max(100),
  bold: z.boolean().nullable().optional(),
  italic: z.boolean().nullable().optional(),
  fontSize: z.number().nullable().optional(),
  fontFamily: z.string().nullable().optional(),
  alignment: z.string().optional(),
  editable: z.boolean().default(true),
});
export type Paragraph = z.infer<typeof paragraphSchema>;
export const snapshotSchema = z.object({
  title: z.string().max(250),
  revision: z.string().max(100),
  paragraphs: z.array(paragraphSchema).max(3000),
  selection: z.string().max(30000),
  selectionKey: z.string().max(100),
  documentId: z.string().max(200),
});
export type Snapshot = z.infer<typeof snapshotSchema>;
export const chatSchema = z
  .object({
    profileId: z.string().max(100),
    mode: z.enum(['agent', 'ask']),
    scope: z.enum(['document', 'selection']),
    messages: z
      .array(
        z.object({ role: z.enum(['user', 'assistant']), content: z.string().max(30000) }).strict(),
      )
      .min(1)
      .max(40),
    document: snapshotSchema,
  })
  .strict();
export type ChatRequest = z.infer<typeof chatSchema>;
export const profileSchema = z
  .object({
    id: z.string().regex(/^[\w.-]{1,100}$/),
    name: z.string().trim().min(1).max(80),
    baseUrl: z
      .string()
      .url()
      .max(500)
      .refine((value) => {
        const url = new URL(value);
        return (
          !url.username &&
          !url.password &&
          !url.search &&
          !url.hash &&
          (url.protocol === 'https:' ||
            (url.protocol === 'http:' &&
              ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname)))
        );
      }, '接口必须使用 HTTPS；本机模型允许 HTTP'),
    model: z.string().trim().min(1).max(150),
    apiKey: z.string().max(4096).optional(),
    temperature: z.number().min(0).max(2).default(0.3),
    maxTokens: z.number().int().min(256).max(32768).default(8192),
    reasoningEffort: z.enum(['minimal', 'low', 'medium', 'high', 'xhigh', 'max']).optional(),
  })
  .strict();
export type Profile = z.infer<typeof profileSchema>;
export type PublicProfile = Omit<Profile, 'apiKey'> & { hasKey: boolean; warning?: string };
export type StreamEvent =
  | { type: 'status'; message: string }
  | { type: 'delta'; text: string }
  | { type: 'proposal'; plan: EditPlan }
  | { type: 'error'; message: string }
  | { type: 'done' };

export function revision(paragraphs: Paragraph[]): string {
  const value = JSON.stringify(paragraphs);
  let hash = 2166136261;
  for (let i = 0; i < value.length; i++) hash = Math.imul(hash ^ value.charCodeAt(i), 16777619);
  return `${value.length}-${(hash >>> 0).toString(36)}`;
}

export function validatePlan(
  plan: EditPlan,
  snapshot: Snapshot,
  scope: 'document' | 'selection' = 'document',
): void {
  planSchema.parse(plan);
  const targets = new Set<string>();
  for (const op of plan.operations) {
    if (scope === 'selection' && op.type !== 'replace_selection')
      throw new Error('选区模式只允许修改选中的文字');
    if (op.type === 'replace_selection') {
      if (
        plan.operations.length !== 1 ||
        !snapshot.selection ||
        snapshot.selection !== op.expectedText
      )
        throw new Error('选区已变化或编辑计划不匹配');
      continue;
    }
    const para = snapshot.paragraphs.find((p) => p.id === op.paragraphId);
    if (!para || para.text !== op.expectedText)
      throw new Error(`段落 ${op.paragraphId} 的原文不匹配，请重新生成`);
    if (!para.editable) throw new Error('此段落位于复杂结构中，请在 Word 中直接编辑');
    if (targets.has(op.paragraphId)) throw new Error('同一批次不能重复修改同一段落');
    targets.add(op.paragraphId);
    if (op.type === 'table' && op.rows.some((row) => row.length !== op.rows[0].length))
      throw new Error('表格每行的列数必须一致');
    if (op.type === 'format' && Object.keys(op).length <= 3) throw new Error('缺少格式设置');
  }
}
