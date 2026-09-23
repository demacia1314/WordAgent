import { z } from 'zod';

// Short literal strings work consistently in both ProseMirror and Word search.
// No regex, paragraph breaks, tabs, or Word control characters are executable input.
export const literalTextSchema = z
  .string()
  .min(1)
  .max(120)
  .refine(
    (value) => !/[\u0000-\u001f\u007f\u2028\u2029]/u.test(value),
    '查找文字不能包含换行或控制字符',
  );
export const replacementTextSchema = z
  .string()
  .max(2000)
  .refine(
    (value) => !/[\u0000-\u001f\u007f\u2028\u2029]/u.test(value),
    '局部替换不能包含换行或控制字符',
  );

export type TextEdit = {
  expectedText: string;
  find: string;
  text: string;
  expectedMatches: number;
  occurrence?: number;
};

/** Exact, case-sensitive, non-overlapping UTF-16 offsets into the original text. */
export function literalOffsets(text: string, find: string): number[] {
  if (!find) throw new Error('查找文字不能为空');
  const offsets: number[] = [];
  let from = 0;
  while (from <= text.length - find.length) {
    const index = text.indexOf(find, from);
    if (index < 0) break;
    offsets.push(index);
    from = index + find.length;
  }
  return offsets;
}

export function replacementOffsets(edit: TextEdit): number[] {
  const offsets = literalOffsets(edit.expectedText, edit.find);
  if (offsets.length !== edit.expectedMatches || !offsets.length)
    throw new Error('匹配数量与预期不符，请重新查找后生成修改');
  if (edit.find === edit.text) throw new Error('替换前后文字相同，无需修改');
  if (edit.occurrence !== undefined) {
    if (
      !Number.isInteger(edit.occurrence) ||
      edit.occurrence < 1 ||
      edit.occurrence > offsets.length
    )
      throw new Error('指定的匹配序号不存在');
    return [offsets[edit.occurrence - 1]];
  }
  return offsets;
}

export function replacedText(edit: TextEdit): string {
  let result = edit.expectedText;
  for (const index of replacementOffsets(edit).reverse())
    result = result.slice(0, index) + edit.text + result.slice(index + edit.find.length);
  return result;
}

/** Word interprets caret codes even with wildcards disabled. */
export function wordLiteralSearch(find: string): string {
  return literalTextSchema.parse(find).replace(/\^/g, '^^');
}
