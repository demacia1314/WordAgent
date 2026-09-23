import type { EditPlan, Snapshot } from '../../shared/contracts';

export type Checkpoint = {
  before: unknown;
  after: unknown;
  afterRevision: string;
  documentId: string;
};
export interface DocumentAdapter {
  kind: 'browser' | 'word';
  snapshot(): Promise<Snapshot>;
  capture?(): Promise<Snapshot>;
  apply(plan: EditPlan, base: Snapshot): Promise<Checkpoint>;
  undo(checkpoint: Checkpoint): Promise<void>;
  focus(paragraphId: string): Promise<void>;
}
export const CONFLICT = '文档在生成后发生了变化。为保护你的修改，本次未应用，请重新发送请求。';
export const UNDO_CONFLICT = '文档在应用后又发生了变化，不能直接撤回。请先使用编辑器撤销后续修改。';
export function fingerprint(text: string): string {
  let a = 2166136261;
  let b = 5381;
  for (let i = 0; i < text.length; i++) {
    a = Math.imul(a ^ text.charCodeAt(i), 16777619);
    b = Math.imul(b, 33) ^ text.charCodeAt(i);
  }
  return `${text.length}:${(a >>> 0).toString(36)}:${(b >>> 0).toString(36)}`;
}
