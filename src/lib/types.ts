import type { EditPlan, Snapshot, ToolActivity } from '../../shared/contracts';
export type ChangeStatus =
  'pending' | 'applying' | 'applied' | 'rejected' | 'reverted' | 'conflict';
export type Change = {
  id: string;
  plan: EditPlan;
  base: Snapshot;
  status: ChangeStatus;
  created: number;
  error?: string;
};
export type Message = {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  created: number;
  change?: Change;
  tools?: ToolActivity[];
  error?: string;
  interrupted?: boolean;
};
export type Session = {
  id: string;
  documentId: string;
  title: string;
  created: number;
  updated: number;
  messages: Message[];
};
export const uid = () => crypto.randomUUID();
