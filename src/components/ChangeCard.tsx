import {
  Check,
  ChevronDown,
  ChevronRight,
  FileDiff,
  RotateCcw,
  X,
  CircleAlert,
  LoaderCircle,
} from 'lucide-react';
import { useState } from 'react';
import { diffChars } from 'diff';
import { replacedText } from '../../shared/text-edits';
import type { Change } from '../lib/types';
import type { Operation } from '../../shared/contracts';

const labels = {
  pending: '待审阅',
  applying: '应用中',
  applied: '已应用',
  rejected: '已放弃',
  reverted: '已撤回',
  conflict: '存在冲突',
};
const operationNames = {
  replace: '修改段落',
  replace_text: '精确替换文字',
  insert: '插入段落',
  delete: '删除段落',
  format: '调整格式',
  table: '插入表格',
  replace_selection: '修改选区',
};
const styleNames: Record<string, string> = {
  Normal: '正文',
  Title: '标题',
  Heading1: '标题 1',
  Heading2: '标题 2',
  Heading3: '标题 3',
  Quote: '引用',
};
function OperationDiff({ op }: { op: Operation }) {
  if (op.type === 'format')
    return (
      <div className="format-diff">
        {op.style && <span>{styleNames[op.style]}</span>}
        {op.fontSize && <span>{op.fontSize} pt</span>}
        {op.fontFamily && <span>{op.fontFamily}</span>}
        {op.bold !== undefined && <span>{op.bold ? '加粗' : '取消加粗'}</span>}
        {op.italic !== undefined && <span>{op.italic ? '斜体' : '取消斜体'}</span>}
        {op.alignment && (
          <span>
            {{ left: '左对齐', center: '居中', right: '右对齐', justify: '两端对齐' }[op.alignment]}
          </span>
        )}
      </div>
    );
  if (op.type === 'table')
    return (
      <div className="diff-table-wrap">
        <table className="diff-table">
          <tbody>
            {op.rows.map((row, i) => (
              <tr key={i}>
                {row.map((cell, j) => (
                  <td key={j}>{cell}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  if (op.type === 'insert')
    return (
      <div className="diff-text">
        <ins>{op.text}</ins>
      </div>
    );
  if (op.type === 'delete')
    return (
      <div className="diff-text">
        <del>{op.expectedText || '空段落'}</del>
      </div>
    );
  return (
    <div className="diff-text">
      {diffChars(op.expectedText, op.type === 'replace_text' ? replacedText(op) : op.text).map(
        (part, i) =>
          part.added ? (
            <ins key={i}>{part.value}</ins>
          ) : part.removed ? (
            <del key={i}>{part.value}</del>
          ) : (
            <span key={i}>{part.value}</span>
          ),
      )}
    </div>
  );
}
export function ChangeCard({
  change,
  busy,
  canUndo,
  onApply,
  onReject,
  onUndo,
}: {
  change: Change;
  busy: boolean;
  canUndo: boolean;
  onApply: () => void;
  onReject: () => void;
  onUndo: () => void;
}) {
  const [expanded, setExpanded] = useState(true);
  const pending = change.status === 'pending';
  return (
    <section className={`change-card status-${change.status}`} aria-label="文档变更">
      <button
        className="change-heading"
        onClick={() => setExpanded(!expanded)}
        aria-expanded={expanded}
      >
        <FileDiff size={16} />
        <span>{change.plan.summary}</span>
        {expanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
      </button>
      <div className="change-meta">
        <span>{change.plan.operations.length} 处变更</span>
        <span className={`change-status ${change.status}`}>
          {change.status === 'applying' ? (
            <LoaderCircle size={12} className="spin" />
          ) : change.status === 'conflict' ? (
            <CircleAlert size={12} />
          ) : change.status === 'applied' ? (
            <Check size={12} />
          ) : null}
          {labels[change.status]}
        </span>
      </div>
      {expanded && (
        <div className="operations">
          {change.plan.operations.map((op, i) => (
            <div className="operation" key={i}>
              <div className="operation-label">
                <span>{operationNames[op.type]}</span>
                <span>
                  {'paragraphId' in op ? `段落 ${Number(op.paragraphId.slice(1)) + 1}` : '当前选区'}
                </span>
              </div>
              <OperationDiff op={op} />
            </div>
          ))}
        </div>
      )}
      {change.error && (
        <p className="inline-error" role="alert">
          {change.error}
        </p>
      )}
      {pending && (
        <div className="change-actions">
          <button className="button primary" disabled={busy} onClick={onApply}>
            <Check size={14} />
            应用更改
          </button>
          <button className="button subtle" disabled={busy} onClick={onReject}>
            <X size={14} />
            放弃
          </button>
        </div>
      )}
      {change.status === 'applied' && (
        <div className="change-actions">
          <button
            className="button subtle"
            onClick={onUndo}
            disabled={busy || !canUndo}
            title={
              canUndo ? '撤回此批次；后续改动会触发冲突保护' : '仅当前打开期间最近一批编辑可撤回'
            }
          >
            <RotateCcw size={14} />
            撤回本次
          </button>
        </div>
      )}
    </section>
  );
}
