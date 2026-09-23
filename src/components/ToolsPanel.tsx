import './ToolsPanel.css';
import { useMemo, useState } from 'react';
import { Search, RefreshCw, ShieldCheck, WandSparkles } from 'lucide-react';
import type { EditPlan, Snapshot } from '../../shared/contracts';
import { inspectDocument, searchDocument, planTextReplacement } from '../../shared/document-tools';
import { toolRecipes, type ToolRecipe } from '../lib/tool-recipes';

export function ToolsPanel({
  snapshot,
  busy,
  onRefresh,
  onFocus,
  onStage,
  onRecipe,
}: {
  snapshot: Snapshot;
  busy: boolean;
  onRefresh: () => void;
  onFocus: (id: string) => void;
  onStage: (plan: EditPlan, base: Snapshot) => void;
  onRecipe: (recipe: ToolRecipe) => void;
}) {
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [offset, setOffset] = useState(0);
  const [error, setError] = useState('');
  const inspection = useMemo(() => inspectDocument(snapshot), [snapshot]);
  const search = useMemo(() => {
    if (!query) return { result: null, error: '' };
    try {
      return { result: searchDocument(snapshot, { query, offset, limit: 20 }), error: '' };
    } catch {
      return { result: null, error: '请输入不含换行或控制字符的查找文字（最多 120 字符）' };
    }
  }, [snapshot, query, offset]);
  return (
    <div className="panel-scroll tools-panel">
      <div className="section-label">
        <span>本机文档工具</span>
        <button type="button" className="text-button" disabled={busy} onClick={onRefresh}>
          <RefreshCw size={13} />
          刷新
        </button>
      </div>
      <p className="tool-note">
        <ShieldCheck size={14} />
        以下查找、统计与替换预览不调用模型，不发送文档。
      </p>
      <form
        className="tool-search"
        onSubmit={(event) => {
          event.preventDefault();
          setError('');
          try {
            const plan = planTextReplacement(snapshot, {
              summary: `精确替换“${query}”`,
              find: query,
              text: replacement,
              expectedOccurrences: search.result?.totalMatches ?? 0,
            });
            onStage(plan, snapshot);
          } catch (reason) {
            setError(reason instanceof Error ? reason.message : '无法生成替换方案');
          }
        }}
      >
        <label htmlFor="tool-find">查找文字（区分大小写、精确子串）</label>
        <div className="search-field">
          <Search size={14} />
          <input
            id="tool-find"
            maxLength={120}
            value={query}
            onChange={(event) => {
              setQuery(event.target.value);
              setOffset(0);
              setError('');
            }}
            placeholder="例如：旧项目名称"
          />
        </div>
        <label htmlFor="tool-replacement">替换为（留空表示删除匹配文字）</label>
        <input
          id="tool-replacement"
          maxLength={2000}
          value={replacement}
          onChange={(event) => {
            setReplacement(event.target.value);
            setError('');
          }}
        />
        <button
          className="button primary"
          type="submit"
          disabled={
            busy || !snapshot.documentId || !search.result?.totalMatches || query === replacement
          }
        >
          预览全部替换{search.result ? ` · ${search.result.totalMatches} 处` : ''}
        </button>
        <small className="muted">
          仅生成待审阅方案，不立即写入。包含受保护内容时整批拒绝，不会悄悄跳过。
        </small>
      </form>
      {(error || search.error) && (
        <p role="alert" className="inline-error">
          {error || search.error}
        </p>
      )}
      {search.result && (
        <section aria-label="查找结果">
          <div className="section-label">找到 {search.result.totalMatches} 处匹配</div>
          {search.result.hits.map((hit) => (
            <button
              type="button"
              className="tool-hit"
              key={`${hit.paragraphId}:${hit.start}`}
              disabled={busy}
              onClick={() => onFocus(hit.paragraphId)}
            >
              <small>
                段落 {Number(hit.paragraphId.slice(1)) + 1} · 第 {hit.occurrence} 处
                {hit.editable ? '' : ' · 受保护'}
              </small>
              <span>{hit.excerpt}</span>
            </button>
          ))}
          <div className="tool-pages">
            <button
              type="button"
              className="text-button"
              disabled={!offset}
              onClick={() => setOffset(Math.max(0, offset - 20))}
            >
              上一页
            </button>
            <button
              type="button"
              className="text-button"
              disabled={search.result.nextOffset === null}
              onClick={() => setOffset(search.result!.nextOffset!)}
            >
              下一页
            </button>
          </div>
        </section>
      )}
      <section aria-label="文档检查">
        <div className="section-label">结构检查 · {inspection.totalFindings} 项提示</div>
        <p className="tool-note">
          {inspection.statistics.paragraphs} 段 · {inspection.statistics.nonWhitespaceCharacters}{' '}
          非空白字符 · {inspection.statistics.headings} 个标题 ·{' '}
          {inspection.statistics.protectedParagraphs} 段受保护
        </p>
        {inspection.findings.map((finding, index) => (
          <button
            type="button"
            className="tool-hit"
            key={index}
            disabled={busy}
            onClick={() => onFocus(finding.paragraphId)}
          >
            <small>
              段落 {Number(finding.paragraphId.slice(1)) + 1} · {finding.message}
            </small>
            <span>{finding.excerpt || '（空内容）'}</span>
          </button>
        ))}
        {inspection.findingsTruncated && <p className="muted">仅展示前 30 项提示。</p>}
        <small className="muted">
          规则提示可能是有意的写法；不代表事实、语义或引文核验。字符统计不是 Word 字数。
        </small>
      </section>
      <div className="section-label">AI 写作工具</div>
      <p className="tool-note">点击只填入提示词。你可以修改范围和要求，再发送给已授权的模型。</p>
      <div className="tool-recipes">
        {toolRecipes.map((recipe) => (
          <button type="button" key={recipe.id} disabled={busy} onClick={() => onRecipe(recipe)}>
            <WandSparkles size={16} />
            <span>
              <strong>{recipe.title}</strong>
              <small>{recipe.description}</small>
            </span>
            <small>{recipe.mode === 'ask' ? '问答' : '编辑'}</small>
          </button>
        ))}
      </div>
    </div>
  );
}
