import { useCallback, useEffect, useRef, useState } from 'react';
import { useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { TableKit } from '@tiptap/extension-table';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyleKit } from '@tiptap/extension-text-style';
import DOMPurify from 'dompurify';
import Markdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import {
  ArrowDown,
  ArrowUp,
  Check,
  ChevronDown,
  ChevronRight,
  CircleAlert,
  Copy,
  Download,
  FileCheck2,
  FileText,
  GitCompareArrows,
  History,
  ListTree,
  LoaderCircle,
  MessageSquare,
  PanelRightClose,
  Plus,
  RefreshCw,
  Search,
  Settings2,
  ShieldCheck,
  Sparkles,
  Square,
  TextCursorInput,
  Trash2,
  WandSparkles,
  Wrench,
  X,
  Zap,
  PencilLine,
} from 'lucide-react';
import { BrowserAdapter } from './document/browser';
import { WordAdapter } from './document/word';
import type { Checkpoint, DocumentAdapter } from './document/adapter';
import { api, streamChat, type PublicSettings } from './lib/api';
import { download, readLocal, writeLocal } from './lib/storage';
import { uid, type Change, type Message, type Session } from './lib/types';
import type { EditPlan, Snapshot } from '../shared/contracts';
import { ToolsPanel } from './components/ToolsPanel';
import { IconButton } from './components/IconButton';
import { ChangeCard } from './components/ChangeCard';
import { SettingsPanel } from './components/SettingsPanel';
import { EditorWorkspace } from './components/EditorWorkspace';

const sample =
  '<h1>面向文档协作的智能编辑系统</h1><p>项目研究笔记 · 初稿</p><h2>1. 研究背景</h2><p>文档是知识工作的重要载体。从资料整理、内容撰写到多轮修改，一份文档往往需要在多个工具之间流转。作者不仅要关注文字本身，还需要持续维护内容结构与表达的一致性。</p><p>现有的 AI 写作工具能够生成和改写文本，但生成结果通常需要手动复制回文档。这个过程打断了写作节奏，也使局部修改与上下文之间的关系变得不够清晰。</p><h2>2. 设计目标</h2><p>本项目希望把对话式交互带入文档编辑现场，让作者在保留编辑控制权的同时，更自然地完成内容修改。</p><ul><li><p>理解当前文档与选中内容，准确定位需要修改的位置。</p></li><li><p>在同一个工作区完成讨论、修改和结果审阅。</p></li><li><p>保留清晰的变更记录，使每一次修改都可以被检查。</p></li></ul><h2>3. 交互流程</h2><p>作者提出修改意图后，系统读取文档上下文并生成具体的编辑计划。修改内容经过校验后应用到文档，作者可以查看原文与新内容的差异，并决定是否保留。</p><h2>4. 待验证的问题</h2><p>如何在提高修改效率的同时，让作者始终清楚文档发生了哪些变化？这将是后续原型验证的重点。</p><p></p>';
const emptySnapshot: Snapshot = {
  title: '',
  documentId: '',
  revision: '',
  paragraphs: [],
  selection: '',
  selectionKey: '',
};
const navItems = [
  { id: 'chat', name: '对话', icon: MessageSquare },
  { id: 'outline', name: '文档大纲', icon: ListTree },
  { id: 'tools', name: '文档工具', icon: Wrench },
  { id: 'changes', name: '变更记录', icon: GitCompareArrows },
  { id: 'history', name: '历史会话', icon: History },
] as const;
type Panel = 'chat' | 'outline' | 'tools' | 'changes' | 'history' | 'settings';
type Modal =
  | { kind: 'newDocument' }
  | { kind: 'deleteSession'; id: string }
  | { kind: 'consent'; prompt: string }
  | { kind: 'import'; file: File }
  | { kind: 'clearHistory' };

export default function App({ host }: { host: 'browser' | 'word' }) {
  const [initial] = useState(() =>
    readLocal('document', { title: '研究笔记 · 示例.docx', id: uid(), content: sample }),
  );
  const [title, setTitle] = useState(initial.title);
  const [docId, setDocId] = useState(initial.id);
  const [saved, setSaved] = useState(true);
  const [, setEditorTick] = useState(0);
  const [adapter, setAdapter] = useState<DocumentAdapter | null>(null);
  const [snapshot, setSnapshot] = useState<Snapshot>(emptySnapshot);
  const [panel, setPanel] = useState<Panel>('chat');
  const [mobileView, setMobileView] = useState<'document' | 'assistant'>('assistant');
  const [settings, setSettings] = useState<PublicSettings>({ defaultProfile: '', profiles: [] });
  const [profileId, setProfileId] = useState('');
  const [connected, setConnected] = useState(false);
  const [sessions, setSessions] = useState<Session[]>(() => {
    const value = readLocal<Session[]>('sessions', []);
    return Array.isArray(value)
      ? value
          .filter((s) => s.id && Array.isArray(s.messages))
          .slice(0, 60)
          .map((s) => ({
            ...s,
            messages: s.messages.map((m) =>
              m.change?.status === 'applying'
                ? {
                    ...m,
                    change: {
                      ...m.change,
                      status: 'conflict',
                      error: '上次应用过程被中断，请检查文档后重新生成',
                    },
                  }
                : m,
            ),
          }))
      : [];
  });
  const [activeId, setActiveId] = useState('');
  const [prompt, setPrompt] = useState('');
  const [scope, setScope] = useState<'document' | 'selection'>('document');
  const [mode, setMode] = useState<'agent' | 'ask'>('agent');
  const [live, setLive] = useState(() => readLocal('liveEditing', true));
  const [theme, setTheme] = useState(() => readLocal('theme', 'light'));
  const [saveHistory, setSaveHistory] = useState(() => readLocal('saveHistory', true));
  const [consentFor, setConsentFor] = useState(() => readLocal('consentFor', ''));
  const [busy, setBusy] = useState(false);
  const [actionBusy, setActionBusy] = useState(false);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');
  const [toast, setToast] = useState('');
  const [search, setSearch] = useState('');
  const [contextOpen, setContextOpen] = useState(false);
  const [modal, setModal] = useState<Modal | null>(null);
  const [lastUndoId, setLastUndoId] = useState('');
  const [autoScroll, setAutoScroll] = useState(true);
  const checkpoints = useRef(new Map<string, Checkpoint>());
  const busyRef = useRef(false);
  const abortRef = useRef<AbortController | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const messagesRef = useRef<HTMLDivElement>(null);
  const composerRef = useRef<HTMLTextAreaElement>(null);
  const modalRef = useRef<HTMLDialogElement>(null);
  const toastTimer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const editor = useEditor({
    extensions: [
      StarterKit,
      TableKit.configure({ table: { resizable: true } }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      TextStyleKit,
    ],
    content: initial.content,
    editable: host === 'browser',
    editorProps: { attributes: { 'aria-label': '文档正文', spellcheck: 'false' } },
    onUpdate: () => {
      setSaved(false);
      setEditorTick((n) => n + 1);
    },
    onSelectionUpdate: () => setEditorTick((n) => n + 1),
    onTransaction: () => setEditorTick((n) => n + 1),
  });
  const notify = useCallback((message: string) => {
    setToast(message);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(''), 5000);
  }, []);
  const report = useCallback((reason: unknown) => {
    setError(reason instanceof Error ? reason.message : String(reason));
  }, []);
  const refreshSettings = useCallback(async () => {
    try {
      const result = await api<PublicSettings>('/settings');
      setSettings(result);
      setConnected(true);
      setError('');
      setProfileId((current) =>
        result.profiles.some((p) => p.id === current) ? current : result.defaultProfile,
      );
    } catch (reason) {
      setConnected(false);
      report(reason);
    }
  }, [report]);

  useEffect(() => {
    void refreshSettings();
  }, [refreshSettings]);
  useEffect(() => {
    if (connected) return;
    const timer = setInterval(() => void refreshSettings(), 3000);
    return () => clearInterval(timer);
  }, [connected, refreshSettings]);
  useEffect(() => {
    const media = matchMedia('(prefers-color-scheme: dark)');
    const update = () => {
      document.documentElement.dataset.theme =
        theme === 'system' ? (media.matches ? 'dark' : 'light') : theme;
    };
    update();
    writeLocal('theme', theme);
    media.addEventListener('change', update);
    return () => media.removeEventListener('change', update);
  }, [theme]);
  useEffect(() => {
    writeLocal('consentFor', consentFor);
  }, [consentFor]);
  useEffect(() => {
    writeLocal('saveHistory', saveHistory);
  }, [saveHistory]);
  useEffect(() => {
    const timer = setTimeout(() => {
      if (!writeLocal('sessions', saveHistory ? sessions.slice(0, 60) : []))
        notify('本机存储空间不足，对话历史未保存');
    }, 400);
    return () => clearTimeout(timer);
  }, [sessions, saveHistory, notify]);
  useEffect(() => {
    if (host === 'word') {
      setAdapter(new WordAdapter());
      return;
    }
    if (editor) setAdapter(new BrowserAdapter(editor, title, docId));
  }, [host, editor, docId]);
  useEffect(() => {
    if (adapter instanceof BrowserAdapter) adapter.title = title;
  }, [title, adapter]);
  const refreshDocument = useCallback(async () => {
    if (!adapter) return;
    try {
      const next = await adapter.snapshot();
      setSnapshot(next);
      setError((current) => (current.startsWith('Word 文档连接') ? '' : current));
    } catch (reason) {
      report(new Error(`Word 文档连接失败：${(reason as Error).message}`));
    }
  }, [adapter, report]);
  useEffect(() => {
    if (!adapter) return;
    void refreshDocument();
    if (adapter.kind === 'browser' && editor) {
      const update = () => void refreshDocument();
      editor.on('update', update);
      editor.on('selectionUpdate', update);
      return () => {
        editor.off('update', update);
        editor.off('selectionUpdate', update);
      };
    }
    const timer = setInterval(() => {
      if (!busyRef.current && document.visibilityState === 'visible') void refreshDocument();
    }, 2500);
    const onSelection = () => {
      if (!busyRef.current) void refreshDocument();
    };
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelection);
    return () => {
      clearInterval(timer);
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
        handler: onSelection,
      });
    };
  }, [adapter, editor, refreshDocument]);
  useEffect(() => {
    if (host !== 'browser' || !editor) return;
    const save = () => {
      const ok = writeLocal('document', { id: docId, title, content: editor.getJSON() });
      setSaved(ok);
      if (!ok) notify('本机存储空间不足，请导出文档以免丢失');
    };
    const timer = setTimeout(save, 500);
    const onUnload = (event: BeforeUnloadEvent) => {
      save();
      if (busyRef.current || !saved) {
        event.preventDefault();
      }
    };
    window.addEventListener('beforeunload', onUnload);
    return () => {
      clearTimeout(timer);
      window.removeEventListener('beforeunload', onUnload);
    };
  }, [editor, editor?.state.doc, title, docId, host, saved, notify]);
  useEffect(() => {
    if (snapshot.documentId && !activeId) {
      const session = sessions.find((s) => s.documentId === snapshot.documentId);
      if (session) setActiveId(session.id);
    }
  }, [snapshot.documentId, activeId, sessions]);
  useEffect(() => {
    if (modal && !modalRef.current?.open) modalRef.current?.showModal();
    else if (!modal) modalRef.current?.close();
  }, [modal]);
  useEffect(
    () => () => {
      abortRef.current?.abort();
      clearTimeout(toastTimer.current);
    },
    [],
  );
  const session = sessions.find((s) => s.id === activeId);
  const messages = session?.messages ?? [];
  const changes = sessions
    .filter((s) => s.documentId === snapshot.documentId)
    .flatMap((s) => s.messages)
    .filter((m) => m.change)
    .map((m) => m.change!)
    .sort((a, b) => b.created - a.created);
  const pendingCount = changes.filter((c) => c.status === 'pending').length;
  const currentProfile = settings.profiles.find((p) => p.id === profileId);
  const consent = Boolean(currentProfile && consentFor === currentProfile.baseUrl);
  const setConsent = (allowed: boolean) =>
    setConsentFor(allowed ? currentProfile?.baseUrl || '' : '');
  useEffect(() => {
    if (autoScroll && messagesRef.current)
      messagesRef.current.scrollTop = messages.length ? messagesRef.current.scrollHeight : 0;
  }, [messages, status, autoScroll]);

  function patchMessage(sessionId: string, id: string, update: (m: Message) => Message) {
    setSessions((previous) =>
      previous.map((s) =>
        s.id === sessionId
          ? {
              ...s,
              updated: Date.now(),
              messages: s.messages.map((m) => (m.id === id ? update(m) : m)),
            }
          : s,
      ),
    );
  }
  function patchChange(id: string, update: Partial<Change>) {
    setSessions((previous) =>
      previous.map((s) => ({
        ...s,
        messages: s.messages.map((m) =>
          m.change?.id === id ? { ...m, change: { ...m.change, ...update } } : m,
        ),
      })),
    );
  }
  function createSession() {
    if (busyRef.current) return;
    const fresh: Session = {
      id: uid(),
      documentId: snapshot.documentId,
      title: '新的对话',
      created: Date.now(),
      updated: Date.now(),
      messages: [],
    };
    setSessions((previous) => [fresh, ...previous]);
    setActiveId(fresh.id);
    setPanel('chat');
    setPrompt('');
    setError('');
    setAutoScroll(true);
    composerRef.current?.focus();
  }
  function stageLocalPlan(plan: EditPlan, base: Snapshot) {
    if (busyRef.current || actionBusy) return;
    if (pendingCount) return report(new Error('请先应用或放弃待审阅的修改'));
    if (base.documentId !== snapshot.documentId) return report(new Error('文档已切换，请重新预览'));
    const change: Change = { id: uid(), plan, base, status: 'pending', created: Date.now() };
    const message: Message = {
      id: uid(),
      role: 'assistant',
      content: '本机工具生成的替换预览，尚未修改文档。',
      change,
      created: Date.now(),
    };
    const current = session?.documentId === base.documentId ? session : undefined;
    const id = current?.id || uid();
    setSessions((previous) =>
      current
        ? previous.map((item) =>
            item.id === id
              ? { ...item, updated: Date.now(), messages: [...item.messages, message] }
              : item,
          )
        : [
            {
              id,
              documentId: base.documentId,
              title: plan.summary.slice(0, 30),
              created: Date.now(),
              updated: Date.now(),
              messages: [message],
            },
            ...previous,
          ],
    );
    setActiveId(id);
    setPanel('changes');
    setError('');
  }
  async function applyChange(change: Change, automatic = false) {
    if (!adapter || (!automatic && busyRef.current)) return;
    if (!automatic) busyRef.current = true;
    setActionBusy(true);
    patchChange(change.id, { status: 'applying', error: undefined });
    try {
      const checkpoint = await adapter.apply(change.plan, change.base);
      checkpoints.current.set(change.id, checkpoint);
      while (checkpoints.current.size > 10)
        checkpoints.current.delete(checkpoints.current.keys().next().value!);
      setLastUndoId(change.id);
      patchChange(change.id, { status: 'applied' });
      notify(`已应用 ${change.plan.operations.length} 处更改`);
      await refreshDocument();
    } catch (reason) {
      patchChange(change.id, { status: 'conflict', error: (reason as Error).message });
    } finally {
      setActionBusy(false);
      if (!automatic) busyRef.current = false;
    }
  }
  async function undoChange(change: Change) {
    if (!adapter || busyRef.current) return;
    const checkpoint = checkpoints.current.get(change.id);
    if (!checkpoint) return notify('此编辑的撤回快照已释放，请使用编辑器自身的撤销');
    busyRef.current = true;
    setActionBusy(true);
    try {
      await adapter.undo(checkpoint);
      patchChange(change.id, { status: 'reverted', error: undefined });
      checkpoints.current.delete(change.id);
      setLastUndoId([...checkpoints.current.keys()].at(-1) || '');
      await refreshDocument();
      notify('已撤回本次更改');
    } catch (reason) {
      report(reason);
    } finally {
      busyRef.current = false;
      setActionBusy(false);
    }
  }
  async function send(text = prompt, permission = false) {
    text = text.trim();
    if (!text || !adapter || busyRef.current) return;
    if (!profileId) {
      setPanel('settings');
      return notify('请先添加模型连接');
    }
    if (!consent && !permission) {
      setModal({ kind: 'consent', prompt: text });
      return;
    }
    if (session && session.documentId !== snapshot.documentId)
      return report(new Error('该会话属于另一份文档，请新建对话'));
    if (pendingCount)
      return report(new Error('请先在变更记录中应用或放弃待审阅的修改，再继续对话'));
    busyRef.current = true;
    setBusy(true);
    setError('');
    setPanel('chat');
    setPrompt('');
    setAutoScroll(true);
    setStatus('正在获取文档上下文');
    const controller = new AbortController();
    abortRef.current = controller;
    const sessionId = session?.id || uid();
    const assistantId = uid();
    let created = false;
    try {
      const base = await (adapter.capture?.() ?? adapter.snapshot());
      if (scope === 'selection' && !base.selection) throw new Error('请先在文档中选中文字');
      setSnapshot(base);
      const userMessage: Message = { id: uid(), role: 'user', content: text, created: Date.now() };
      const assistantMessage: Message = {
        id: assistantId,
        role: 'assistant',
        content: '',
        created: Date.now(),
      };
      const history = [...(session?.messages ?? [])];
      setSessions((previous) =>
        session
          ? previous.map((s) =>
              s.id === sessionId
                ? {
                    ...s,
                    title: s.messages.length ? s.title : text.slice(0, 30),
                    updated: Date.now(),
                    messages: [...s.messages, userMessage, assistantMessage],
                  }
                : s,
            )
          : [
              {
                id: sessionId,
                documentId: base.documentId,
                title: text.slice(0, 30),
                created: Date.now(),
                updated: Date.now(),
                messages: [userMessage, assistantMessage],
              },
              ...previous,
            ],
      );
      setActiveId(sessionId);
      created = true;
      const historyPayload = history
        .filter((m) => !m.error && !m.interrupted && (m.content || m.change))
        .slice(-16)
        .map((m) => ({
          role: m.role,
          content:
            m.content.slice(0, 10000) +
            (m.change ? `\n[编辑结果: ${m.change.status}; ${m.change.plan.summary}]` : ''),
        }));
      await streamChat(
        {
          profileId,
          scope,
          mode,
          document: scope === 'selection' ? { ...base, paragraphs: [] } : base,
          messages: [...historyPayload, { role: 'user', content: text }],
        },
        controller.signal,
        async (event) => {
          if (controller.signal.aborted) return;
          if (event.type === 'status') setStatus(event.message);
          if (event.type === 'delta')
            patchMessage(sessionId, assistantId, (m) => ({
              ...m,
              content: m.content + event.text,
            }));
          if (event.type === 'tool')
            patchMessage(sessionId, assistantId, (m) => ({
              ...m,
              tools: [
                ...(m.tools ?? []).filter((tool) => tool.id !== event.activity.id),
                event.activity,
              ].slice(-5),
            }));
          if (event.type === 'proposal') {
            const change: Change = {
              id: uid(),
              plan: event.plan,
              base,
              status: 'pending',
              created: Date.now(),
            };
            patchMessage(sessionId, assistantId, (m) => ({ ...m, change }));
            if (live) {
              setStatus('正在应用到文档');
              await applyChange(change, true);
            } else setStatus('修改已生成，等待审阅');
          }
        },
      );
    } catch (reason) {
      if (controller.signal.aborted) {
        if (created) patchMessage(sessionId, assistantId, (m) => ({ ...m, interrupted: true }));
      } else {
        if (created)
          patchMessage(sessionId, assistantId, (m) => ({ ...m, error: (reason as Error).message }));
        else {
          report(reason);
          setPrompt(text);
        }
      }
    } finally {
      busyRef.current = false;
      setBusy(false);
      setStatus('');
      abortRef.current = null;
      void refreshDocument();
    }
  }
  function stop() {
    abortRef.current?.abort();
    setStatus('正在停止');
  }
  const renderChange = (change: Change) => (
    <ChangeCard
      key={change.id}
      change={change}
      busy={busy || actionBusy}
      canUndo={lastUndoId === change.id && checkpoints.current.has(change.id)}
      onApply={() => void applyChange(change)}
      onReject={() => patchChange(change.id, { status: 'rejected' })}
      onUndo={() => void undoChange(change)}
    />
  );
  async function exportDocument() {
    if (!editor) return;
    setActionBusy(true);
    try {
      const response = await fetch('/api/documents/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-WordAgent-Client': '1' },
        body: JSON.stringify({ document: editor.getJSON() }),
      });
      if (!response.ok) throw new Error((await response.json()).error || '导出失败');
      download(await response.blob(), `${title.replace(/\.docx$/i, '') || '未命名文档'}.docx`);
      notify('Word 文档已导出');
    } catch (reason) {
      report(reason);
    } finally {
      setActionBusy(false);
    }
  }
  function newDocument() {
    editor?.commands.setContent('<p></p>');
    setDocId(uid());
    setTitle('未命名文档.docx');
    setActiveId('');
    setSnapshot(emptySnapshot);
    checkpoints.current.clear();
    setLastUndoId('');
    setModal(null);
  }
  async function importDocument(file: File) {
    if (!file.name.toLowerCase().endsWith('.docx')) return report(new Error('仅支持 .docx 文件'));
    if (file.size > 12 * 1024 * 1024) return report(new Error('文档不能超过 12 MB'));
    setActionBusy(true);
    setModal(null);
    try {
      const response = await fetch('/api/documents/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream', 'X-WordAgent-Client': '1' },
        body: file,
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error);
      editor?.commands.setContent(
        DOMPurify.sanitize(result.html, {
          FORBID_TAGS: ['img', 'iframe', 'script', 'style', 'object'],
        }),
      );
      setTitle(file.name);
      setDocId(uid());
      setActiveId('');
      setSnapshot(emptySnapshot);
      checkpoints.current.clear();
      setLastUndoId('');
      notify('文档正文已导入，原文件未修改');
    } catch (reason) {
      report(reason);
    } finally {
      setActionBusy(false);
    }
  }
  const headings = snapshot.paragraphs.filter((p) => /^(Heading[1-6]|Title)$/.test(p.style));
  const chars = snapshot.paragraphs.reduce((n, p) => n + p.text.length, 0);
  const panelTitle =
    panel === 'settings' ? '设置' : navItems.find((n) => n.id === panel)?.name || '对话';
  return (
    <div className={`app host-${host} mobile-${mobileView}`}>
      {host === 'browser' && (
        <header className="app-header">
          <div className="brand">
            <img src="/assets/icon-32.png" alt="" />
            <span>WordAgent</span>
            <span className="workspace-label">文档工作台</span>
          </div>
          <div className="mobile-view-toggle" role="tablist" aria-label="工作区视图">
            <button
              role="tab"
              aria-selected={mobileView === 'document'}
              onClick={() => setMobileView('document')}
            >
              <FileText size={14} />
              文档
            </button>
            <button
              role="tab"
              aria-selected={mobileView === 'assistant'}
              onClick={() => setMobileView('assistant')}
            >
              <Sparkles size={14} />
              助手
            </button>
          </div>
          <div className="app-header-right">
            <span className="local-badge">
              <span />
              本地工作区
            </span>
            <IconButton
              label="设置"
              onClick={() => {
                setPanel('settings');
                setMobileView('assistant');
              }}
            >
              <Settings2 size={17} />
            </IconButton>
          </div>
        </header>
      )}
      <div className="workbench">
        {host === 'browser' && editor && (
          <EditorWorkspace
            editor={editor}
            title={title}
            setTitle={setTitle}
            saved={saved}
            chars={chars}
            fileRef={fileRef}
            onNew={() => setModal({ kind: 'newDocument' })}
            onExport={() => void exportDocument()}
            onImport={(file) => setModal({ kind: 'import', file })}
            busy={busy || actionBusy}
          />
        )}
        <aside className="assistant-shell" aria-label="WordAgent 侧边栏">
          <nav className="activity-rail" aria-label="侧边栏导航">
            <div className="rail-top">
              {host === 'word' && (
                <img className="rail-brand" src="/assets/icon-32.png" alt="WordAgent" />
              )}
              {navItems.map(({ id, name, icon: Icon }) => (
                <button
                  key={id}
                  className={`rail-button ${panel === id ? 'selected' : ''}`}
                  aria-label={name}
                  title={name}
                  aria-current={panel === id ? 'page' : undefined}
                  onClick={() => {
                    setPanel(id);
                    setSearch('');
                  }}
                >
                  <Icon size={20} />
                  {id === 'changes' && pendingCount > 0 && (
                    <span className="rail-badge">{pendingCount}</span>
                  )}
                </button>
              ))}
            </div>
            <div className="rail-bottom">
              <button
                className={`rail-button ${panel === 'settings' ? 'selected' : ''}`}
                aria-label="设置"
                title="设置"
                onClick={() => setPanel('settings')}
              >
                <Settings2 size={20} />
              </button>
              <div className="avatar" title="本地用户">
                W
              </div>
            </div>
          </nav>
          <section className="assistant-panel">
            <header className="panel-header">
              <div>
                <h2>{panelTitle}</h2>
                {panel === 'chat' && <span className="panel-badge">AI</span>}
              </div>
              <div>
                <IconButton label="新建对话" disabled={busy || actionBusy} onClick={createSession}>
                  <Plus size={18} />
                </IconButton>
                {panel === 'chat' ? (
                  <IconButton label="打开历史会话" onClick={() => setPanel('history')}>
                    <History size={16} />
                  </IconButton>
                ) : (
                  <IconButton label="返回对话" onClick={() => setPanel('chat')}>
                    <MessageSquare size={16} />
                  </IconButton>
                )}
                {host === 'browser' && (
                  <IconButton
                    className="mobile-only"
                    label="显示文档"
                    onClick={() => setMobileView('document')}
                  >
                    <PanelRightClose size={16} />
                  </IconButton>
                )}
              </div>
            </header>
            {error && (
              <div className="error-banner" role="alert">
                <CircleAlert size={15} />
                <span>{error}</span>
                <IconButton label="关闭错误提示" onClick={() => setError('')}>
                  <X size={14} />
                </IconButton>
              </div>
            )}
            {panel === 'chat' && (
              <>
                <div className="context-strip">
                  <button onClick={() => setContextOpen(!contextOpen)} aria-expanded={contextOpen}>
                    <FileText size={14} />
                    <span>{snapshot.title || '正在连接文档'}</span>
                    {contextOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                  </button>
                  <span className="context-count">{snapshot.paragraphs.length} 段</span>
                </div>
                {contextOpen && (
                  <div className="context-detail">
                    <div>
                      <span>上下文范围</span>
                      <strong>{scope === 'selection' ? '当前选区' : '文档正文'}</strong>
                    </div>
                    <div>
                      <span>字符数</span>
                      <strong>{scope === 'selection' ? snapshot.selection.length : chars}</strong>
                    </div>
                    <div>
                      <span>同步状态</span>
                      <strong>{snapshot.documentId ? '已连接' : '未连接'}</strong>
                    </div>
                    {snapshot.selection && (
                      <blockquote>
                        {snapshot.selection.slice(0, 300)}
                        {snapshot.selection.length > 300 ? '…' : ''}
                      </blockquote>
                    )}
                  </div>
                )}
                <div
                  className="messages"
                  ref={messagesRef}
                  onScroll={(e) => {
                    const el = e.currentTarget;
                    setAutoScroll(el.scrollHeight - el.scrollTop - el.clientHeight < 100);
                  }}
                  aria-label="对话消息"
                >
                  {!messages.length && (
                    <div className="chat-empty">
                      <div className="assistant-symbol">
                        <Sparkles size={23} />
                      </div>
                      <h3>新的对话</h3>
                      <div className="empty-document">
                        <FileText size={13} />
                        <span>{snapshot.title || '当前文档'}</span>
                      </div>
                      <div className="starter-actions">
                        <button
                          onClick={() => {
                            setMode('agent');
                            setPrompt('请润色文档中的正文段落，保留原意和事实，使表达更简洁自然。');
                            composerRef.current?.focus();
                          }}
                        >
                          <WandSparkles size={16} />
                          <span>润色文字</span>
                          <ChevronRight size={14} />
                        </button>
                        <button
                          onClick={() => {
                            setMode('ask');
                            setPrompt('请检查这份文档的逻辑与结构，指出具体问题并给出建议。');
                            composerRef.current?.focus();
                          }}
                        >
                          <ListTree size={16} />
                          <span>检查逻辑与结构</span>
                          <ChevronRight size={14} />
                        </button>
                        <button
                          onClick={() => {
                            setMode('agent');
                            setPrompt('请为这份文档撰写一段简洁的摘要，插入到文档标题之后。');
                            composerRef.current?.focus();
                          }}
                        >
                          <PencilLine size={16} />
                          <span>撰写摘要</span>
                          <ChevronRight size={14} />
                        </button>
                      </div>
                    </div>
                  )}
                  {messages.map((message) => (
                    <article key={message.id} className={`message message-${message.role}`}>
                      <div className="message-author">
                        {message.role === 'assistant' ? (
                          <Sparkles size={14} />
                        ) : (
                          <span className="user-dot" />
                        )}
                        <strong>{message.role === 'assistant' ? 'WordAgent' : '你'}</strong>
                        <time>
                          {new Date(message.created).toLocaleTimeString('zh-CN', {
                            hour: '2-digit',
                            minute: '2-digit',
                          })}
                        </time>
                      </div>
                      {message.content && (
                        <div className="markdown">
                          {message.role === 'assistant' ? (
                            <Markdown
                              remarkPlugins={[remarkGfm]}
                              components={{
                                a: (props) => (
                                  <a {...props} target="_blank" rel="noopener noreferrer" />
                                ),
                                img: (props) => <span>{props.alt || '[图片]'}</span>,
                              }}
                            >
                              {message.content}
                            </Markdown>
                          ) : (
                            <p>{message.content}</p>
                          )}
                        </div>
                      )}
                      {!!message.tools?.length && (
                        <div className="tool-activity" aria-label="工具执行记录">
                          {message.tools.map((tool) => (
                            <div key={tool.id}>
                              {tool.status === 'completed' ? (
                                <Check size={12} />
                              ) : (
                                <Wrench size={12} />
                              )}
                              <span>{tool.label}</span>
                              <small>
                                {tool.status === 'completed'
                                  ? '完成'
                                  : busy && message.id === messages.at(-1)?.id
                                    ? '处理中'
                                    : '未完成'}
                              </small>
                            </div>
                          ))}
                        </div>
                      )}
                      {message.change && renderChange(message.change)}
                      {message.error && (
                        <div className="message-error">
                          <CircleAlert size={14} />
                          <span>{message.error}</span>
                        </div>
                      )}
                      {message.interrupted && (
                        <p className="interrupted">
                          已停止生成
                          {message.change?.status === 'applied' ? '，已应用的修改保留' : ''}
                        </p>
                      )}
                      {message.role === 'assistant' &&
                        (message.content || message.error || message.interrupted) && (
                          <div className="message-tools">
                            {message.content && (
                              <IconButton
                                label="复制回复"
                                onClick={() =>
                                  void navigator.clipboard
                                    .writeText(message.content)
                                    .then(() => notify('已复制回复'))
                                    .catch(() => notify('无法访问剪贴板'))
                                }
                              >
                                <Copy size={13} />
                              </IconButton>
                            )}
                            {(message.error || message.interrupted) && (
                              <button
                                className="text-button"
                                disabled={busy || actionBusy}
                                onClick={() => {
                                  const index = messages.findIndex((m) => m.id === message.id);
                                  const text = messages
                                    .slice(0, index)
                                    .reverse()
                                    .find((m) => m.role === 'user')?.content;
                                  if (text) void send(text);
                                }}
                              >
                                <RefreshCw size={12} />
                                重试
                              </button>
                            )}
                          </div>
                        )}
                    </article>
                  ))}
                  {busy && (
                    <div className="generation-status" role="status">
                      <LoaderCircle size={14} className="spin" />
                      <span>{status || '正在生成'}</span>
                    </div>
                  )}
                </div>
                {!autoScroll && (
                  <button
                    className="scroll-latest"
                    onClick={() => {
                      setAutoScroll(true);
                      messagesRef.current?.scrollTo({
                        top: messagesRef.current.scrollHeight,
                        behavior: 'smooth',
                      });
                    }}
                  >
                    <ArrowDown size={13} />
                    最新消息
                  </button>
                )}
                <form
                  className="composer-area"
                  onSubmit={(e) => {
                    e.preventDefault();
                    void send();
                  }}
                >
                  <div className="composer-options">
                    <div className="segmented" role="group" aria-label="对话模式">
                      <button
                        type="button"
                        className={mode === 'agent' ? 'active' : ''}
                        onClick={() => setMode('agent')}
                        disabled={busy}
                      >
                        <Zap size={12} />
                        Agent
                      </button>
                      <button
                        type="button"
                        className={mode === 'ask' ? 'active' : ''}
                        onClick={() => setMode('ask')}
                        disabled={busy}
                      >
                        <MessageSquare size={12} />
                        Ask
                      </button>
                    </div>
                    <label className="live-setting">
                      <input
                        type="checkbox"
                        aria-label="实时应用"
                        checked={live}
                        disabled={busy || actionBusy}
                        onChange={(event) => {
                          setLive(event.target.checked);
                          writeLocal('liveEditing', event.target.checked);
                        }}
                      />
                      <span>{live ? '实时应用' : '审阅后应用'}</span>
                    </label>
                  </div>
                  <div className="composer-box">
                    <div className="composer-context">
                      <FileText size={12} />
                      <select
                        aria-label="上下文范围"
                        value={scope}
                        onChange={(e) => setScope(e.target.value as 'document' | 'selection')}
                        disabled={busy}
                      >
                        <option value="document">当前文档</option>
                        <option value="selection">
                          当前选区
                          {snapshot.selection
                            ? ` · ${snapshot.selection.length} 字符`
                            : ' · 未选择'}
                        </option>
                      </select>
                    </div>
                    <textarea
                      ref={composerRef}
                      aria-label="发送给 AI 的消息"
                      placeholder={
                        mode === 'agent'
                          ? '描述你想对文档做的修改…'
                          : '关于这份文档，你想了解什么？'
                      }
                      value={prompt}
                      maxLength={30000}
                      onChange={(e) => setPrompt(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
                          e.preventDefault();
                          void send();
                        }
                      }}
                      disabled={actionBusy}
                    />
                    <div className="composer-bottom">
                      <select
                        className="model-select"
                        aria-label="选择模型"
                        value={profileId}
                        onChange={(e) => setProfileId(e.target.value)}
                        disabled={busy || !settings.profiles.length}
                      >
                        {!settings.profiles.length && <option value="">未配置模型</option>}
                        {settings.profiles.map((p) => (
                          <option key={p.id} value={p.id}>
                            {p.name}
                          </option>
                        ))}
                      </select>
                      {busy ? (
                        <IconButton
                          label="停止生成"
                          className="send-button stop-button"
                          onClick={stop}
                        >
                          <Square size={14} fill="currentColor" />
                        </IconButton>
                      ) : (
                        <button
                          type="submit"
                          className="send-button"
                          aria-label="发送消息"
                          title="发送消息"
                          disabled={!prompt.trim() || actionBusy || !snapshot.documentId}
                        >
                          <ArrowUp size={18} />
                        </button>
                      )}
                    </div>
                  </div>
                  {mode === 'agent' && (
                    <div className={`approval-status ${live ? 'live' : ''}`}>
                      <ShieldCheck size={11} />
                      <span>{live ? '生成的修改将自动写入文档' : '修改需经你审阅后应用'}</span>
                    </div>
                  )}
                </form>
              </>
            )}
            {panel === 'tools' && (
              <ToolsPanel
                key={snapshot.documentId}
                snapshot={snapshot}
                busy={busy || actionBusy}
                onRefresh={() => void refreshDocument()}
                onFocus={(id) => {
                  void adapter?.focus(id).catch(report);
                }}
                onStage={stageLocalPlan}
                onRecipe={(recipe) => {
                  setMode(recipe.mode);
                  setPrompt(recipe.prompt);
                  setPanel('chat');
                }}
              />
            )}
            {panel === 'outline' && (
              <div className="panel-scroll outline-panel">
                <div className="section-label">
                  <span>当前文档</span>
                  <IconButton
                    label="刷新文档"
                    onClick={() => void refreshDocument()}
                    disabled={busy}
                  >
                    <RefreshCw size={14} />
                  </IconButton>
                </div>
                <div className="document-summary">
                  <FileText size={22} />
                  <strong>{snapshot.title}</strong>
                  <span>
                    {snapshot.paragraphs.length} 段落 · {chars.toLocaleString()} 字符
                  </span>
                </div>
                <div className="section-label">
                  大纲 <span>{headings.length}</span>
                </div>
                {headings.length ? (
                  headings.map((p) => (
                    <button
                      className="outline-item"
                      style={{ paddingLeft: 10 + (Number(p.style.slice(-1)) || 1) * 12 }}
                      key={p.id}
                      onClick={() => void adapter?.focus(p.id).catch(report)}
                    >
                      <span className="heading-level">H{Number(p.style.slice(-1)) || 1}</span>
                      <span>{p.text || '空标题'}</span>
                    </button>
                  ))
                ) : (
                  <p className="muted empty-inline">文档中暂无标题</p>
                )}
                <div className="section-label">
                  当前选区
                  <TextCursorInput size={14} />
                </div>
                <blockquote className="selection-preview">
                  {snapshot.selection || '无选中内容'}
                </blockquote>
              </div>
            )}
            {panel === 'changes' && (
              <div className="panel-scroll changes-panel">
                <div className="section-label">
                  <span>{changes.length} 批修改</span>
                  {pendingCount > 0 && (
                    <button
                      className="text-button"
                      disabled={busy || actionBusy}
                      onClick={() =>
                        changes
                          .filter((c) => c.status === 'pending')
                          .forEach((c) => patchChange(c.id, { status: 'rejected' }))
                      }
                    >
                      放弃全部待审阅
                    </button>
                  )}
                </div>
                {changes.length ? (
                  changes.map(renderChange)
                ) : (
                  <div className="panel-empty">
                    <FileCheck2 size={30} />
                    <h3>暂无变更</h3>
                  </div>
                )}
              </div>
            )}
            {panel === 'history' && (
              <div className="panel-scroll history-panel">
                <div className="search-field">
                  <Search size={14} />
                  <input
                    aria-label="搜索历史会话"
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="搜索会话"
                  />
                </div>
                <div className="section-label">
                  <span>当前文档</span>
                  <div>
                    <IconButton
                      label="导出对话历史"
                      onClick={() =>
                        download(
                          new Blob(
                            [
                              JSON.stringify(
                                sessions.filter((s) => s.documentId === snapshot.documentId),
                                null,
                                2,
                              ),
                            ],
                            { type: 'application/json' },
                          ),
                          'wordagent-conversations.json',
                        )
                      }
                    >
                      <Download size={14} />
                    </IconButton>
                    <IconButton
                      label="清空全部历史"
                      disabled={busy || actionBusy}
                      onClick={() => setModal({ kind: 'clearHistory' })}
                    >
                      <Trash2 size={14} />
                    </IconButton>
                  </div>
                </div>
                {sessions
                  .filter(
                    (s) =>
                      s.documentId === snapshot.documentId &&
                      (s.title + s.messages.map((m) => m.content).join(''))
                        .toLowerCase()
                        .includes(search.toLowerCase()),
                  )
                  .map((s) => (
                    <div className={`session-row ${s.id === activeId ? 'active' : ''}`} key={s.id}>
                      <button
                        disabled={busy || actionBusy}
                        onClick={() => {
                          setActiveId(s.id);
                          setPanel('chat');
                          setAutoScroll(true);
                        }}
                      >
                        <MessageSquare size={15} />
                        <span>
                          <strong>{s.title}</strong>
                          <small>
                            {new Date(s.updated).toLocaleDateString('zh-CN')} ·{' '}
                            {s.messages.filter((m) => m.role === 'user').length} 条消息
                          </small>
                        </span>
                      </button>
                      <IconButton
                        label={`删除会话 ${s.title}`}
                        disabled={busy || actionBusy}
                        onClick={() => setModal({ kind: 'deleteSession', id: s.id })}
                      >
                        <Trash2 size={14} />
                      </IconButton>
                    </div>
                  ))}
                {!sessions.some((s) => s.documentId === snapshot.documentId) && (
                  <div className="panel-empty">
                    <History size={30} />
                    <h3>暂无会话</h3>
                  </div>
                )}
              </div>
            )}
            {panel === 'settings' && (
              <SettingsPanel
                settings={settings}
                refresh={refreshSettings}
                theme={theme}
                setTheme={setTheme}
                saveHistory={saveHistory}
                setSaveHistory={setSaveHistory}
                consent={consent}
                setConsent={setConsent}
                notify={notify}
              />
            )}
            <footer className="panel-status">
              <span>
                <i className={snapshot.documentId ? 'status-dot' : 'status-dot offline'} />
                {host === 'word' ? 'Word 已连接' : '浏览器文档'}
              </span>
              <button onClick={() => void refreshSettings()} title="刷新服务连接">
                <i className={connected ? 'status-dot' : 'status-dot offline'} />
                {connected ? '服务已连接' : '服务未连接'}
              </button>
            </footer>
          </section>
        </aside>
      </div>
      {toast && (
        <div className="toast" role="status">
          <Check size={15} />
          <span>{toast}</span>
          <IconButton label="关闭通知" onClick={() => setToast('')}>
            <X size={14} />
          </IconButton>
        </div>
      )}
      <dialog
        ref={modalRef}
        className="dialog"
        onCancel={() => setModal(null)}
        onClick={(e) => {
          if (e.target === modalRef.current) setModal(null);
        }}
      >
        {modal && (
          <>
            <div className="dialog-heading">
              <h2>
                {modal.kind === 'consent'
                  ? '发送文档上下文'
                  : modal.kind === 'newDocument'
                    ? '新建文档'
                    : modal.kind === 'import'
                      ? '导入 Word 文档'
                      : '删除历史记录'}
              </h2>
              <IconButton label="关闭对话框" onClick={() => setModal(null)}>
                <X size={18} />
              </IconButton>
            </div>
            <p>
              {modal.kind === 'consent'
                ? `本次请求将把${scope === 'selection' ? '选中的文字' : '文档正文'}与对话发送给 ${currentProfile?.name || '所选模型'}（${currentProfile ? new URL(currentProfile.baseUrl).host : ''}）。API Key 仅保存在本机服务端。`
                : modal.kind === 'newDocument'
                  ? '当前浏览器文档将被替换。需要保留时，请先导出为 .docx。'
                  : modal.kind === 'import'
                    ? '导入将替换当前浏览器文档。浏览器仅保留正文、标题、列表和简单表格；图片、批注、修订、页眉页脚与复杂排版不会保留。原文件不会被修改。'
                    : '此操作只删除本机对话记录，不会撤销已经写入文档的修改。删除后无法恢复。'}
            </p>
            <div className="dialog-actions">
              <button className="button subtle" onClick={() => setModal(null)}>
                取消
              </button>
              <button
                className={`button ${modal.kind === 'deleteSession' || modal.kind === 'clearHistory' ? 'danger' : 'primary'}`}
                onClick={() => {
                  if (modal.kind === 'consent') {
                    setConsent(true);
                    setModal(null);
                    void send(modal.prompt, true);
                  }
                  if (modal.kind === 'newDocument') newDocument();
                  if (modal.kind === 'import') void importDocument(modal.file);
                  if (modal.kind === 'deleteSession') {
                    setSessions((prev) => prev.filter((s) => s.id !== modal.id));
                    if (activeId === modal.id) setActiveId('');
                    setModal(null);
                  }
                  if (modal.kind === 'clearHistory') {
                    setSessions([]);
                    setActiveId('');
                    setModal(null);
                  }
                }}
              >
                {modal.kind === 'consent'
                  ? '同意并发送'
                  : modal.kind === 'newDocument'
                    ? '新建空白文档'
                    : modal.kind === 'import'
                      ? '导入正文'
                      : '确认删除'}
              </button>
            </div>
          </>
        )}
      </dialog>
    </div>
  );
}
