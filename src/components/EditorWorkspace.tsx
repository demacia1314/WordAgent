import { EditorContent, type Editor } from '@tiptap/react';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  ChevronRight,
  Download,
  FilePlus2,
  FolderOpen,
  Italic,
  List,
  ListOrdered,
  Redo2,
  Table2,
  Underline,
  Undo2,
  Minus,
  Plus,
  Check,
  FileText,
} from 'lucide-react';
import { IconButton } from './IconButton';
import { useState, type RefObject } from 'react';

export function EditorWorkspace({
  editor,
  title,
  setTitle,
  saved,
  chars,
  fileRef,
  onNew,
  onExport,
  onImport,
  busy,
}: {
  editor: Editor;
  title: string;
  setTitle: (s: string) => void;
  saved: boolean;
  chars: number;
  fileRef: RefObject<HTMLInputElement | null>;
  onNew: () => void;
  onExport: () => void;
  onImport: (file: File) => void;
  busy: boolean;
}) {
  const [zoom, setZoom] = useState(100);
  return (
    <main className="editor-workspace" aria-label="文档编辑器">
      <header className="document-header">
        <div className="document-path">
          <FileText size={17} />
          <span>工作区</span>
          <ChevronRight size={13} />
          <input
            aria-label="文档名称"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            onBlur={() => {
              if (!title.trim()) setTitle('未命名文档.docx');
            }}
          />
        </div>
        <div className="document-actions">
          <span className={`save-indicator ${saved ? '' : 'unsaved'}`}>
            <Check size={13} />
            {saved ? '已保存到本机' : '未保存'}
          </span>
          <IconButton label="新建文档" onClick={onNew} disabled={busy}>
            <FilePlus2 size={17} />
          </IconButton>
          <IconButton
            label="导入 Word 文档"
            onClick={() => fileRef.current?.click()}
            disabled={busy}
          >
            <FolderOpen size={17} />
          </IconButton>
          <button className="button export-button" onClick={onExport} disabled={busy}>
            <Download size={14} />
            <span>导出 .docx</span>
          </button>
        </div>
      </header>
      <input
        className="hidden"
        type="file"
        accept=".docx"
        ref={fileRef}
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) onImport(file);
          e.target.value = '';
        }}
      />
      <div className="editor-toolbar" role="toolbar" aria-label="文档格式">
        <IconButton
          label="撤销"
          disabled={!editor.can().undo()}
          onClick={() => editor.chain().focus().undo().run()}
        >
          <Undo2 size={16} />
        </IconButton>
        <IconButton
          label="重做"
          disabled={!editor.can().redo()}
          onClick={() => editor.chain().focus().redo().run()}
        >
          <Redo2 size={16} />
        </IconButton>
        <i className="toolbar-divider" />
        <select
          aria-label="段落样式"
          value={editor.isActive('heading') ? `h${editor.getAttributes('heading').level}` : 'p'}
          onChange={(e) =>
            e.target.value === 'p'
              ? editor.chain().focus().setParagraph().run()
              : editor
                  .chain()
                  .focus()
                  .setHeading({ level: Number(e.target.value[1]) as 1 | 2 | 3 })
                  .run()
          }
        >
          <option value="p">正文</option>
          <option value="h1">标题 1</option>
          <option value="h2">标题 2</option>
          <option value="h3">标题 3</option>
        </select>
        <i className="toolbar-divider" />
        <IconButton
          label="加粗"
          className={editor.isActive('bold') ? 'active' : ''}
          aria-pressed={editor.isActive('bold')}
          onClick={() => editor.chain().focus().toggleBold().run()}
        >
          <Bold size={16} />
        </IconButton>
        <IconButton
          label="斜体"
          className={editor.isActive('italic') ? 'active' : ''}
          aria-pressed={editor.isActive('italic')}
          onClick={() => editor.chain().focus().toggleItalic().run()}
        >
          <Italic size={16} />
        </IconButton>
        <IconButton
          label="下划线"
          className={editor.isActive('underline') ? 'active' : ''}
          aria-pressed={editor.isActive('underline')}
          onClick={() => editor.chain().focus().toggleUnderline().run()}
        >
          <Underline size={16} />
        </IconButton>
        <i className="toolbar-divider" />
        {(
          [
            ['left', AlignLeft, '左对齐'],
            ['center', AlignCenter, '居中'],
            ['right', AlignRight, '右对齐'],
            ['justify', AlignJustify, '两端对齐'],
          ] as const
        ).map(([align, Icon, label]) => (
          <IconButton
            key={align}
            label={label}
            className={editor.isActive({ textAlign: align }) ? 'active' : ''}
            onClick={() => editor.chain().focus().setTextAlign(align).run()}
          >
            <Icon size={16} />
          </IconButton>
        ))}
        <i className="toolbar-divider" />
        <IconButton
          label="无序列表"
          className={editor.isActive('bulletList') ? 'active' : ''}
          onClick={() => editor.chain().focus().toggleBulletList().run()}
        >
          <List size={16} />
        </IconButton>
        <IconButton
          label="有序列表"
          className={editor.isActive('orderedList') ? 'active' : ''}
          onClick={() => editor.chain().focus().toggleOrderedList().run()}
        >
          <ListOrdered size={16} />
        </IconButton>
        <IconButton
          label="插入表格"
          onClick={() =>
            editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()
          }
        >
          <Table2 size={16} />
        </IconButton>
        {editor.isActive('table') && (
          <select
            aria-label="表格操作"
            value=""
            onChange={(e) => {
              const chain = editor.chain().focus();
              if (e.target.value === 'row') chain.addRowAfter().run();
              if (e.target.value === 'column') chain.addColumnAfter().run();
              if (e.target.value === 'deleteRow') chain.deleteRow().run();
              if (e.target.value === 'deleteColumn') chain.deleteColumn().run();
              if (e.target.value === 'delete') chain.deleteTable().run();
            }}
          >
            <option value="">表格操作</option>
            <option value="row">插入行</option>
            <option value="column">插入列</option>
            <option value="deleteRow">删除行</option>
            <option value="deleteColumn">删除列</option>
            <option value="delete">删除表格</option>
          </select>
        )}
      </div>
      <div className="page-stage">
        <div className="page-ruler" aria-hidden="true">
          <span>0</span>
          <span>2</span>
          <span>4</span>
          <span>6</span>
          <span>8</span>
          <span>10</span>
          <span>12</span>
          <span>14</span>
          <span>16</span>
        </div>
        <div className="document-page" style={{ zoom: zoom / 100 }}>
          <EditorContent editor={editor} />
          <footer className="page-footer">
            <span>WordAgent</span>
            <span>工作文档</span>
          </footer>
        </div>
      </div>
      <footer className="document-status">
        <span>{chars.toLocaleString()} 字符</span>
        <span>中文（简体）</span>
        <div className="zoom-controls">
          <IconButton label="缩小" disabled={zoom <= 60} onClick={() => setZoom(zoom - 10)}>
            <Minus size={13} />
          </IconButton>
          <span>{zoom}%</span>
          <IconButton label="放大" disabled={zoom >= 150} onClick={() => setZoom(zoom + 10)}>
            <Plus size={13} />
          </IconButton>
        </div>
      </footer>
    </main>
  );
}
