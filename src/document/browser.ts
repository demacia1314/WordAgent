import type { Editor, JSONContent } from '@tiptap/react';
import type { Node as ProseNode } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';
import {
  validatePlan,
  type EditPlan,
  type Paragraph,
  type ParagraphStyle,
  type Snapshot,
} from '../../shared/contracts';
import {
  CONFLICT,
  UNDO_CONFLICT,
  fingerprint,
  type Checkpoint,
  type DocumentAdapter,
} from './adapter';

export function paragraphNode(
  text: string,
  style: ParagraphStyle = 'Normal',
  attrs: Record<string, unknown> = {},
  marks?: JSONContent['marks'],
): JSONContent[] {
  return text.split(/\r?\n/).map((line) => ({
    type: style.startsWith('Heading') || style === 'Title' ? 'heading' : 'paragraph',
    attrs: {
      ...attrs,
      ...(style.startsWith('Heading')
        ? { level: Number(style.slice(-1)) }
        : style === 'Title'
          ? { level: 1 }
          : {}),
    },
    ...(line ? { content: [{ type: 'text', text: line, marks }] } : {}),
  }));
}

export class BrowserAdapter implements DocumentAdapter {
  readonly kind = 'browser' as const;
  constructor(
    public editor: Editor,
    public title: string,
    public documentId: string,
  ) {}
  entries() {
    const entries: { paragraph: Paragraph; pos: number; node: ProseNode }[] = [];
    this.editor.state.doc.descendants((node, pos, parent) => {
      if (!node.isTextblock) return;
      entries.push({
        paragraph: {
          id: `p${entries.length}`,
          text: node.textContent,
          style: node.type.name === 'heading' ? `Heading${node.attrs.level}` : 'Normal',
          alignment: node.attrs.textAlign || 'left',
          fontFamily:
            node.firstChild?.marks.find((mark) => mark.type.name === 'textStyle')?.attrs.fontFamily ||
            null,
          editable: parent?.type.name === 'doc',
        },
        pos,
        node,
      });
    });
    return entries;
  }
  async snapshot(): Promise<Snapshot> {
    const { from, to } = this.editor.state.selection;
    return {
      documentId: this.documentId,
      title: this.title,
      revision: fingerprint(JSON.stringify(this.editor.getJSON())),
      paragraphs: this.entries().map((e) => e.paragraph),
      selection: this.editor.state.doc.textBetween(from, to, '\n'),
      selectionKey: `${from}:${to}`,
    };
  }
  async apply(plan: EditPlan, base: Snapshot): Promise<Checkpoint> {
    const current = await this.snapshot();
    const sameParagraphState =
      JSON.stringify(current.paragraphs) === JSON.stringify(base.paragraphs) &&
      current.selection === base.selection;
    if (
      current.documentId !== base.documentId ||
      (current.revision !== base.revision && !sameParagraphState)
    )
      throw new Error(CONFLICT);
    validatePlan(plan, base);
    const before = this.editor.getJSON();
    const { schema } = this.editor;
    const entries = this.entries();
    const tr = this.editor.state.tr;
    const ordered = [...plan.operations].sort(
      (a, b) =>
        ('paragraphId' in b ? Number(b.paragraphId.slice(1)) : 0) -
        ('paragraphId' in a ? Number(a.paragraphId.slice(1)) : 0),
    );
    for (const op of ordered) {
      if (op.type === 'replace_selection') {
        const [from, to] = base.selectionKey.split(':').map(Number);
        if (tr.doc.textBetween(from, to, '\n') !== op.expectedText) throw new Error(CONFLICT);
        if (op.text.includes('\n'))
          tr.replaceWith(
            from,
            to,
            paragraphNode(op.text).map((n) => schema.nodeFromJSON(n)),
          );
        else tr.insertText(op.text, from, to);
        continue;
      }
      const entry = entries.find((e) => e.paragraph.id === op.paragraphId)!;
      const { pos, node } = entry;
      if (op.type === 'delete') tr.delete(pos, pos + node.nodeSize);
      if (op.type === 'replace') {
        if (!op.text.includes('\n')) {
          const replacement = op.text ? schema.text(op.text, node.firstChild?.marks) : null;
          if (replacement) tr.replaceWith(pos + 1, pos + node.nodeSize - 1, replacement);
          else tr.delete(pos + 1, pos + node.nodeSize - 1);
        } else
          tr.replaceWith(
            pos,
            pos + node.nodeSize,
            paragraphNode(
              op.text,
              entry.paragraph.style as ParagraphStyle,
              node.attrs,
              node.firstChild?.marks.map((m) => m.toJSON()),
            ).map((n) => schema.nodeFromJSON(n)),
          );
      }
      if (op.type === 'insert')
        tr.insert(
          op.position === 'before' ? pos : pos + node.nodeSize,
          paragraphNode(op.text, op.style).map((n) => schema.nodeFromJSON(n)),
        );
      if (op.type === 'table') {
        const table = schema.nodeFromJSON({
          type: 'table',
          content: op.rows.map((row, index) => ({
            type: 'tableRow',
            content: row.map((cell) => ({
              type: index === 0 ? 'tableHeader' : 'tableCell',
              content: paragraphNode(cell),
            })),
          })),
        });
        tr.insert(pos + node.nodeSize, table);
      }
      if (op.type === 'format') {
        if (op.style || op.alignment) {
          const heading = op.style?.startsWith('Heading') || op.style === 'Title';
          tr.setNodeMarkup(
            pos,
            op.style ? schema.nodes[heading ? 'heading' : 'paragraph'] : undefined,
            {
              ...node.attrs,
              ...(op.style ? { level: heading ? Number(op.style.slice(-1)) || 1 : null } : {}),
              ...(op.alignment ? { textAlign: op.alignment } : {}),
            },
          );
        }
        for (const type of ['bold', 'italic'] as const) {
          if (op[type] === true)
            tr.addMark(pos + 1, pos + node.nodeSize - 1, schema.marks[type].create());
          if (op[type] === false)
            tr.removeMark(pos + 1, pos + node.nodeSize - 1, schema.marks[type]);
        }
        if (op.fontSize)
          tr.addMark(
            pos + 1,
            pos + node.nodeSize - 1,
            schema.marks.textStyle.create({ fontSize: `${op.fontSize}pt` }),
          );
        if (op.fontFamily)
          tr.addMark(
            pos + 1,
            pos + node.nodeSize - 1,
            schema.marks.textStyle.create({ fontFamily: op.fontFamily }),
          );
      }
    }
    if (!tr.doc.childCount) tr.insert(0, schema.nodes.paragraph.create());
    this.editor.view.dispatch(tr.scrollIntoView());
    const after = this.editor.getJSON();
    return {
      before,
      after,
      afterRevision: fingerprint(JSON.stringify(after)),
      documentId: this.documentId,
    };
  }
  async undo(checkpoint: Checkpoint) {
    if (
      this.documentId !== checkpoint.documentId ||
      fingerprint(JSON.stringify(this.editor.getJSON())) !== checkpoint.afterRevision
    )
      throw new Error(UNDO_CONFLICT);
    this.editor.commands.setContent(checkpoint.before as JSONContent);
  }
  async focus(paragraphId: string) {
    const entry = this.entries().find((e) => e.paragraph.id === paragraphId);
    if (!entry) throw new Error('段落不存在');
    const tr = this.editor.state.tr
      .setSelection(TextSelection.create(this.editor.state.doc, entry.pos + 1))
      .scrollIntoView();
    this.editor.view.dispatch(tr);
    this.editor.view.focus();
  }
}
