import { wordLiteralSearch } from '../../shared/text-edits';
import { validatePlan, type EditPlan, type Snapshot, type Paragraph } from '../../shared/contracts';
import {
  CONFLICT,
  UNDO_CONFLICT,
  fingerprint,
  type Checkpoint,
  type DocumentAdapter,
} from './adapter';

const alignment = {
  left: 'Left',
  center: 'Centered',
  right: 'Right',
  justify: 'Justified',
} as const;
export class WordAdapter implements DocumentAdapter {
  readonly kind = 'word' as const;
  private documentId: string;
  private queue = Promise.resolve();
  private capturedSelection?: Word.Range;
  private capturedKey = '';
  constructor() {
    this.documentId =
      Office.context.document.settings.get('wordagent.documentId') || crypto.randomUUID();
    Office.context.document.settings.set('wordagent.documentId', this.documentId);
    Office.context.document.settings.saveAsync();
  }
  private serial<T>(callback: () => Promise<T>): Promise<T> {
    const result = this.queue.then(callback);
    this.queue = result.then(
      () => {},
      () => {},
    );
    return result;
  }
  private async read(context: Word.RequestContext) {
    const paragraphs = context.document.body.paragraphs;
    const selection = context.document.getSelection();
    const xml = context.document.body.getOoxml();
    paragraphs.load(
      'items/text,items/styleBuiltIn,items/alignment,items/font/bold,items/font/italic,items/font/size,items/font/name',
    );
    selection.load('text');
    await context.sync();
    const tables = paragraphs.items.map((p) => p.parentTableOrNullObject);
    tables.forEach((t) => t.load('isNullObject'));
    await context.sync();
    const items: Paragraph[] = paragraphs.items.map((p, i) => ({
      id: `p${i}`,
      text: p.text,
      style: p.styleBuiltIn,
      bold: p.font.bold ?? null,
      italic: p.font.italic ?? null,
      fontSize: p.font.size ?? null,
      fontFamily: p.font.name ?? null,
      alignment: p.alignment,
      editable: tables[i].isNullObject,
    }));
    const raw = Office.context.document.url || '';
    let title = '未命名文档.docx';
    try {
      title = decodeURIComponent(raw.split(/[\\/]/).at(-1) || title).split('?')[0];
    } catch {
      /* Unsaved or non-URL document names are valid. */
    }
    const snapshot: Snapshot = {
      title,
      documentId: this.documentId,
      revision: fingerprint(xml.value),
      paragraphs: items,
      selection: selection.text,
      selectionKey: fingerprint(selection.text),
    };
    return { snapshot, paragraphs, selection, xml: xml.value };
  }
  snapshot(): Promise<Snapshot> {
    return this.serial(() => Word.run(async (context) => (await this.read(context)).snapshot));
  }
  capture(): Promise<Snapshot> {
    return this.serial(async () => {
      if (this.capturedSelection) {
        const previous = this.capturedSelection;
        await Word.run(previous, async (context) => {
          previous.untrack();
          await context.sync();
        });
        this.capturedSelection = undefined;
      }
      return Word.run(async (context) => {
        const current = await this.read(context);
        if (current.snapshot.selection) {
          current.selection.track();
          await context.sync();
          this.capturedSelection = current.selection;
          this.capturedKey = crypto.randomUUID();
          current.snapshot.selectionKey = this.capturedKey;
        }
        return current.snapshot;
      });
    });
  }
  apply(plan: EditPlan, base: Snapshot): Promise<Checkpoint> {
    return this.serial(async () => {
      const selectionEdit = plan.operations.some((op) => op.type === 'replace_selection');
      if (selectionEdit && (!this.capturedSelection || base.selectionKey !== this.capturedKey))
        throw new Error('选区定位已失效，请重新选择文字并生成修改');
      const execute = async (context: Word.RequestContext): Promise<Checkpoint> => {
        const current = await this.read(context);
        if (base.documentId !== this.documentId || current.snapshot.revision !== base.revision)
          throw new Error(CONFLICT);
        validatePlan(plan, base);
        const selection = selectionEdit ? this.capturedSelection! : current.selection;
        if (selectionEdit) {
          selection.load('text');
          await context.sync();
          if (selection.text !== base.selection) throw new Error('原始选区已改变，请重新生成修改');
        }
        // Preflight every literal range before queuing ANY write. One search sync,
        // not one round trip per occurrence (Office correlated-objects pattern).
        const replacements = plan.operations
          .filter((op) => op.type === 'replace_text')
          .map((op) => {
            const paragraph = current.paragraphs.items[Number(op.paragraphId.slice(1))];
            const ranges = paragraph.search(wordLiteralSearch(op.find), {
              matchCase: true,
              matchWholeWord: false,
              matchWildcards: false,
              matchPrefix: false,
              matchSuffix: false,
              ignorePunct: false,
              ignoreSpace: false,
            });
            ranges.load('items/text');
            return { op, ranges };
          });
        const replacementRanges = new Map<string, Word.Range[]>();
        if (replacements.length) {
          const checkedXml = context.document.body.getOoxml();
          await context.sync();
          if (fingerprint(checkedXml.value) !== current.snapshot.revision)
            throw new Error(CONFLICT);
          for (const { op, ranges } of replacements) {
            if (
              ranges.items.length !== op.expectedMatches ||
              ranges.items.some((range) => range.text !== op.find)
            )
              throw new Error('Word 匹配结果与原文不一致，未应用本批修改，请重新查找');
            replacementRanges.set(
              op.paragraphId,
              op.occurrence === undefined ? ranges.items : [ranges.items[op.occurrence - 1]],
            );
          }
        }
        // Resolve all targets against the same snapshot and edit from bottom to top.
        const sorted = [...plan.operations].sort(
          (a, b) =>
            ('paragraphId' in b ? Number(b.paragraphId.slice(1)) : 0) -
            ('paragraphId' in a ? Number(a.paragraphId.slice(1)) : 0),
        );
        for (const op of sorted) {
          if (op.type === 'replace_selection') {
            selection.insertText(op.text.replace(/\n/g, '\r'), 'Replace');
            continue;
          }
          const p = current.paragraphs.items[Number(op.paragraphId.slice(1))];
          if (op.type === 'replace_text') {
            for (const range of [...replacementRanges.get(op.paragraphId)!].reverse())
              range.insertText(op.text, 'Replace');
          }
          if (op.type === 'replace')
            p.getRange('Content').insertText(op.text.replace(/\n/g, '\r'), 'Replace');
          if (op.type === 'delete') {
            if (current.paragraphs.items.length === 1) p.getRange('Content').clear();
            else p.delete();
          }
          if (op.type === 'insert') {
            const lines = op.text.split(/\r?\n/);
            let anchor = p;
            for (const line of op.position === 'before' ? [...lines].reverse() : lines) {
              anchor = anchor.insertParagraph(line, op.position === 'before' ? 'Before' : 'After');
              if (op.style) anchor.styleBuiltIn = op.style;
            }
          }
          if (op.type === 'table') {
            const table = p.insertTable(op.rows.length, op.rows[0].length, 'After', op.rows);
            table.styleBuiltIn = 'TableGrid';
            table.headerRowCount = 1;
          }
          if (op.type === 'format') {
            if (op.style) p.styleBuiltIn = op.style;
            if (op.bold !== undefined) p.font.bold = op.bold;
            if (op.italic !== undefined) p.font.italic = op.italic;
            if (op.fontSize) p.font.size = op.fontSize;
            if (op.fontFamily) p.font.name = op.fontFamily;
            if (op.alignment) p.alignment = alignment[op.alignment];
          }
        }
        try {
          await context.sync();
        } catch {
          throw new Error(
            'Word 未能完成编辑，可能有部分操作已执行。请立即使用 Word 的撤销检查并恢复，然后重新生成。',
          );
        }
        const after = context.document.body.getOoxml();
        await context.sync();
        return {
          before: current.xml,
          after: after.value,
          afterRevision: fingerprint(after.value),
          documentId: this.documentId,
        };
      };
      return selectionEdit ? Word.run(this.capturedSelection!, execute) : Word.run(execute);
    });
  }
  undo(checkpoint: Checkpoint): Promise<void> {
    return this.serial(() =>
      Word.run(async (context) => {
        const xml = context.document.body.getOoxml();
        await context.sync();
        if (
          checkpoint.documentId !== this.documentId ||
          fingerprint(xml.value) !== checkpoint.afterRevision
        )
          throw new Error(UNDO_CONFLICT);
        context.document.body.insertOoxml(String(checkpoint.before), 'Replace');
        await context.sync();
      }),
    );
  }
  focus(id: string): Promise<void> {
    return this.serial(() =>
      Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        const p = paragraphs.items[Number(id.slice(1))];
        if (!p) throw new Error('段落不存在，请刷新大纲');
        p.select();
        await context.sync();
      }),
    );
  }
}
