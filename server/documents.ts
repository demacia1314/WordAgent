import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  ShadingType,
  type IRunOptions,
} from 'docx';
import { z } from 'zod';

export type EditorNode = {
  type: string;
  text?: string;
  attrs?: Record<string, unknown>;
  marks?: { type: string; attrs?: Record<string, unknown> }[];
  content?: EditorNode[];
};
const nodeSchema: z.ZodType<EditorNode> = z.lazy(() =>
  z.object({
    type: z.string().max(40),
    text: z.string().max(100000).optional(),
    attrs: z.record(z.unknown()).optional(),
    marks: z
      .array(z.object({ type: z.string(), attrs: z.record(z.unknown()).optional() }))
      .max(20)
      .optional(),
    content: z.array(nodeSchema).max(5000).optional(),
  }),
);
export function parseEditorDocument(input: unknown): EditorNode {
  function depth(value: unknown, level: number) {
    if (level > 24) throw new Error('文档嵌套过深');
    if (value && typeof value === 'object' && 'content' in value && Array.isArray(value.content))
      for (const node of value.content) depth(node, level + 1);
  }
  depth(input, 0);
  return nodeSchema.parse(input);
}
function runs(node: EditorNode): TextRun[] {
  if (node.type === 'hardBreak') return [new TextRun({ break: 1 })];
  if (node.type === 'text') {
    const options: { -readonly [K in keyof IRunOptions]: IRunOptions[K] } = {
      text: node.text || '',
      font: 'Arial',
    };
    for (const mark of node.marks ?? []) {
      if (mark.type === 'bold') options.bold = true;
      if (mark.type === 'italic') options.italics = true;
      if (mark.type === 'underline') options.underline = {};
      if (mark.type === 'strike') options.strike = true;
      if (mark.type === 'textStyle' && mark.attrs?.fontSize)
        options.size = Math.min(144, Math.max(16, parseFloat(String(mark.attrs.fontSize)) * 2));
    }
    return [new TextRun(options)];
  }
  return (node.content ?? []).flatMap(runs);
}
function blocks(
  nodes: EditorNode[],
  numbering?: { reference: string; level: number },
): (Paragraph | Table)[] {
  return nodes.flatMap((node): (Paragraph | Table)[] => {
    if (node.type === 'table') {
      const rows = node.content ?? [];
      const count = rows[0]?.content?.length || 1;
      const widths = Array.from(
        { length: count },
        (_, i) => Math.floor(9026 / count) + (i < 9026 % count ? 1 : 0),
      );
      const border = { style: BorderStyle.SINGLE, size: 4, color: 'D4D8DF' };
      return [
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: widths,
          rows: rows.map(
            (row) =>
              new TableRow({
                children: (row.content ?? []).map(
                  (cell, index) =>
                    new TableCell({
                      width: { size: widths[index] || widths[0], type: WidthType.DXA },
                      columnSpan: Number(cell.attrs?.colspan) || 1,
                      rowSpan: Number(cell.attrs?.rowspan) || 1,
                      borders: { top: border, bottom: border, left: border, right: border },
                      margins: { top: 80, bottom: 80, left: 120, right: 120 },
                      ...(cell.type === 'tableHeader'
                        ? { shading: { type: ShadingType.CLEAR, fill: 'EFF3F6' } }
                        : {}),
                      children: blocks(cell.content ?? []).length
                        ? blocks(cell.content ?? [])
                        : [new Paragraph('')],
                    }),
                ),
              }),
          ),
        }),
      ];
    }
    if (node.type === 'bulletList' || node.type === 'orderedList')
      return blocks(node.content ?? [], {
        reference: node.type === 'bulletList' ? 'bullets' : 'numbers',
        level: numbering ? Math.min(5, numbering.level + 1) : 0,
      });
    if (node.type === 'listItem' || node.type === 'blockquote')
      return blocks(node.content ?? [], numbering);
    const alignment =
      (
        {
          left: AlignmentType.LEFT,
          center: AlignmentType.CENTER,
          right: AlignmentType.RIGHT,
          justify: AlignmentType.JUSTIFIED,
        } as Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]>
      )[String(node.attrs?.textAlign)] || AlignmentType.LEFT;
    const levels = [
      HeadingLevel.HEADING_1,
      HeadingLevel.HEADING_2,
      HeadingLevel.HEADING_3,
      HeadingLevel.HEADING_4,
      HeadingLevel.HEADING_5,
      HeadingLevel.HEADING_6,
    ];
    return [
      new Paragraph({
        children: runs(node),
        alignment,
        numbering,
        ...(node.type === 'heading'
          ? { heading: levels[Math.min(5, Math.max(0, Number(node.attrs?.level || 1) - 1))] }
          : {}),
        spacing: { after: 160, line: 360 },
        ...(node.type === 'horizontalRule'
          ? { border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D4D8DF' } } }
          : {}),
      }),
    ];
  });
}
export async function exportDocx(node: EditorNode): Promise<Buffer> {
  const doc = new Document({
    creator: 'WordAgent',
    styles: {
      default: { document: { run: { font: 'Arial', size: 24 } } },
      paragraphStyles: [1, 2, 3].map((level) => ({
        id: `Heading${level}`,
        name: `Heading ${level}`,
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { bold: true, color: '20252B', size: [36, 30, 26][level - 1] },
        paragraph: { outlineLevel: level - 1, spacing: { before: 240, after: 180 } },
      })),
    },
    numbering: {
      config: ['bullets', 'numbers'].map((reference) => ({
        reference,
        levels: Array.from({ length: 6 }, (_, level) => ({
          level,
          format: reference === 'bullets' ? LevelFormat.BULLET : LevelFormat.DECIMAL,
          text: reference === 'bullets' ? '\u2022' : `%${level + 1}.`,
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
        })),
      })),
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        children: blocks(node.content ?? []),
      },
    ],
  });
  return Packer.toBuffer(doc);
}
