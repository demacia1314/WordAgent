import { afterEach, describe, expect, it, vi } from 'vitest';
import { parseSse, PublicText, runAgent } from '../server/agent';
import type { ChatRequest, Profile, StreamEvent } from '../shared/contracts';

const profile: Profile = {
  id: 'test',
  name: 'test',
  baseUrl: 'http://localhost:1234/v1',
  model: 'test-model',
  apiKey: 'test-secret',
  temperature: 0.3,
  maxTokens: 4096,
};
const request: ChatRequest = {
  profileId: 'test',
  mode: 'agent',
  scope: 'document',
  document: {
    documentId: 'doc',
    title: 'test',
    revision: 'rev',
    selection: '',
    selectionKey: '',
    paragraphs: [{ id: 'p0', text: '原文', style: 'Normal', editable: true }],
  },
  messages: [{ role: 'user', content: '润色' }],
};
const delta = (data: object) => `data: ${JSON.stringify({ choices: [data] })}\n\n`;
function streaming(text: string, chunk = 7) {
  const bytes = new TextEncoder().encode(text);
  return new ReadableStream<Uint8Array>({
    start(c) {
      for (let i = 0; i < bytes.length; i += chunk) c.enqueue(bytes.slice(i, i + chunk));
      c.close();
    },
  });
}
async function collect(input = request) {
  const result: StreamEvent[] = [];
  for await (const e of runAgent(input, profile, new AbortController().signal)) result.push(e);
  return result;
}
afterEach(() => vi.unstubAllGlobals());
describe('streaming agent', () => {
  it('decodes fragmented UTF-8 and CRLF SSE with heartbeats', async () => {
    const frames = [];
    for await (const frame of parseSse(
      streaming(': keepalive\r\n\r\ndata: 你好\r\n\r\ndata: [DONE]\r\n\r\n', 1),
    ))
      frames.push(frame);
    expect(frames).toEqual(['你好', '[DONE]']);
  });
  it('hides split reasoning envelopes', () => {
    const p = new PublicText();
    expect(
      [
        p.feed('回答<th'),
        p.feed('ink>隐藏'),
        p.feed('内容</thi'),
        p.feed('nk>公开'),
        p.feed('', true),
      ].join(''),
    ).toBe('回答公开');
  });
  it('streams public text but never native reasoning fields', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          new Response(
            streaming(
              delta({ delta: { reasoning_content: 'private', content: '你好' } }) +
                delta({ delta: {}, finish_reason: 'stop' }) +
                'data: [DONE]\n\n',
            ),
          ),
        ),
    );
    const events = await collect();
    expect(events).toContainEqual({ type: 'delta', text: '你好' });
    expect(JSON.stringify(events)).not.toContain('private');
    expect(events.at(-1)?.type).toBe('done');
  });
  it('assembles and validates fragmented tool calls before emitting a plan', async () => {
    const plan = {
      summary: '改写',
      operations: [
        { type: 'replace', paragraphId: 'p0', expectedText: '原文', text: '修改后的内容' },
      ],
    };
    const args = JSON.stringify(plan);
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        new Response(
          streaming(
            delta({
              delta: {
                tool_calls: [
                  { index: 0, function: { name: 'edit_document', arguments: args.slice(0, 30) } },
                ],
              },
            }) +
              delta({
                delta: { tool_calls: [{ index: 0, function: { arguments: args.slice(30) } }] },
              }) +
              delta({ delta: {}, finish_reason: 'tool_calls' }),
          ),
        ),
      ),
    );
    expect(await collect()).toContainEqual({ type: 'proposal', plan });
  });
  it('does not propose truncated or disconnected edits', async () => {
    for (const ending of ['', delta({ delta: {}, finish_reason: 'length' })]) {
      vi.stubGlobal(
        'fetch',
        vi
          .fn()
          .mockResolvedValue(
            new Response(streaming(delta({ delta: { content: 'partial' } }) + ending)),
          ),
      );
      await expect(collect()).rejects.toThrow(/截断|中断/);
    }
  });
  it('redacts upstream authentication bodies', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(new Response('api key secret', { status: 401 })),
    );
    await expect(collect()).rejects.toThrow('认证失败');
  });
  it('does not send tools for Ask mode', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValue(
        new Response(
          streaming(delta({ delta: { content: '解释' } }) + delta({ finish_reason: 'stop' })),
        ),
      );
    vi.stubGlobal('fetch', fetcher);
    await collect({ ...request, mode: 'ask' });
    expect(JSON.parse(fetcher.mock.calls[0][1].body).tools).toBeUndefined();
  });
  it('selection scope sends no unselected document paragraphs', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValue(
        new Response(
          streaming(delta({ delta: { content: '解释' } }) + delta({ finish_reason: 'stop' })),
        ),
      );
    vi.stubGlobal('fetch', fetcher);
    await collect({
      ...request,
      scope: 'selection',
      document: { ...request.document, selection: 'selected' },
    });
    expect(fetcher.mock.calls[0][1].body).not.toContain('原文');
  });
  it('rejects an empty selection before contacting the model', async () => {
    const mock = vi.fn();
    vi.stubGlobal('fetch', mock);
    await expect(collect({ ...request, scope: 'selection' })).rejects.toThrow('选中文字');
    expect(mock).not.toHaveBeenCalled();
  });
});
