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
  function toolResponse(name: string, args: unknown, id = 'call_1', finish = 'tool_calls') {
    return new Response(
      streaming(
        delta({
          delta: {
            tool_calls: [{ index: 0, id, function: { name, arguments: JSON.stringify(args) } }],
          },
        }) + delta({ finish_reason: finish }),
      ),
    );
  }
  const finalResponse = () =>
    new Response(
      streaming(delta({ delta: { content: '已检查' } }) + delta({ finish_reason: 'stop' })),
    );
  it('continues a read-tool conversation with the exact call id and tool result', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(toolResponse('search_document', { query: '原文' }, 'search_1'))
      .mockResolvedValueOnce(finalResponse());
    vi.stubGlobal('fetch', fetcher);
    const events = await collect();
    expect(fetcher).toHaveBeenCalledTimes(2);
    const body = JSON.parse(fetcher.mock.calls[1][1].body);
    expect(body.messages.at(-2).tool_calls[0].id).toBe('search_1');
    expect(body.messages.at(-1)).toMatchObject({ role: 'tool', tool_call_id: 'search_1' });
    expect(JSON.parse(body.messages.at(-1).content).totalMatches).toBe(1);
    expect(events.filter((e) => e.type === 'tool')).toHaveLength(2);
    expect(events.some((e) => e.type === 'proposal')).toBe(false);
    expect(events.at(-1)?.type).toBe('done');
  });
  it('searches then compiles one final replacement proposal without a third model call', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(toolResponse('search_document', { query: '原文' }))
      .mockResolvedValueOnce(
        toolResponse(
          'replace_in_document',
          { summary: '更名', find: '原文', text: '新文', expectedOccurrences: 1 },
          'replace_1',
        ),
      );
    vi.stubGlobal('fetch', fetcher);
    const events = await collect();
    expect(events.filter((e) => e.type === 'proposal')).toHaveLength(1);
    const proposal = events.find((e) => e.type === 'proposal');
    expect(proposal?.type === 'proposal' && proposal.plan.operations[0].type).toBe('replace_text');
    expect(fetcher).toHaveBeenCalledTimes(2);
  });
  it('uses one deadline across tool rounds', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(toolResponse('inspect_document', {}))
      .mockResolvedValueOnce(finalResponse());
    vi.stubGlobal('fetch', fetcher);
    await collect();
    expect(fetcher.mock.calls[0][1].signal).toBe(fetcher.mock.calls[1][1].signal);
  });
  it('bounds repeated read tools instead of looping forever', async () => {
    let calls = 0;
    const fetcher = vi
      .fn()
      .mockImplementation(() =>
        Promise.resolve(toolResponse('inspect_document', {}, `read_${calls++}`)),
      );
    vi.stubGlobal('fetch', fetcher);
    await expect(collect()).rejects.toThrow('次数达到上限');
    expect(fetcher).toHaveBeenCalledTimes(5);
  });
  it('rejects duplicate or missing read call ids', async () => {
    const fetcher = vi
      .fn()
      .mockImplementation(() => Promise.resolve(toolResponse('inspect_document', {}, 'duplicate')));
    vi.stubGlobal('fetch', fetcher);
    await expect(collect()).rejects.toThrow('缺失或重复');
    expect(fetcher).toHaveBeenCalledTimes(2);
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(toolResponse('inspect_document', {}, '')));
    await expect(collect()).rejects.toThrow('缺失或重复');
  });
  it('refuses body tools in selection scope even when the provider invents them', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(toolResponse('search_document', { query: '原文' }));
    vi.stubGlobal('fetch', fetcher);
    await expect(
      collect({
        ...request,
        scope: 'selection',
        document: { ...request.document, selection: 'selected' },
      }),
    ).rejects.toThrow('不支持');
    const body = JSON.parse(fetcher.mock.calls[0][1].body);
    expect(body.tools.map((t: { function: { name: string } }) => t.function.name)).toEqual([
      'edit_document',
    ]);
  });
  it('rejects malformed tool inputs, unknown tools, and non-tool completion reasons', async () => {
    for (const response of [
      toolResponse('inspect_document', { script: 'x' }),
      toolResponse('shell', {}),
      toolResponse('inspect_document', {}, 'id', 'stop'),
    ]) {
      vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(response));
      await expect(collect()).rejects.toThrow();
    }
  });
  it('cancels after a read without starting another upstream request', async () => {
    const fetcher = vi.fn().mockResolvedValueOnce(toolResponse('inspect_document', {}));
    vi.stubGlobal('fetch', fetcher);
    const controller = new AbortController();
    const execute = async () => {
      for await (const event of runAgent(request, profile, controller.signal)) {
        if (event.type === 'tool' && event.activity.status === 'completed') controller.abort();
        expect(event.type).not.toBe('proposal');
      }
    };
    await expect(execute()).rejects.toThrow();
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
  it('bounds cumulative multiline SSE data, not just unfinished lines', async () => {
    const execute = async () => {
      for await (const _frame of parseSse(
        streaming(('data: ' + 'x'.repeat(1000) + '\n').repeat(2100), 65536),
      )) {
        /* consume */
      }
    };
    await expect(execute()).rejects.toThrow('响应帧过大');
  });
  it('rejects an empty selection before contacting the model', async () => {
    const mock = vi.fn();
    vi.stubGlobal('fetch', mock);
    await expect(collect({ ...request, scope: 'selection' })).rejects.toThrow('选中文字');
    expect(mock).not.toHaveBeenCalled();
  });
});
