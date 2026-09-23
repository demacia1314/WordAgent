import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import request from 'supertest';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import mammoth from 'mammoth';
import { createApp } from '../server/app';
import { exportDocx, parseEditorDocument } from '../server/documents';

const app = createApp();
let temp: string;
const headers = { 'X-WordAgent-Client': '1', Host: 'localhost' };
beforeAll(async () => {
  temp = await fs.mkdtemp(path.join(os.tmpdir(), 'wordagent-tests-'));
  process.env.WORDAGENT_DATA_DIR = temp;
});
afterAll(async () => {
  delete process.env.WORDAGENT_DATA_DIR;
  await fs.rm(temp, { recursive: true, force: true });
});
afterEach(() => vi.unstubAllGlobals());
describe('local API and DOCX round trip', () => {
  it('requires an explicit same-site client header', async () => {
    expect((await request(app).get('/api/settings').set('Host', 'localhost')).status).toBe(403);
  });
  it('blocks foreign origins and DNS-rebound hosts', async () => {
    expect(
      (await request(app).get('/api/settings').set(headers).set('Origin', 'https://evil.example'))
        .status,
    ).toBe(403);
    expect(
      (await request(app).get('/api/settings').set(headers).set('Host', 'evil.example')).status,
    ).toBe(403);
  });
  it('returns a health check and handles missing profiles', async () => {
    expect((await request(app).get('/api/health').set(headers)).body).toMatchObject({
      ok: true,
      version: '2.2.0',
    });
    expect((await request(app).get('/api/settings').set(headers)).body.profiles).toEqual([]);
  });
  it('stores keys on the server only and preserves an omitted key on update', async () => {
    const config = {
      name: 'Local',
      baseUrl: 'http://localhost:11434/v1',
      model: 'test',
      apiKey: 'SECRET_VALUE',
    };
    expect(
      (await request(app).put('/api/settings/profiles/local').set(headers).send(config)).status,
    ).toBe(200);
    const result = await request(app).get('/api/settings').set(headers);
    expect(result.body.profiles[0].hasKey).toBe(true);
    expect(JSON.stringify(result.body)).not.toContain('SECRET_VALUE');
    expect(result.body.profiles[0].apiKey).toBeUndefined();
    const { apiKey, ...withoutKey } = config;
    await request(app).put('/api/settings/profiles/local').set(headers).send(withoutKey);
    expect(
      JSON.parse(await fs.readFile(path.join(temp, 'settings.json'), 'utf8')).profiles[0].apiKey,
    ).toBe(apiKey);
  });
  it('rejects malformed model settings and chat requests', async () => {
    expect(
      (await request(app).put('/api/settings/profiles/bad').set(headers).send({ name: 'bad' }))
        .status,
    ).toBe(400);
    expect((await request(app).post('/api/chat').set(headers).send({})).status).toBe(400);
  });
  it('tests plain connectivity without tools and Agent capability with a no-document function probe', async () => {
    const config = {
      name: 'Probe',
      baseUrl: 'http://localhost:11434/v1',
      model: 'probe',
      apiKey: 'SECRET_VALUE',
    };
    await request(app).put('/api/settings/profiles/probe').set(headers).send(config);
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ choices: [{ message: { content: 'OK' } }] }), {
          status: 200,
        }),
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            choices: [
              {
                message: {
                  tool_calls: [
                    {
                      id: 'call_1',
                      type: 'function',
                      function: { name: 'wordagent_capability_probe', arguments: '{}' },
                    },
                  ],
                },
              },
            ],
          }),
          { status: 200 },
        ),
      );
    vi.stubGlobal('fetch', fetcher);
    const plain = await request(app).post('/api/settings/profiles/probe/test').set(headers);
    expect(plain.status).toBe(200);
    expect(JSON.parse(fetcher.mock.calls[0][1].body).tools).toBeUndefined();
    const agent = await request(app).post('/api/settings/profiles/probe/test?agent=1').set(headers);
    expect(agent.status).toBe(200);
    expect(agent.body.agentTools).toBe(true);
    const body = JSON.parse(fetcher.mock.calls[1][1].body);
    expect(body.messages[0].content).not.toContain('document');
    expect(body.tools[0].function.name).toBe('wordagent_capability_probe');
    expect(body.tool_choice.function.name).toBe('wordagent_capability_probe');
    expect(body.parallel_tool_calls).toBe(false);
  });
  it('reports a clear Agent capability failure without exposing upstream bodies', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValueOnce(
        new Response(JSON.stringify({ choices: [{ message: { content: 'no tool' } }] }), {
          status: 200,
        }),
      ),
    );
    const response = await request(app)
      .post('/api/settings/profiles/probe/test?agent=1')
      .set(headers);
    expect(response.status).toBe(502);
    expect(response.body.error).toContain('Agent 编辑可能不可用');
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValueOnce(new Response('secret upstream error', { status: 400 })),
    );
    const rejected = await request(app)
      .post('/api/settings/profiles/probe/test?agent=1')
      .set(headers);
    expect(rejected.status).toBe(502);
    expect(rejected.body.error).toContain('Agent 工具测试失败');
    expect(JSON.stringify(rejected.body)).not.toContain('secret upstream error');
    await request(app).delete('/api/settings/profiles/probe').set(headers);
  });
  it('deletes credentials and adjusts the default', async () => {
    await request(app).delete('/api/settings/profiles/local').set(headers);
    expect((await request(app).get('/api/settings').set(headers)).body).toEqual({
      defaultProfile: '',
      profiles: [],
    });
  });
  it('exports valid Chinese headings, marks, lists, and fixed-width tables', async () => {
    const node = parseEditorDocument({
      type: 'doc',
      content: [
        { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: '中文标题' }] },
        {
          type: 'paragraph',
          content: [{ type: 'text', text: '测试正文', marks: [{ type: 'bold' }] }],
        },
        {
          type: 'table',
          content: [
            {
              type: 'tableRow',
              content: [
                {
                  type: 'tableHeader',
                  content: [{ type: 'paragraph', content: [{ type: 'text', text: '表头' }] }],
                },
              ],
            },
          ],
        },
      ],
    });
    const file = await exportDocx(node);
    expect(file.readUInt32LE(0)).toBe(0x04034b50);
    const roundTrip = await mammoth.convertToHtml({ buffer: file });
    expect(roundTrip.value).toContain('<h1>中文标题</h1>');
    expect(roundTrip.value).toContain('<strong>测试正文</strong>');
    expect(roundTrip.value).toContain('<table>');
    expect(roundTrip.value).toContain('表头');
    const imported = await request(app)
      .post('/api/documents/import')
      .set(headers)
      .set('Content-Type', 'application/octet-stream')
      .send(file);
    expect(imported.status).toBe(200);
    expect(imported.body.html).toContain('中文标题');
    expect(imported.body.warnings.length).toBeGreaterThan(0);
  });
  it('rejects invalid archives and deeply nested editor input', async () => {
    expect(
      (
        await request(app)
          .post('/api/documents/import')
          .set(headers)
          .set('Content-Type', 'application/octet-stream')
          .send(Buffer.from('not a word document'))
      ).status,
    ).toBe(400);
    let node: object = { type: 'paragraph' };
    for (let i = 0; i < 30; i++) node = { type: 'doc', content: [node] };
    expect(() => parseEditorDocument(node)).toThrow('嵌套过深');
  });
});
