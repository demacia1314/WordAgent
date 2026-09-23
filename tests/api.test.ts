import { afterAll, beforeAll, describe, expect, it } from 'vitest';
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
      version: '2.1.1',
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
