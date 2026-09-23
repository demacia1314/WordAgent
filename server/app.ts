import express from 'express';
import mammoth from 'mammoth';
import { chatSchema, profileSchema } from '../shared/contracts';
import { runAgent, headers, upstreamUrl, upstreamError } from './agent';
import { publicProfile, readSettings, resolveProfile, updateSettings } from './settings';
import { exportDocx, parseEditorDocument } from './documents';
import { ZodError } from 'zod';
import { validateDocxArchive } from './archive';

export function createApp() {
  const app = express();
  app.disable('x-powered-by');
  app.use('/api', (req, res, next) => {
    res.set({ 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
    const host = req.hostname;
    if (!['localhost', '127.0.0.1', '[::1]', '::1'].includes(host))
      return void res.status(403).json({ error: '仅允许本机访问' });
    if (req.headers['x-wordagent-client'] !== '1')
      return void res.status(403).json({ error: '缺少客户端验证头' });
    const origin = req.headers.origin;
    if (origin) {
      try {
        if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(origin).hostname))
          throw new Error();
      } catch {
        return void res.status(403).json({ error: '不允许跨站访问' });
      }
    }
    next();
  });
  app.get('/api/health', (_req, res) => {
    res.json({ ok: true, version: '2.2.0', instance: process.env.WORDAGENT_INSTANCE });
  });
  app.post(
    '/api/documents/import',
    express.raw({ type: 'application/octet-stream', limit: '12mb' }),
    async (req, res) => {
      if (
        !Buffer.isBuffer(req.body) ||
        req.body.length < 4 ||
        req.body.readUInt32LE(0) !== 0x04034b50
      )
        return void res.status(400).json({ error: '请选择有效的 .docx 文件' });
      try {
        await validateDocxArchive(req.body);
        const result = await mammoth.convertToHtml(
          { buffer: req.body },
          {
            externalFileAccess: false,
            convertImage: mammoth.images.imgElement(async () => ({
              src: '',
              alt: '[图片仅在 Word 原文档中保留]',
            })),
          },
        );
        res.json({
          html: result.value,
          warnings: [
            '浏览器导入仅保留可编辑的正文结构；页眉页脚、图片、批注、修订和复杂排版不会导入。原文件未修改。',
            ...result.messages.map((m) => m.message),
          ],
        });
      } catch {
        res.status(400).json({ error: '无法读取此文档，请确认文件未加密且为有效的 .docx' });
      }
    },
  );
  app.use(express.json({ limit: '3mb' }));
  app.get('/api/settings', async (_req, res) => {
    const settings = await readSettings();
    res.json({
      defaultProfile: settings.defaultProfile,
      profiles: settings.profiles.map(publicProfile),
    });
  });
  app.put('/api/settings/profiles/:id', async (req, res) => {
    const profile = profileSchema.parse({ ...req.body, id: req.params.id });
    await updateSettings((settings) => {
      const current = settings.profiles.find((p) => p.id === profile.id);
      const saved = {
        ...profile,
        apiKey: profile.apiKey === undefined ? current?.apiKey || '' : profile.apiKey,
      };
      return {
        defaultProfile: settings.defaultProfile || saved.id,
        profiles: [...settings.profiles.filter((p) => p.id !== saved.id), saved],
      };
    });
    res.json({ ok: true });
  });
  app.delete('/api/settings/profiles/:id', async (req, res) => {
    await updateSettings((s) => {
      const profiles = s.profiles.filter((p) => p.id !== req.params.id);
      return {
        profiles,
        defaultProfile:
          s.defaultProfile === req.params.id ? profiles[0]?.id || '' : s.defaultProfile,
      };
    });
    res.json({ ok: true });
  });
  app.put('/api/settings/default', async (req, res) => {
    await resolveProfile(req.body.id);
    await updateSettings((s) => ({ ...s, defaultProfile: req.body.id }));
    res.json({ ok: true });
  });
  app.post('/api/settings/profiles/:id/test', async (req, res) => {
    const profile = await resolveProfile(String(req.params.id));
    const agentProbe = req.query.agent === '1';
    const start = Date.now();
    try {
      const probeName = 'wordagent_capability_probe';
      const response = await fetch(upstreamUrl(profile), {
        method: 'POST',
        headers: headers(profile),
        signal: AbortSignal.timeout(25000),
        body: JSON.stringify({
          model: profile.model,
          messages: [
            {
              role: 'user',
              content: agentProbe
                ? 'Call the provided capability probe tool exactly once with no arguments.'
                : 'Reply with OK.',
            },
          ],
          max_tokens: agentProbe ? 128 : 32,
          stream: false,
          ...(agentProbe
            ? {
                tools: [
                  {
                    type: 'function',
                    function: {
                      name: probeName,
                      description: 'No-op probe used only to verify function tool calling support.',
                      parameters: { type: 'object', additionalProperties: false, properties: {} },
                    },
                  },
                ],
                tool_choice: { type: 'function', function: { name: probeName } },
                parallel_tool_calls: false,
              }
            : {}),
        }),
      });
      if (!response.ok) {
        await response.body?.cancel();
        throw new Error(
          agentProbe
            ? `Agent 工具测试失败：${upstreamError(response.status)}`
            : upstreamError(response.status),
        );
      }
      const payload = (await response.json()) as {
        choices?: Array<{
          message?: { tool_calls?: Array<{ function?: { name?: string; arguments?: string } }> };
        }>;
      };
      if (!payload.choices?.length) throw new Error('接口返回了无效的 Chat Completions 响应');
      if (agentProbe) {
        const calls = payload.choices[0]?.message?.tool_calls;
        if (calls?.length !== 1 || calls[0]?.function?.name !== probeName)
          throw new Error('模型没有返回所要求的函数工具调用；Ask 可用，但 Agent 编辑可能不可用');
        try {
          const args = JSON.parse(calls[0].function?.arguments || '{}');
          if (!args || typeof args !== 'object' || Array.isArray(args) || Object.keys(args).length)
            throw new Error();
        } catch {
          throw new Error('模型返回的函数工具参数无效；Agent 编辑可能不可用');
        }
      }
      res.json({
        ok: true,
        latency: Date.now() - start,
        ...(agentProbe ? { agentTools: true } : {}),
      });
    } catch (error) {
      res.status(502).json({
        error:
          error instanceof Error && !['TypeError', 'TimeoutError'].includes(error.name)
            ? error.message
            : '无法连接模型，请检查网络、代理和接口地址',
      });
    }
  });
  let active = 0;
  app.post('/api/chat', async (req, res) => {
    const input = chatSchema.parse(req.body);
    const profile = await resolveProfile(input.profileId);
    if (active >= 3) return void res.status(429).json({ error: '已有多个生成任务，请稍后再试' });
    active++;
    const abort = new AbortController();
    res.on('close', () => abort.abort());
    res.status(200).set({
      'Content-Type': 'text/event-stream; charset=utf-8',
      Connection: 'keep-alive',
      'X-Accel-Buffering': 'no',
    });
    res.flushHeaders();
    const heartbeat = setInterval(() => {
      if (!res.destroyed) res.write(': heartbeat\n\n');
    }, 10000);
    try {
      for await (const event of runAgent(input, profile, abort.signal)) {
        if (res.destroyed) break;
        res.write(`data: ${JSON.stringify(event)}\n\n`);
      }
    } catch (error) {
      if (!abort.signal.aborted) {
        const message =
          error instanceof Error &&
          !['TypeError', 'TimeoutError', 'AbortError'].includes(error.name)
            ? error.message
            : '模型连接超时或中断，请检查网络后重试';
        res.write(`data: ${JSON.stringify({ type: 'error', message })}\n\n`);
      }
    } finally {
      clearInterval(heartbeat);
      active--;
      res.end();
    }
  });
  app.post('/api/documents/export', async (req, res) => {
    const document = parseEditorDocument(req.body.document);
    const file = await exportDocx(document);
    res
      .set({
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="document.docx"',
      })
      .send(file);
  });
  app.use('/api', (_req, res) => {
    res.status(404).json({ error: 'API 路径不存在' });
  });
  app.use(
    (
      error: Error & { status?: number },
      _req: express.Request,
      res: express.Response,
      _next: express.NextFunction,
    ) => {
      if (res.headersSent) return void res.end();
      const status = error instanceof ZodError ? 400 : error.status || 400;
      res.status(status).json({
        error:
          error instanceof ZodError
            ? `请求参数无效：${error.issues[0]?.path.join('.')} ${error.issues[0]?.message}`
            : status === 413
              ? '文档或请求过大，请缩小内容范围'
              : error.message || '请求失败',
      });
    },
  );
  return app;
}
