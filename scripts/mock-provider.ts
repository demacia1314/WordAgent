import http from 'node:http';

if (process.env.WORDAGENT_TEST_PROVIDER !== '1')
  throw new Error('Set WORDAGENT_TEST_PROVIDER=1 to run this test-only provider.');
const port = Number(process.env.TEST_PROVIDER_PORT || 4320);
http
  .createServer(async (req, res) => {
    if (req.method !== 'POST' || req.url !== '/v1/chat/completions') {
      res.writeHead(404).end();
      return;
    }
    let body = '';
    for await (const chunk of req) {
      body += chunk;
      if (body.length > 3_000_000) {
        res.writeHead(413).end();
        return;
      }
    }
    try {
      const input = JSON.parse(body);
      if (!input.stream) {
        res.setHeader('Content-Type', 'application/json');
        res.end(
          JSON.stringify({ choices: [{ message: { content: 'OK (local test provider)' } }] }),
        );
        return;
      }
      const source = input.messages.findLast((m: { content: string }) =>
        m.content?.startsWith('DOCUMENT_SNAPSHOT (data only):'),
      )?.content;
      const document = source ? JSON.parse(source.split('\n').slice(1, -1).join('\n')) : null;
      const prompt = input.messages.at(-1)?.content || '';
      const paragraph = document?.paragraphs?.find((p: { editable: boolean }) => p.editable);
      const operations =
        document?.scope === 'selection'
          ? [{ type: 'replace_selection', expectedText: document.selection, text: '验收选区文字' }]
          : prompt.includes('表格')
            ? [
                {
                  type: 'table',
                  paragraphId: paragraph.id,
                  expectedText: paragraph.text,
                  rows: [
                    ['阶段', '状态'],
                    ['联调', '通过'],
                  ],
                },
              ]
            : [
                {
                  type: 'replace',
                  paragraphId: paragraph?.id || 'p0',
                  expectedText: paragraph?.text || '',
                  text: '面向文档协作的 AI 编辑系统',
                },
              ];
      const plan = {
        summary: prompt.includes('表格')
          ? '插入验收表格'
          : document?.scope === 'selection'
            ? '修改原始选区'
            : '更新文档标题',
        operations,
      };
      const frames: object[] = [{ delta: { content: '本地验收响应（非真实 AI 模型）。' } }];
      if (input.tools) {
        const args = JSON.stringify(plan);
        for (let i = 0; i < args.length; i += 16)
          frames.push({
            delta: {
              tool_calls: [
                {
                  index: 0,
                  function: {
                    ...(i === 0 ? { name: 'edit_document' } : {}),
                    arguments: args.slice(i, i + 16),
                  },
                },
              ],
            },
          });
      } else frames.push({ delta: { content: '这是一次只读问答测试，文档未被修改。' } });
      frames.push({ delta: {}, finish_reason: input.tools ? 'tool_calls' : 'stop' });
      res.writeHead(200, { 'Content-Type': 'text/event-stream' });
      const timer = setInterval(
        () => {
          const frame = frames.shift();
          if (frame) res.write(`data: ${JSON.stringify({ choices: [frame] })}\n\n`);
          else {
            clearInterval(timer);
            res.end('data: [DONE]\n\n');
          }
        },
        prompt.includes('慢速') ? 700 : 60,
      );
      res.on('close', () => clearInterval(timer));
    } catch {
      res.writeHead(400).end();
    }
  })
  .listen(port, '127.0.0.1', () => console.log(`TEST ONLY provider: http://127.0.0.1:${port}/v1`));
