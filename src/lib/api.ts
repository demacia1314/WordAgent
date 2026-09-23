import type { ChatRequest, PublicProfile, StreamEvent } from '../../shared/contracts';

export async function api<T>(path: string, options: RequestInit = {}): Promise<T> {
  const response = await fetch(`/api${path}`, {
    ...options,
    headers: { 'Content-Type': 'application/json', 'X-WordAgent-Client': '1', ...options.headers },
  });
  if (!response.ok) {
    let message = `请求失败 (${response.status})`;
    try {
      message = (await response.json()).error || message;
    } catch {
      /* Non-JSON proxy errors. */
    }
    throw new Error(message);
  }
  return response.json();
}
export type PublicSettings = { defaultProfile: string; profiles: PublicProfile[] };
export async function streamChat(
  request: ChatRequest,
  signal: AbortSignal,
  onEvent: (event: StreamEvent) => Promise<void>,
): Promise<void> {
  const response = await fetch('/api/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'X-WordAgent-Client': '1' },
    body: JSON.stringify(request),
    signal,
  });
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || `服务不可用 (${response.status})`);
  }
  if (!response.body) throw new Error('浏览器不支持流式响应');
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let complete = false;
  try {
    while (true) {
      const { value, done } = await reader.read();
      buffer += done ? decoder.decode() : decoder.decode(value, { stream: true });
      let index: number;
      while ((index = buffer.indexOf('\n\n')) >= 0) {
        const frame = buffer.slice(0, index);
        buffer = buffer.slice(index + 2);
        const data = frame
          .split('\n')
          .filter((line) => line.startsWith('data:'))
          .map((line) => line.slice(5).trimStart())
          .join('\n');
        if (!data) continue;
        const event = JSON.parse(data) as StreamEvent;
        if (event.type === 'error') throw new Error(event.message);
        if (event.type === 'done') complete = true;
        await onEvent(event);
      }
      if (done) break;
    }
    if (!complete && !signal.aborted) throw new Error('响应流提前结束，请重试');
  } finally {
    await reader.cancel().catch(() => {});
    reader.releaseLock();
  }
}
