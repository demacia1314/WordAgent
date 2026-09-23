import { spawn } from 'node:child_process';
import { createServer } from 'vite';
import net from 'node:net';
import { randomUUID } from 'node:crypto';

const apiPort = Number(process.env.API_PORT || 4318);
await new Promise<void>((resolve, reject) => {
  const probe = net.createServer();
  probe.once('error', () =>
    reject(new Error(`API port ${apiPort} is occupied. Set API_PORT to a free port.`)),
  );
  probe.listen(apiPort, '127.0.0.1', () => probe.close(() => resolve()));
});
const instance = randomUUID();

const backend = spawn(process.execPath, ['--watch', '--import', 'tsx', 'server/index.ts'], {
  stdio: 'inherit',
  env: { ...process.env, WORDAGENT_INSTANCE: instance },
  windowsHide: true,
});
let server: Awaited<ReturnType<typeof createServer>> | undefined;
let stopping = false;
async function stop(code = 0) {
  if (stopping) return;
  stopping = true;
  if (process.platform === 'win32' && backend.pid) {
    await new Promise<void>((resolve) => {
      const taskkill = spawn('taskkill', ['/PID', String(backend.pid), '/T', '/F'], {
        stdio: 'ignore',
        windowsHide: true,
      });
      taskkill.once('exit', () => resolve());
      taskkill.once('error', () => resolve());
    });
  } else backend.kill();
  await server?.close();
  process.exit(code);
}
backend.on('exit', (code) => {
  if (!stopping) void stop(code || 1);
});
process.on('SIGINT', () => void stop());
process.on('SIGTERM', () => void stop());
try {
  let ready = false;
  for (let attempt = 0; attempt < 60 && !stopping; attempt++) {
    try {
      const response = await fetch(`http://127.0.0.1:${apiPort}/api/health`, {
        headers: { 'X-WordAgent-Client': '1' },
        signal: AbortSignal.timeout(1000),
      });
      const health = (await response.json()) as { instance?: string };
      if (health.instance === instance) {
        ready = true;
        break;
      }
    } catch {
      /* The backend may still be loading TypeScript. */
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  if (!ready) throw new Error('API startup failed. Check the backend error above.');
  server = await createServer();
  await server.listen();
  server.printUrls();
} catch (error) {
  console.error(error);
  await stop(1);
}
