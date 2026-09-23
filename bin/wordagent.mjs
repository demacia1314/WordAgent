#!/usr/bin/env node
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { once } from 'node:events';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import https from 'node:https';
import net from 'node:net';
import readline from 'node:readline/promises';

const root = fileURLToPath(new URL('../', import.meta.url));
const require = createRequire(import.meta.url);
const command = process.argv[2] || '--help';
const version = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).version;
const dataDir = process.env.WORDAGENT_DATA_DIR || path.join(os.homedir(), '.wordagent');
const env = { ...process.env, NODE_ENV: 'production', PORT: '3001', WORDAGENT_DATA_DIR: dataDir,
  PATH: path.dirname(process.execPath) + path.delimiter + process.env.PATH };
const children = new Set();
let stopping = false;
for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, () => {
  stopping = true;
  for (const child of children) child.kill();
});

function launch(file, args = []) {
  const child = spawn(process.execPath, [file, ...args], { cwd: root, env, stdio: 'inherit', windowsHide: true });
  children.add(child);
  child.once('exit', () => children.delete(child));
  return child;
}
async function tool(name, args) {
  const child = launch(require.resolve(`${name}/cli.js`), args);
  const [code] = await once(child, 'exit');
  if (code !== 0 || stopping) throw new Error(`${name} did not finish successfully.`);
}
async function certificateConsent() {
  if (!process.stdin.isTTY) throw new Error('Run install in an interactive terminal to approve the localhost certificate.');
  console.log('This will trust a localhost HTTPS certificate for your account and sideload WordAgent into Word.');
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  let answer;
  try { answer = await rl.question('Continue? [y/N] '); } finally { rl.close(); }
  return answer.trim().toLowerCase() === 'y';
}
async function ready(child) {
  const ca = fs.readFileSync(path.join(os.homedir(), '.office-addin-dev-certs/ca.crt'));
  for (let n = 0; n < 40 && !stopping; n++) {
    if (child.exitCode !== null || child.signalCode) throw new Error('Local server exited before startup.');
    const ok = await new Promise(resolve => {
      const request = https.get({ hostname: 'localhost', port: 3001, path: '/api/health', ca,
        headers: { 'X-WordAgent-Client': '1' } }, response => {
        let body = '';
        response.on('data', value => { body += value; });
        response.on('end', () => { try { resolve(response.statusCode === 200 && JSON.parse(body).ok === true); } catch { resolve(false); } });
        response.on('error', () => resolve(false));
      });
      request.setTimeout(1000, () => request.destroy());
      request.on('error', () => resolve(false));
    });
    if (ok) return;
    await new Promise(resolve => setTimeout(resolve, 250));
  }
  throw new Error('HTTPS startup failed. Run wordagent install to configure certificates.');
}
async function start(install) {
  if (install && process.platform !== 'win32') throw new Error('Automatic sideload is supported by this release on Windows only.');
  await new Promise((resolve, reject) => {
    const probe = net.createServer();
    probe.once('error', () => reject(new Error('Port 3001 is occupied. Stop the other server first.')));
    probe.listen(3001, 'localhost', () => probe.close(resolve));
  });
  if (install) {
    if (!await certificateConsent()) return;
    await tool('office-addin-dev-certs', ['install', '--days', '365']);
  }
  const child = launch(path.join(root, 'dist-server/index.js'));
  const ended = once(child, 'exit');
  // Attach a handler immediately, including failures before the readiness poll completes.
  ended.catch(() => {});
  try {
    await ready(child);
    if (install) await tool('office-addin-debugging', ['start', 'manifest.xml', 'desktop', '--app', 'word', '--no-debug', '--no-live-reload', '--dev-server-port', '3001']);
    console.log(`WordAgent ${version} ready. Settings: ${dataDir}`);
    console.log('Keep this terminal open. Ctrl+C stops the service.');
    const [code] = await ended;
    if (code && !stopping) throw new Error(`Server exited with code ${code}.`);
  } finally {
    if (child.exitCode === null && !child.signalCode) child.kill();
    await ended.catch(() => {});
  }
}

try {
  if (['--help', '-h', 'help'].includes(command)) {
    console.log(`WordAgent ${version}\n\nCommands:\n  install   Approve certificate setup, open Word and run the local service (Windows)\n  start     Run the local HTTPS service after initial setup\n  --version Print the installed version\n\nRequires Node.js 22+, Internet and a supported Word installation.\nModel settings are stored in ~/.wordagent, not in the npm package.\nThis is a local trial with developer sideloading, not a Store add-in.`);
  } else if (['--version', '-v'].includes(command)) console.log(version);
  else if (command === 'start' || command === 'install') {
    if (Number(process.versions.node.split('.')[0]) < 22) throw new Error('Node.js 22+ is required.');
    await start(command === 'install');
  } else throw new Error(`Unknown command: ${command}. Run wordagent --help.`);
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
}
