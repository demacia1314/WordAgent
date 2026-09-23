import { describe, expect, it } from 'vitest';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import os from 'node:os';
import fs from 'node:fs';

const cli = path.resolve('bin/wordagent.mjs');
const version = JSON.parse(fs.readFileSync('package.json', 'utf8')).version;
const run = (argument: string) => spawnSync(process.execPath, [cli, argument], {
  cwd: os.tmpdir(), encoding: 'utf8', windowsHide: true, timeout: 10000,
});
describe('npm command line entry', () => {
  it('prints help outside the package directory without starting a server', () => {
    const result = run('--help');
    expect(result.status).toBe(0);
    expect(result.stdout).toContain('install');
    expect(result.stdout).toContain('start');
  });
  it('reports the package version', () => {
    const result = run('--version');
    expect(result.status).toBe(0);
    expect(result.stdout.trim()).toBe(version);
  });
  it('rejects unknown commands', () => {
    const result = run('invalid-command');
    expect(result.status).toBe(1);
    expect(result.stderr).toContain('Unknown command');
  });
});
