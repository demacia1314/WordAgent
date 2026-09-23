import { cp, mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';

const root = process.cwd();
const pkg = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
if (process.platform !== 'win32' || process.arch !== 'x64')
  throw new Error('Build on Windows x64.');
const npmCli = process.env.npm_execpath;
if (!npmCli) throw new Error('Run npm run package:release.');
const stamp = new Date().toISOString().replace(/[:.]/g, '-');
const out = path.join(root, 'release', `WordAgent-${pkg.version}-windows-x64-${stamp}`);
function run(args: string[], cwd = root) {
  return new Promise<void>((resolve, reject) => {
    const child = spawn(process.execPath, [npmCli!, ...args], {
      cwd,
      stdio: 'inherit',
      windowsHide: true,
    });
    child.once('exit', (code) =>
      code === 0 ? resolve() : reject(new Error(`npm exited ${code}`)),
    );
    child.once('error', reject);
  });
}
await run(['run', 'build']);
await run(['test']);
await run(['run', 'manifest:validate']);
await mkdir(out, { recursive: true });
// Explicit allowlist: no personal settings, documents, certificates or logs.
for (const file of ['manifest.xml']) await cp(path.join(root, file), path.join(out, file));
for (const dir of ['dist', 'dist-server'])
  await cp(path.join(root, dir), path.join(out, dir), { recursive: true });
await cp(path.join(root, 'scripts/release'), out, { recursive: true });
await cp(path.join(root, 'docs/distribution.md'), path.join(out, 'INSTALL.md'));
await mkdir(path.join(out, 'runtime'));
await cp(process.execPath, path.join(out, 'runtime/node.exe'));
await cp(path.join(path.dirname(process.execPath), 'LICENSE'), path.join(out, 'runtime/LICENSE'));
const dependencies = Object.fromEntries(
  [
    'express',
    'mammoth',
    'docx',
    'yauzl',
    'zod',
    'office-addin-debugging',
    'office-addin-dev-certs',
  ].map((name) => [name, pkg.dependencies[name] || pkg.devDependencies[name]]),
);
await writeFile(
  path.join(out, 'package.json'),
  JSON.stringify(
    {
      name: 'wordagent-local-trial',
      version: pkg.version,
      private: true,
      type: 'module',
      scripts: { start: 'node dist-server/index.js' },
      dependencies,
      overrides: { 'adm-zip': pkg.overrides['adm-zip'], qs: pkg.overrides.qs },
    },
    null,
    2,
  ),
);
await run(['install', '--omit=dev', '--no-fund'], out);
await run(['audit', '--omit=dev'], out);
const require = createRequire(import.meta.url);
const Zip = require('adm-zip');
const zip = new Zip();
zip.addLocalFolder(out, path.basename(out));
zip.writeZip(`${out}.zip`);
const checksum = createHash('sha256').update(await readFile(`${out}.zip`)).digest('hex');
await writeFile(`${out}.zip.sha256`, `${checksum}  ${path.basename(out)}.zip\n`);
await new Promise<void>((resolve, reject) => {
  const child = spawn(process.execPath, ['scripts/verify-release.mjs', `${out}.zip`], { cwd: root, stdio: 'inherit', windowsHide: true });
  child.once('error', reject);
  child.once('exit', code => code === 0 ? resolve() : reject(new Error('Package verification failed.')));
});
console.log(`Windows trial package: ${out}.zip`);
