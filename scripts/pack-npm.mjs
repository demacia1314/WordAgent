import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';

const root = fileURLToPath(new URL('../', import.meta.url));
const require = createRequire(import.meta.url);
const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
const stage = await fs.mkdtemp(path.join(os.tmpdir(), 'wordagent-npm-'));
const bundled = ['office-addin-debugging', 'office-addin-dev-certs'];
const npm = process.env.npm_execpath;
if (!npm) throw new Error('Run this script through npm run package:npm.');

try {
  for (const entry of [...manifest.files, 'README.md']) {
    const relative = entry.replace(/\/$/, '');
    const destination = path.join(stage, relative);
    await fs.mkdir(path.dirname(destination), { recursive: true });
    await fs.cp(path.join(root, relative), destination, { recursive: true });
  }
  // Keep bundling out of the development manifest: npm loops on bundled overrides.
  const published = { ...manifest, bundleDependencies: bundled };
  delete published.scripts;
  delete published.devDependencies;
  delete published.overrides;
  await fs.writeFile(path.join(stage, 'package.json'), JSON.stringify(published, null, 2) + '\n');
  await fs.cp(path.join(root, 'node_modules'), path.join(stage, 'node_modules'),
    { recursive: true, dereference: true });
  const output = execFileSync(process.execPath, [npm, 'pack', stage, '--ignore-scripts', '--json',
    '--pack-destination', root], { cwd: root, encoding: 'utf8', maxBuffer: 20 * 1024 * 1024 });
  const report = JSON.parse(output);
  const zip = JSON.parse(await fs.readFile(require.resolve('adm-zip/package.json'), 'utf8'));
  if (zip.version !== '0.6.1') throw new Error('Review adm-zip version before publishing: ' + zip.version);
  await fs.mkdir(path.join(root, '.local'), { recursive: true });
  await fs.writeFile(path.join(root, '.local/npm-pack-report.json'), JSON.stringify(report, null, 2));
  console.log(`Packed ${report[0].filename} (${report[0].size} bytes)`);
} finally {
  if (path.dirname(stage) !== path.resolve(os.tmpdir()) || !path.basename(stage).startsWith('wordagent-npm-')) {
    throw new Error('Unexpected staging path; cleanup refused.');
  }
  await fs.rm(stage, { recursive: true, force: true });
}
