import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const Zip = require('adm-zip');
const file = process.argv[2];
if (!file) throw new Error('Pass the ZIP path.');
const zip = new Zip(file);
const entries = zip.getEntries();
const root = entries[0].entryName.split('/')[0] + '/';
for (const name of ['Install.cmd', 'Start.cmd', 'Stop.cmd', 'manifest.xml', 'runtime/node.exe', 'runtime/LICENSE', 'dist/index.html', 'dist/taskpane.html', 'dist-server/index.js', 'node_modules/express/package.json', 'node_modules/office-addin-debugging/lib/cli.js']) {
  if (!zip.getEntry(root + name)) throw new Error(`Missing package file: ${name}`);
}
for (const entry of entries) {
  const name = entry.entryName.slice(root.length);
  if (!entry.entryName.startsWith(root) || name.split('/').includes('..')) throw new Error('Unsafe ZIP path.');
  if (/^(\.local|certs|\.git|artifacts|src|tests)(\/|$)/.test(name) || /^\.env($|\.)/.test(name)) throw new Error('Private/development data found in package.');
}
for (const name of ['install.ps1', 'start.ps1', 'stop.ps1', 'health.mjs']) {
  if (zip.readAsText(root + name) !== fs.readFileSync(path.join('scripts/release', name), 'utf8')) throw new Error(`Stale launcher: ${name}`);
}
console.log(`PASS: ${entries.length} ZIP entries; bundled runtime and launchers present; no personal data directories.`);
