import https from 'node:https';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
const ca = fs.readFileSync(path.join(os.homedir(), '.office-addin-dev-certs/ca.crt'));
let ready = false;
for (let n = 0; n < 30; n++) {
  ready = await new Promise(resolve => {
    const req = https.get({ hostname: 'localhost', port: 3001, path: '/api/health', ca,
      headers: { 'X-WordAgent-Client': '1' } }, res => {
      let body = '';
      res.on('data', c => { body += c; });
      res.on('end', () => { try { resolve(res.statusCode === 200 && JSON.parse(body).ok === true); } catch { resolve(false); } });
      res.on('error', () => resolve(false));
    });
    req.setTimeout(1000, () => req.destroy());
    req.on('error', () => resolve(false));
  });
  if (ready) break;
  await new Promise(resolve => setTimeout(resolve, 300));
}
process.exitCode = ready ? 0 : 1;
