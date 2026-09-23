import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

export function loadTls() {
  const home = path.join(os.homedir(), '.office-addin-dev-certs');
  const resolve = (name: string, override?: string) =>
    override ||
    (fs.existsSync(path.join(home, name)) ? path.join(home, name) : path.join('certs', name));
  try {
    return {
      key: fs.readFileSync(resolve('localhost.key', process.env.HTTPS_KEY)),
      cert: fs.readFileSync(resolve('localhost.crt', process.env.HTTPS_CERT)),
    };
  } catch {
    throw new Error('无法读取 HTTPS 证书。请先运行 npm run certs -- --days 365。');
  }
}
