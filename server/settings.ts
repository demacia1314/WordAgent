import fs from 'node:fs/promises';
import path from 'node:path';
import { profileSchema, type Profile, type PublicProfile } from '../shared/contracts';
import { z } from 'zod';

export type Settings = { defaultProfile: string; profiles: Profile[] };
const settingsPath = () =>
  path.resolve(process.env.WORDAGENT_DATA_DIR || '.local', 'settings.json');
export async function readSettings(): Promise<Settings> {
  try {
    const raw = JSON.parse((await fs.readFile(settingsPath(), 'utf8')).replace(/^\uFEFF/, ''));
    const profiles: Profile[] = raw.profiles.map((p: unknown) =>
      profileSchema
        .omit({ baseUrl: true })
        .extend({ baseUrl: z.string().url().max(500) })
        .parse(p),
    );
    const configuredDefault = profiles.find((p) => p.id === raw.defaultProfile);
    const defaultProfile =
      configuredDefault && profileSchema.safeParse(configuredDefault).success
        ? configuredDefault.id
        : profiles.find((p) => profileSchema.safeParse(p).success)?.id || '';
    return { defaultProfile, profiles };
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT')
      return { defaultProfile: '', profiles: [] };
    throw new Error('本地模型配置无法读取，请检查 .local/settings.json');
  }
}
let writing = Promise.resolve();
export function updateSettings(update: (s: Settings) => Settings): Promise<void> {
  const next = writing.then(async () => {
    const settings = update(await readSettings());
    const dest = settingsPath();
    await fs.mkdir(path.dirname(dest), { recursive: true });
    await fs.writeFile(`${dest}.tmp`, JSON.stringify(settings, null, 2), { mode: 0o600 });
    await fs.rename(`${dest}.tmp`, dest);
  });
  writing = next.catch(() => {});
  return next;
}
export function publicProfile({ apiKey, ...profile }: Profile): PublicProfile {
  const validation = profileSchema.safeParse(profile);
  const url = new URL(profile.baseUrl);
  url.username = '';
  url.password = '';
  url.search = '';
  url.hash = '';
  return {
    ...profile,
    baseUrl: url.toString().replace(/\/$/, ''),
    hasKey: Boolean(apiKey),
    ...(!validation.success ? { warning: '旧接口配置需要更新为 HTTPS 后才能使用' } : {}),
  };
}
export async function resolveProfile(id: string): Promise<Profile> {
  const settings = await readSettings();
  const profile = settings.profiles.find((p) => p.id === (id || settings.defaultProfile));
  if (!profile) throw new Error('请先在设置中添加并选择模型');
  if (!profileSchema.safeParse(profile).success)
    throw new Error('此模型的旧接口地址不安全，请在设置中更新为 HTTPS；本机服务可使用 HTTP');
  return profile;
}
