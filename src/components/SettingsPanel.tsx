import { useState } from 'react';
import {
  Check,
  ChevronLeft,
  CirclePlus,
  Eye,
  EyeOff,
  LoaderCircle,
  PlugZap,
  Save,
  ShieldCheck,
  Trash2,
  Wrench,
} from 'lucide-react';
import { api, type PublicSettings } from '../lib/api';
import type { PublicProfile } from '../../shared/contracts';
import { IconButton } from './IconButton';

type Form = {
  id: string;
  name: string;
  baseUrl: string;
  model: string;
  apiKey: string;
  temperature: number;
  maxTokens: number;
  reasoningEffort?: 'minimal' | 'low' | 'medium' | 'high' | 'xhigh' | 'max';
  hasKey: boolean;
};
const empty = (): Form => ({
  id: `model-${Date.now().toString(36)}`,
  name: '',
  baseUrl: 'https://api.openai.com/v1',
  model: '',
  apiKey: '',
  temperature: 0.3,
  maxTokens: 8192,
  reasoningEffort: 'medium',
  hasKey: false,
});
export function SettingsPanel({
  settings,
  refresh,
  theme,
  setTheme,
  saveHistory,
  setSaveHistory,
  consent,
  setConsent,
  notify,
}: {
  settings: PublicSettings;
  refresh: () => Promise<void>;
  theme: string;
  setTheme: (v: string) => void;
  saveHistory: boolean;
  setSaveHistory: (v: boolean) => void;
  consent: boolean;
  setConsent: (v: boolean) => void;
  notify: (s: string) => void;
}) {
  const [form, setForm] = useState<Form | null>(null);
  const [showKey, setShowKey] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState('');
  const [deleting, setDeleting] = useState('');
  const run = async (fn: () => Promise<void>) => {
    setBusy(true);
    setStatus('');
    try {
      await fn();
    } catch (e) {
      setStatus((e as Error).message);
    } finally {
      setBusy(false);
    }
  };
  const edit = ({ warning: _warning, ...p }: PublicProfile) => {
    setForm({ ...p, apiKey: '' });
    setStatus('');
    setShowKey(false);
  };
  if (form)
    return (
      <div className="panel-scroll settings-panel">
        <div className="subheading">
          <IconButton
            label="返回设置"
            onClick={() => {
              setForm(null);
              setStatus('');
            }}
          >
            <ChevronLeft size={17} />
          </IconButton>
          <h3>{settings.profiles.some((p) => p.id === form.id) ? '编辑模型' : '添加模型'}</h3>
        </div>
        <form
          onSubmit={(event) => {
            event.preventDefault();
            void run(async () => {
              const { hasKey, ...payload } = form;
              await api(`/settings/profiles/${form.id}`, {
                method: 'PUT',
                body: JSON.stringify({
                  ...payload,
                  apiKey: form.apiKey || (hasKey ? undefined : ''),
                }),
              });
              await refresh();
              setForm(null);
              notify('模型配置已保存');
            });
          }}
        >
          <label className="field">
            显示名称
            <input
              required
              maxLength={80}
              value={form.name}
              onChange={(e) => setForm({ ...form, name: e.target.value })}
              placeholder="我的模型"
            />
          </label>
          <label className="field">
            Base URL
            <input
              required
              type="url"
              value={form.baseUrl}
              onChange={(e) => setForm({ ...form, baseUrl: e.target.value })}
              placeholder="https://api.example.com/v1"
            />
          </label>
          <label className="field">
            模型 ID
            <input
              required
              value={form.model}
              onChange={(e) => setForm({ ...form, model: e.target.value })}
              placeholder="服务商提供的模型 ID"
            />
          </label>
          <label className="field">
            API Key{' '}
            <span className="input-with-button">
              <input
                autoComplete="off"
                type={showKey ? 'text' : 'password'}
                value={form.apiKey}
                onChange={(e) => setForm({ ...form, apiKey: e.target.value })}
                placeholder={form.hasKey ? '已配置，留空保持不变' : '本机模型可留空'}
              />
              <IconButton
                label={showKey ? '隐藏密钥' : '显示密钥'}
                onClick={() => setShowKey(!showKey)}
              >
                {showKey ? <EyeOff size={16} /> : <Eye size={16} />}
              </IconButton>
            </span>
          </label>
          <label className="field">
            温度{' '}
            <div className="range-field">
              <input
                type="range"
                min="0"
                max="2"
                step="0.1"
                value={form.temperature}
                onChange={(e) => setForm({ ...form, temperature: Number(e.target.value) })}
              />
              <output>{form.temperature.toFixed(1)}</output>
            </div>
          </label>
          <label className="field">
            最大输出 Token
            <input
              type="number"
              min="256"
              max="32768"
              step="256"
              required
              value={form.maxTokens}
              onChange={(e) => setForm({ ...form, maxTokens: Number(e.target.value) })}
            />
          </label>
          <label className="field">
            推理强度
            <select
              value={form.reasoningEffort || ''}
              onChange={(e) =>
                setForm({
                  ...form,
                  reasoningEffort: (e.target.value || undefined) as Form['reasoningEffort'],
                })
              }
            >
              <option value="">服务商默认</option>
              <option value="minimal">Minimal</option>
              <option value="low">Low</option>
              <option value="medium">Medium</option>
              <option value="high">High</option>
              <option value="xhigh">XHigh</option>
              <option value="max">Max</option>
            </select>
          </label>
          {status && (
            <p className="inline-error" role="alert">
              {status}
            </p>
          )}
          <button className="button primary full" type="submit" disabled={busy}>
            {busy ? <LoaderCircle className="spin" size={15} /> : <Save size={15} />}保存模型
          </button>
        </form>
      </div>
    );
  return (
    <div className="panel-scroll settings-panel">
      <div className="section-label">
        <span>模型连接</span>
        <IconButton label="添加模型" onClick={() => setForm(empty())}>
          <CirclePlus size={16} />
        </IconButton>
      </div>
      {!settings.profiles.length && <p className="muted empty-inline">尚未配置模型</p>}
      <div className="profile-list">
        {settings.profiles.map((profile) => (
          <div className="profile-row" key={profile.id}>
            <button className="profile-info" onClick={() => edit(profile)}>
              <span>
                {profile.name}
                {profile.id === settings.defaultProfile && <span className="tiny-tag">默认</span>}
              </span>
              <small>{profile.model}</small>
            </button>
            <div className="profile-tools">
              <IconButton
                label={`测试 ${profile.name}`}
                disabled={busy}
                onClick={() =>
                  void run(async () => {
                    const result = await api<{ latency: number }>(
                      `/settings/profiles/${profile.id}/test`,
                      { method: 'POST' },
                    );
                    setStatus(`连接正常 · ${result.latency} ms`);
                  })
                }
              >
                <PlugZap size={15} />
              </IconButton>
              <IconButton
                label={`测试 Agent 工具 ${profile.name}`}
                disabled={busy}
                onClick={() =>
                  void run(async () => {
                    const result = await api<{ latency: number; agentTools: boolean }>(
                      `/settings/profiles/${profile.id}/test?agent=1`,
                      { method: 'POST' },
                    );
                    setStatus(`Agent 工具正常 · ${result.latency} ms · 未发送文档`);
                  })
                }
              >
                <Wrench size={15} />
              </IconButton>
              <IconButton
                label={`设为默认 ${profile.name}`}
                disabled={busy || profile.id === settings.defaultProfile}
                onClick={() =>
                  void run(async () => {
                    await api('/settings/default', {
                      method: 'PUT',
                      body: JSON.stringify({ id: profile.id }),
                    });
                    await refresh();
                  })
                }
              >
                <Check size={15} />
              </IconButton>
              <IconButton
                label={`删除 ${profile.name}`}
                disabled={busy}
                onClick={() => setDeleting(profile.id)}
              >
                <Trash2 size={15} />
              </IconButton>
            </div>
            {profile.warning && <p className="inline-error">{profile.warning}</p>}
            {deleting === profile.id && (
              <div className="delete-confirm">
                <span>删除此模型及密钥？</span>
                <button
                  className="button danger"
                  disabled={busy}
                  onClick={() =>
                    void run(async () => {
                      await api(`/settings/profiles/${profile.id}`, { method: 'DELETE' });
                      await refresh();
                      setDeleting('');
                    })
                  }
                >
                  删除
                </button>
                <button className="button subtle" onClick={() => setDeleting('')}>
                  取消
                </button>
              </div>
            )}
          </div>
        ))}
      </div>
      {busy && (
        <p className="connection-message">
          <LoaderCircle size={14} className="spin" />
          正在连接
        </p>
      )}
      {status && (
        <p className="connection-message" role="status">
          {status}
        </p>
      )}
      <div className="section-label">工作区</div>
      <label className="setting-row">
        <span>外观</span>
        <select aria-label="外观" value={theme} onChange={(e) => setTheme(e.target.value)}>
          <option value="light">浅色</option>
          <option value="dark">深色</option>
          <option value="system">跟随系统</option>
        </select>
      </label>
      <label className="setting-row">
        <span>在本机保存对话历史</span>
        <input
          className="switch"
          role="switch"
          type="checkbox"
          checked={saveHistory}
          onChange={(e) => setSaveHistory(e.target.checked)}
        />
      </label>
      <div className="section-label">
        <ShieldCheck size={14} />
        隐私与权限
      </div>
      <label className="setting-row">
        <span>允许向所选模型发送文档上下文</span>
        <input
          className="switch"
          role="switch"
          type="checkbox"
          checked={consent}
          onChange={(e) => setConsent(e.target.checked)}
        />
      </label>
      <div className="privacy-facts">
        <div>
          <span>密钥存储</span>
          <strong>本机服务端</strong>
        </div>
        <div>
          <span>对话存储</span>
          <strong>{saveHistory ? '本机浏览器' : '仅本次会话'}</strong>
        </div>
        <div>
          <span>文档发送</span>
          <strong>仅发送请求时</strong>
        </div>
      </div>
      <div className="version-label">
        WordAgent <span>2.2.0</span>
      </div>
    </div>
  );
}
