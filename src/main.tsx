import { Component, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './styles.css';

class ErrorBoundary extends Component<{ children: ReactNode }, { error: string }> {
  state = { error: '' };
  static getDerivedStateFromError(error: Error) {
    return { error: error.message };
  }
  render() {
    return this.state.error ? (
      <div className="boot-screen">
        <h2>工作区暂时无法打开</h2>
        <p>{this.state.error}</p>
        <button className="button primary" onClick={() => location.reload()}>
          重新加载
        </button>
      </div>
    ) : (
      this.props.children
    );
  }
}
const root = createRoot(document.getElementById('root')!);
async function start() {
  let host: 'word' | 'browser' = 'browser';
  if (location.pathname.includes('taskpane')) {
    root.render(
      <div className="boot-screen">
        <img src="/assets/icon-80.png" alt="WordAgent" />
        <p>正在连接 Word…</p>
      </div>,
    );
    try {
      if (!globalThis.Office) throw new Error('Office.js 未加载，请检查网络连接');
      const info = await Promise.race([
        Office.onReady(),
        new Promise<never>((_, reject) =>
          setTimeout(
            () => reject(new Error('未连接到 Word。请从 Word 的加载项中打开此窗格。')),
            15000,
          ),
        ),
      ]);
      if (info.host !== Office.HostType.Word) throw new Error('此页面需要在 Microsoft Word 中打开');
      if (!Office.context.requirements.isSetSupported('WordApi', '1.3'))
        throw new Error('当前 Word 版本不支持 WordApi 1.3，请更新 Office');
      host = 'word';
    } catch (error) {
      root.render(
        <div className="boot-screen">
          <img src="/assets/icon-80.png" alt="WordAgent" />
          <h2>尚未连接 Word</h2>
          <p>{(error as Error).message}</p>
          <button className="button primary" onClick={() => location.reload()}>
            重新连接
          </button>
          <a href="/">打开浏览器工作台</a>
        </div>,
      );
      return;
    }
  }
  root.render(
    <ErrorBoundary>
      <App host={host} />
    </ErrorBoundary>,
  );
}
void start();
