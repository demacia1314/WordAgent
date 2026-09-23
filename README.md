# WordAgent 2

嵌入 Microsoft Word 的 AI 编辑侧边栏。活动栏、Agent / Ask 对话、文档大纲、变更记录、历史会话和模型设置组成一个工作区。旧 Python 后端与旧前端已移除；新版统一使用 TypeScript。

## 启动

### npm 安装

需要 Node.js 22+。Windows 上首次使用：

```powershell
npx @demacia1314/wordagent install
```

确认本机 HTTPS 证书后会启动服务并尝试打开 Word 加载项，保持终端打开。以后运行 `npx @demacia1314/wordagent start`，按 Ctrl+C 停止。也可 `npm install -g @demacia1314/wordagent` 后使用 `wordagent install` / `wordagent start`。

npm 版本的模型配置位于用户目录 `.wordagent/settings.json`，升级不会写进或删除项目原有 `.local/settings.json`。安装时不带个人密钥，请在侧边栏填写自己的模型连接。这仍是开发侧载试用方式，不是商店安装；目标为 Windows 上的 Microsoft 365 / Office 2021 / 2024 Word，组织需要允许加载项。首次证书信任必须由使用者确认，不在 npm 安装阶段自动执行。

### 源码启动

需要 Node.js 22+。在项目根目录执行：

```powershell
npm install
npm run certs -- --days 365
npm run dev
```

- 浏览器文档工作台：`https://localhost:3001/`
- Word 任务窗格：`https://localhost:3001/taskpane.html`
- 本机 API：`http://127.0.0.1:4318`，由前端同源代理访问，不对外网监听。

`taskpane.html` 必须由 Word 加载。普通浏览器请打开根路径，不会伪装 Word 连接。

### 发布包安装

在 Windows x64 上执行 `npm run package:release`，会生成带时间戳的 `release/WordAgent-2.1.1-windows-x64-*.zip`。包内自带 Node 运行时、生产依赖和构建文件，接收者无需安装开发环境。解压后双击 `Install.cmd`，后续使用 `Start.cmd` / `Stop.cmd`。

本包通过开发侧载提供本地试用，不是商店安装或签名安装程序。首次安装需要用户确认本机证书信任和侧载操作。目标为 Windows x64 上的 Microsoft 365、Office 2021/2024 Word；必须支持 WordApi 1.3 和现代 WebView2。旧版本不纳入支持范围，Mac/Web 不属于该 ZIP 的交付范围，各 Office 版本尚待逐一实机验收。详见 `docs/distribution.md`。

### 安装到 Word

保持开发服务运行，在另一个终端执行：

```powershell
npm run sideload
```

需要已安装并激活的 Microsoft Word 桌面版，支持 WordApi 1.3 及现代 WebView2。清单为根目录 `manifest.xml`。安装后，在 Word 的“开始”选项卡中打开“AI 助手”。此命令由 Microsoft 的 `office-addin-debugging` 工具提供；受组织策略限制时，使用管理员批准的加载项目录部署清单。

如旧加载项被缓存，关闭原窗格，重新加载根目录清单。清单版本已更新到 `2.1.1.0`，不再使用旧 `frontend/office-addin` 路径。

### 端口与证书

默认前端端口固定为 3001，与清单一致。端口被占用时程序明确退出，不会连接未知服务。浏览器开发可单独指定其他端口：

```powershell
$env:PORT = '3002'
$env:API_PORT = '4319'
npm run dev
```

Word 使用其他前端端口时，需要同时更新清单内的 URL。`HTTPS_KEY` / `HTTPS_CERT` 可以覆盖证书路径。默认优先使用用户目录 `.office-addin-dev-certs`，其次读取根目录 `certs`。证书过期时重新执行 `npm run certs -- --days 365` 并重启服务。纯浏览器本地测试可设置 `HTTP_ONLY=1`；Word 加载项仍须使用 HTTPS。

## 模型连接

从侧边栏的“设置”添加显示名称、Base URL、模型 ID、API Key，并保存。Base URL 包含 `/v1` 等接口前缀，不包含 `/chat/completions`。支持流式 Chat Completions 协议；Agent 模式需要模型支持函数工具调用，Ask 模式不发送工具。支持 `minimal / low / medium / high / xhigh / max` 推理强度；服务商不支持时可选择“服务商默认”。

设置中可测试连接、调整温度与最大输出、切换默认模型和删除密钥。连接测试只发送 `Reply with OK.`，不发送文档，也不保证模型支持工具调用。

模型配置位于 `.local/settings.json`，不会进入前端构建或 Git。首次重构已迁移旧模型配置。远程 HTTP 配置保留用于手动修正，但禁止发送凭据；本机 `localhost` / `127.0.0.1` 模型允许 HTTP。

密钥文件是本机明文配置，不是加密保险库，请保护操作系统账户与文件权限。接口不返回密钥值。前端不在浏览器存储密钥，服务端不记录请求正文或模型密钥。

## 编辑工作流

1. 选择 Agent 编辑模式或 Ask 问答模式，选择模型和“当前文档 / 当前选区”。
2. 输入要求。首次发送前确认向当前模型发送文档上下文；授权绑定接口地址，切换服务商或关闭授权后会重新询问。
3. 实时应用固定开启：完整且通过校验的一批编辑会自动写入文档，不会将未完成的文本片段写入文档。
4. 变更记录保存每批操作与状态。当前打开期间最近一批修改可撤回；后续手工修改会阻止直接撤回，以免覆盖新内容。

支持段落替换、插入、删除，原始选区替换，标题样式、加粗、斜体、字号、对齐，以及简单表格插入。每一批编辑先核对完整文档版本和目标原文，同一段落不能在同一批次重复操作。冲突时重新生成，不会强行覆盖。

Word 选区在请求时用跟踪 Range 固定，切换焦点不会将编辑写到另一个同名选区；刷新后旧选区建议需要重新生成。Word 撤回使用正文 OOXML 快照，浏览器撤回使用完整 ProseMirror 文档快照。快照仅在本次打开期间保留，最多 10 批。Word 原生撤销仍可使用。

## 两种运行环境

| 功能                    | Word 内任务窗格          | 浏览器工作台                 |
| ----------------------- | ------------------------ | ---------------------------- |
| AI 对话、审阅、实时应用 | 当前打开的真实 Word 文档 | 当前富文本文档               |
| 大纲和选区              | Office.js 读取           | Tiptap / ProseMirror 读取    |
| 保存                    | Word 自身保存 / 自动保存 | 浏览器本机存储并可导出 .docx |
| 原有复杂排版            | 未涉及的内容由 Word 保留 | 导入仅保留支持的正文结构     |
| 导入 / 导出             | 使用 Word 自身文件功能   | 正文、标题、列表、简单表格   |

浏览器工作台不是完整 Word 渲染器：导入不会保留图片、页眉页脚、批注、修订、域和复杂排版。原文件不被修改。真实文档应优先直接在 Word 插件中编辑。

## 限制与恢复

- 全文上下文上限约 160,000 个序列化字符、3,000 个段落，超过后应改用选区分段处理。单段最多 30,000 字符。
- 每次最多 60 个操作，表格最多 50 行、12 列；对复杂表格内部、列表和受保护内容的修改不保证跨环境一致，无法安全处理的目标会拒绝。
- 不生成或执行任意 JavaScript，不提供联网检索、自动引用核验或旧版 RAG。
- Word API 不保证跨多个操作的数据库式事务。若 Office 在写入中途报错，侧边栏会明确提示可能部分执行，需立即使用 Word 原生撤销检查恢复。不会谎报整批成功。
- 多人协作时，服务端校验到 Word 最终同步之间仍有并发窗口。重要文档请保留 Word 版本历史；本插件不替代备份。
- 对话历史默认保存在本机浏览器，可关闭或清空；包含发送的文本、回复和变更上下文。清除对话不撤回文档编辑。关闭本机历史后刷新页面将不保留会话。
- 浏览器导入文件最大 12 MB，解压后最大 48 MB。加密或损坏文档会拒绝。

## 质量检查

```powershell
npm run typecheck
npm test
npm run build
npm run manifest:validate
npm audit
```

测试使用隔离临时配置，不读取真实密钥，不调用付费模型。覆盖协议校验、真实 ProseMirror 编辑、流式文本与工具解析、中止/截断、冲突与撤回、API 访问限制、密钥保护、DOCX 往返。

本次验收结果与尚待真实模型验证的项目见 `docs/verification.md`。现有模型的认证或接口错误需要先在设置中修正；本地测试服务仅用于验收，不会作为 AI 返回假结果。

构建后 `npm start` 在 HTTPS 3001 端口同时提供静态页面与 API。先停止开发服务，确保可用证书存在。可通过 `WORDAGENT_DATA_DIR` 指定本机配置目录。

## 目录

```text
src/                  React 侧边栏与浏览器文档工作台
  components/         工具栏、变更审阅、模型设置等界面
  document/           Word 与 ProseMirror 编辑适配器
  lib/                API、会话类型、本机存储
shared/contracts.ts   前后端共用的校验与编辑协议
server/               本机 API、模型流、配置与 DOCX 转换
scripts/dev.ts        开发进程管理
tests/                自动化测试
manifest.xml          Word 加载项清单
docs/architecture.md  架构、信任边界与验证说明
```
