# WordAgent 2.2 Windows 试用包

## 安装

1. 将整个 ZIP 解压到固定的可写目录，例如用户的 Documents 文件夹。不要在 ZIP 内直接双击脚本。
2. 双击 `Install.cmd`，阅读说明后输入 `y`。首次安装会请求信任本机 HTTPS 证书，再通过 Microsoft 开发侧载工具打开 Word。
3. 在 Word 侧边栏设置中填写自己的模型接口和密钥。包内不带作者的模型配置、证书、文档或历史记录。

包内自带 Windows x64 Node 运行时、依赖和生产构建，不需要安装 Node、npm 或开发环境。不是完全离线软件：Office.js 加载和远程 AI 调用需要联网。

## 日常使用

重启电脑后双击 `Start.cmd`，然后打开 Word 的 WordAgent 窗格。侧载入口丢失时重新运行 `Install.cmd`。用 `Stop.cmd` 关闭本包启动的服务；停止后 Word 窗格不可用。服务不注册为系统服务，也不设置开机自启。

端口固定为 3001。端口被占用时会拒绝启动，不会杀掉其他服务。排错日志在 `.local/server-error.log`。目录移动前先停止服务。模型设置保存在本包 `.local/settings.json`，不要把使用过的目录再次打包分发。

## 兼容范围

- 本包仅适用于 Windows x64；Office 本身可为 32 位或 64 位。
- 目标版本为持续更新的 Microsoft 365 Word，以及 Office/LTSC 2021、2024 Word，要求 WordApi 1.3 和现代 WebView2。该 API 检查不是 Office 产品版本识别；旧版本即使能加载，也不属于支持范围。
- 需要合法安装的 Word，以及允许加载项和开发侧载的组织策略。
- 不支持 WPS、Office 2016/2019、Linux 桌面 Office。此 ZIP 不适用于 Mac、ARM64 或网页版。各目标 Office 版本尚未逐一实机验收，不承诺所有更新通道完全一致。

## 卸载

先运行 `Stop.cmd`。在本目录 PowerShell 执行以下命令停止本清单的开发侧载会话，再从 Word 的加载项界面移除 WordAgent：

```powershell
.\runtime\node.exe .\node_modules\office-addin-debugging\lib\cli.js stop manifest.xml desktop
```

之后可删除解压目录；删除前备份需要保留的模型设置。安装用的 localhost 证书由 Microsoft 开发工具共享管理，不自动卸载，避免影响其他 Office 插件。

## 分发边界

这是便捷的本地试用包，不是签名 MSI/EXE，也不是商店正式安装包。证书信任、侧载许可和 Windows 安全提示不能绕过。组织禁用侧载时需管理员批准，不能靠脚本强行安装。

面向大量终端或跨 Windows/Mac/Web 的正式部署，应单独建设 HTTPS 托管服务、身份认证和租户隔离，再由 Microsoft 365 管理员集中分发清单或发布到商店。目前本地 API 只面向单用户，不可直接暴露公网。
