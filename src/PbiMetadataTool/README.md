# PBI Claw

Power BI Desktop External Tool + Standalone Assistant:

- 自动连接当前 PBIX 模型（也支持扫描本机已打开报告）
- 全新现代化工作台 UI（侧边导航 + 多页面任务流）
- 浏览表/列/度量值/关系并右键触发快捷分析
- 导出模型元数据到 JSON
- 配置任意 OpenAI 兼容模型并进行上下文对话
- AI 回复支持 Markdown 可读渲染（标题、列表、粗体、行内代码、代码块）
- 支持只读建议/允许变更两种权限模式
- 支持“AI 变更计划预览 -> 动作勾选 -> 执行前预检 -> 人工确认 -> 执行写回”（度量值/关系）
- 每次执行前自动生成模型快照与回滚脚本（JSON），并支持“执行回滚文件”
- 新增“备份中心”可浏览历史快照/回滚文件并直接触发回滚

UI 重构需求文档：`docs/ABI_UI_PRD.md`

## Build Installer EXE

Double-click from repo root:

- `build-setup-exe.bat`
- or `生成安装器EXE.bat`

Output:

- `dist\setup\PBIClawSetup.exe`

## Install

1. Double-click `dist\setup\PBIClawSetup.exe`
2. Restart Power BI Desktop
3. Open `External Tools` and click `PBI Claw v<version>`
4. In the UI, click `连接`，然后即可对话或导出元数据

## Debug Mode (No Reinstall)

One-time register debug external tool:

- double-click `注册调试外部工具.bat`
- if your Power BI only reads machine scope, run as admin: `注册调试外部工具_管理员模式.bat`

First time only:

1. restart Power BI Desktop
2. click `PBI Claw (Debug) v<version>` in External Tools

After that, for each code change:

1. double-click `注册调试外部工具.bat`（会自动更新版本号、图标和菜单名称）
2. click `PBI Claw (Debug) v<version>` again (通常无需重装)

Standalone debug run:

- `调试启动ABI助手.bat`

Machine-wide install:

```powershell
.\dist\setup\PBIClawSetup.exe --machine
```

## Uninstall

Double-click:

- `一键卸载.bat`

Then restart Power BI Desktop.

## Cleanup Intermediates

To remove intermediate files and keep only final setup output (`dist\setup`):

- `清理临时文件.bat`
