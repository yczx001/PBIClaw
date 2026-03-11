# PBI Claw（Power BI 智能助手）

PBI Claw 是一个面向 Power BI Desktop 的智能建模助手，形态为：

- Power BI `External Tools` 外部工具
- Windows 独立桌面应用（WinForms + WebView2）

它可以连接当前已打开的 Power BI 模型，读取模型与报表元数据，基于上下文与 AI 对话，并在可控前提下执行模型写回与回滚。

## 1. 核心功能

### 1.1 模型连接与识别
- 自动扫描本机已打开的 Power BI Desktop 实例并识别端口
- 支持外部工具参数自动连接（`--server "%server%" --database "%database%" --external-tool`）
- 支持手动连接 `localhost:port` 或 Tabular Server

### 1.2 元数据读取与上下文构建
- 基于 TOM 读取模型元数据：
  - 表、列、度量值、关系、角色、RLS/OLS 信息
  - 计算表/计算列表达式
  - 分区与来源脚本信息
- 读取报表结构信息（PBIX/PBIP）：
  - 页面、视觉对象类型、标题等
- Power Query 元数据增强：
  - 支持已加载与未加载到模型的查询识别
  - 支持从运行时数据库 TMDL、PBIP 目录、PBIX DataMashup 等多来源提取 M 代码
  - 可解析来源系统、服务器、数据库、Schema、对象名

### 1.3 AI 对话（支持流式）
- 支持 OpenAI 兼容接口和 Anthropic 兼容接口
- 支持流式输出、请求取消、连接测试、错误回传
- 基于问题按需加载深度上下文，减少不必要上下文体积
- 聊天输入支持快捷建议，便于追问

### 1.4 变更计划、预检、执行、回滚
- 支持从 AI 回复中解析 `abi_action_plan`
- 执行前自动预检（错误/警告/信息）
- 写回前自动生成：
  - 模型快照
  - 回滚计划（JSON）
- 支持在“备份中心”查看历史并执行回滚

### 1.5 更新与升级
- 支持应用内检测更新（读取 `latest.json`）
- 支持一键下载并启动安装包升级

## 2. 技术架构

### 2.1 技术栈
- 后端宿主：`.NET 8` + `WinForms`
- 前端界面：`WebView2` + 原生 `HTML/CSS/JavaScript`
- 模型访问：`Microsoft.AnalysisServices.NetCore.retail.amd64`（TOM）
- 通信方式：WebView2 双向消息桥（`AppBridge`）

### 2.2 关键模块
- `Program` / `CliOptions`：入口与命令行参数解析
- `PowerBiInstanceDetector`：Power BI 进程与端口发现
- `TabularMetadataReader`：模型快照读取
- `MetadataPromptBuilder`：系统提示词与上下文构建
- `OnDemandMetadataContextBuilder`：按需深度元数据查询
- `PowerQueryMetadataReader` / `QuerySourceParser`：PQ/M 查询与来源解析
- `AiChatClient`：多供应商 AI 请求与流式解析
- `TabularModelWriter` / `TabularExtendedActionHandler`：模型动作预检与写回
- `AbiActionPlanStorage`：回滚计划落盘与读取
- `PbiMetadataInstaller`：安装向导与外部工具注册

### 2.3 TOM 与 TMDL 在本项目中的使用
- **TOM**：实时连接模型，读取/写回对象
- **TMDL**：用于补充查询定义，尤其是未加载到模型的 Power Query 信息
- 策略：优先走实时模型上下文，在需要时按需补充 TMDL/PQ 细节

## 3. 项目结构

```text
.
├─ docs/
│  └─ PBI_CLAW_UI_PRD.md
├─ external-tools/
│  └─ PBIClaw.pbitool.json
├─ src/
│  ├─ PbiMetadataTool/          # 主程序（PBIClaw.exe）
│  └─ PbiMetadataInstaller/     # 安装程序（PBIClawSetup.exe）
├─ tools/
│  ├─ build-setup-exe.ps1
│  ├─ get-build-version.ps1
│  ├─ register-debug-external-tool.ps1
│  └─ register-tool.ps1
└─ dist/
   ├─ publish/                  # 构建中间产物
   └─ setup/                    # 安装包与更新清单产物
```

## 4. 环境要求

### 4.1 运行环境（用户）
- Windows 10/11
- Power BI Desktop（已安装）
- .NET Desktop Runtime 8.x（若机器未预装，首次启动会提示缺失）
- WebView2 Runtime（通常系统已带）

### 4.2 开发环境（开发者）
- .NET SDK 8.x
- PowerShell 5+ 或 PowerShell 7+

## 5. 安装与使用

### 5.1 安装
1. 运行安装包 `PBIClawSetup.exe`
2. 选择安装目录（默认 `Program Files\PBI Claw`）
3. 按提示授权管理员权限
4. 重启 Power BI Desktop

安装器会自动把外部工具配置写入机器级目录（含 x64/x86 Common Files）：

- `%CommonProgramFiles%\Microsoft Shared\Power BI Desktop\External Tools`
- `%CommonProgramFiles(x86)%\Microsoft Shared\Power BI Desktop\External Tools`

### 5.2 使用
1. 打开 PBIX 报表
2. 在 Power BI Desktop 的 `External Tools` 点击 `PBI Claw`
3. 在工具内连接模型
4. 配置 AI 引擎（Base URL、Model、API Key）
5. 开始对话分析，或在允许写回模式下执行变更计划

### 5.3 快速上手（用户视角）
1. 先用“连接模型”确认状态变为已连接（右上角连接状态可见）
2. 在“AI 引擎”里填好 `Base URL / Model / API Key`，点击“测试连接”
3. 第一次建议先问模型总览类问题，例如：
   - `请先概览当前模型：事实表、维度表、关系方向和潜在风险。`
4. 再逐步深入到对象级问题，例如：
   - `解释度量值 [毛利率] 的DAX逻辑，并指出可能的筛选上下文问题。`
5. 若需要改模型，先确认设置为“允许模型变更”，再让 AI 生成计划并预检执行
6. 执行后到“备份中心”确认已生成快照和回滚文件

### 5.4 典型场景举例
1. 模型梳理与新同事交接
   - 你可以问：`按业务域解释当前模型，每张核心表的用途是什么？`
   - 适合场景：新成员接手旧模型、快速熟悉表和关系
2. DAX 排障与优化
   - 你可以问：`[销售成本] 为什么在按月份汇总时结果异常？请定位可能原因并给修复方案。`
   - 适合场景：总计不对、上下文错乱、性能慢
3. 追查数据来源
   - 你可以问：`事实表_销售明细 的数据来源是什么？给出来源系统、数据库、对象名和M代码。`
   - 适合场景：核对口径、排查源端变更
4. 查询未加载到模型的 Power Query
   - 你可以问：`列出未加载到模型的查询，并显示每个查询的M代码与来源信息。`
   - 适合场景：排查“只做中间转换但未加载”的查询
5. 自动生成变更计划
   - 你可以问：`在表 销售明细 新增度量值 [客单价]，并给出格式字符串。`
   - 适合场景：批量新增/调整度量值、关系、描述、角色权限
6. 风险控制与回滚
   - 你可以问：`执行前先预检，只执行前2条动作，其余保留。`
   - 适合场景：生产模型小步变更、可追溯操作

### 5.5 能实现哪些功能（用户可感知）
- 可以做的事
  - 连接并读取当前 Power BI 模型结构
  - 回答表/列/度量值/关系/角色/RLS 相关问题
  - 识别计算表、计算列、度量值定义并参与分析
  - 读取并解释 Power Query（M）代码与来源链路（含未加载查询）
  - 生成结构化变更计划，执行前预检，执行后自动备份并支持回滚
  - 检测新版本并启动升级安装
- 当前边界
  - 不会替代业务确认，业务口径仍需人工最终审核
  - 写回能力受当前连接模型与权限限制，预检不通过不会强制执行
  - AI 结果质量依赖模型上下文完整度与问题描述清晰度
  - 若网关或模型体量较大，响应时间可能变长

## 6. 开发与调试

### 6.1 构建主程序
```powershell
dotnet build .\src\PbiMetadataTool\PbiMetadataTool.csproj -c Debug
```

### 6.2 注册调试外部工具（免重复安装）
```powershell
.\tools\register-debug-external-tool.ps1
```

如果需要写入机器级 External Tools 目录：
```powershell
.\tools\register-debug-external-tool.ps1 -Machine
```

### 6.3 直接启动主程序
```powershell
dotnet run --project .\src\PbiMetadataTool\PbiMetadataTool.csproj
```

## 7. 打包与发布

### 7.1 一键构建安装包
```powershell
.\tools\build-setup-exe.ps1
```

可选参数示例：
```powershell
.\tools\build-setup-exe.ps1 -Configuration Release -Runtime win-x64 -Version 2026.3.319.1813
```

脚本输出：
- 安装包：`dist\setup\PBIClawSetup.exe`
- 更新清单：`dist\setup\latest.json`

### 7.2 发布流程（建议）
1. 运行打包脚本生成 `PBIClawSetup.exe` 与 `latest.json`
2. 上传安装包到下载地址
3. 上传 `latest.json` 到更新清单地址
4. 创建/更新 GitHub Release（可选，但推荐）

> 本项目默认更新清单地址为：`https://pbihub.cn/downloads/PBIClaw/latest.json`

## 8. 命令行参数（主程序）

```text
--list-instances          列出检测到的实例
--instance-index <index>  指定实例下标
--port <port>             手动指定端口
--server <host:port>      指定服务器（如 localhost:xxxxx）
--database <name>         指定数据库名（默认首个数据库）
--out <path>              导出 JSON 路径
--external-tool           外部工具模式运行
--help                    查看帮助
```

无参数启动时默认进入外部工具模式。

## 9. 配置与数据文件

### 9.1 用户配置
- `%AppData%\PBIClaw\settings.json`
- 主要字段：
  - `provider`（`openai` / `anthropic`）
  - `baseUrl`
  - `model`
  - `apiKey`
  - `temperature`
  - `allowModelChanges`
  - `includeHiddenObjects`
  - `quickPrompts`

### 9.2 历史与备份
- 变更历史：`%AppData%\PBIClaw\change-history.json`
- 模型快照/回滚：`%UserProfile%\Documents\PBIClaw\backups`

## 10. 内置动作类型（写回）

核心动作包括但不限于：

- 度量值：`create_or_update_measure`、`delete_measure`
- 关系：`create_relationship`、`delete_relationship`
- 对象重命名：`rename_table`、`rename_column`、`rename_measure`
- 显隐/格式：`set_table_hidden`、`set_column_hidden`、`set_measure_hidden`、`set_format_string`
- 计算对象：`create_calculated_column`、`delete_column`、`create_calculated_table`
- 关系属性：`set_relationship_active`、`set_relationship_cross_filter`
- 描述更新：`update_description`
- 角色/RLS：`create_role`、`update_role`、`delete_role`、`set_role_table_permission` 等

## 11. 性能与上下文策略

- 聊天请求超时：`180s`
- 对话上下文尾部限制：最多 `8` 条消息、约 `12000` 字符
- 初始模型上下文采用“索引模式”（精简），按提问内容触发深度查询
- 对深度 Power Query 读取结果使用缓存（默认 TTL 20 分钟），降低重复查询开销

## 12. 常见问题

### Q1：连接测试成功，但聊天长时间无回复
- 检查模型是否已连接
- 缩小问题范围（指定表名/度量值名）
- 检查 AI 网关是否支持流式与所选端点格式

### Q2：安装提示进程占用
- 先关闭所有 `PBIClaw.exe` 进程后重试安装

### Q3：为什么需要管理员权限
- 安装器会写入机器级 External Tools 目录，属于系统路径

## 13. 说明

- 本项目主要面向 Windows + Power BI Desktop 场景
- 写回模式涉及模型变更，请在测试环境充分验证后再用于生产模型
