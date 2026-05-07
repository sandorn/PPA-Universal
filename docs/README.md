# 开发者 README（面向贡献者）

本文件面向开发与贡献，和根目录 `README.md`（面向使用者）明确分离。

## 必读

1. [../.cursor/rules/AGENTS.md](../.cursor/rules/AGENTS.md)（项目约束）
2. [仓库目录说明.md](仓库目录说明.md)（架构、技术栈、目录职责）
3. [../CONTRIBUTING.md](../CONTRIBUTING.md)（提交流程与 PR 规范）

## 开发环境与常用命令

在仓库根目录执行：

```powershell
dotnet build src/PPA.Layered.sln -c Release
```

仅构建 COM 宿主：

```powershell
dotnet build src/Hosts/PPA.Universal.ComAddIn/PPA.Universal.ComAddIn.csproj -c Release
```

手动重建并注册（推荐）：

```bat
build\rebuild-register.bat
```

代码风格与仓库根目录 `.editorconfig` 一致（C# 为 Tab）。整理格式可在仓库根执行：`dotnet format src/PPA.Layered.sln`（提交前按需运行，避免与未格式化的分支产生大范围 diff）。

## 持续集成

推送到 `main` / `master` 或针对这些分支的 PR 时，会在 Windows 上执行 `dotnet build src/PPA.Layered.sln -c Release`（见仓库 `.github/workflows/ci.yml`）。本地提交前仍建议先执行一次相同构建。

## 分层约定（简版）

- `src/Core/PPA.Core`：抽象、配置、日志、DI 基础
- `src/Core/PPA.Business`：平台无关业务逻辑
- `src/Adapters/PPA.Adapter.*`：平台差异实现（PPT/WPS）
- `src/Hosts/PPA.Universal`：平台检测与集成
- `src/Hosts/PPA.Universal.ComAddIn`：COM 入口与 Ribbon

## PPAConfig.xml（与代码对齐）

- **默认路径**：`%LOCALAPPDATA%\PPA.Universal\PPAConfig.xml`（由 `UniversalBootstrapper` 在初始化 DI 时 `LoadOrCreate`）。
- **顶层节点**（以当前代码为准）：`Table`、`Text`、`Chart`（内含 `RegularFont` / `TitleFont` / `LegendFont`）、`GlassCard`、`Duplicate`（矩阵/线性复制对话框默认值）、`Logging`。
- **批量服务（DI）**：`ITableBatchService`（`TableBatchService`）、`IShapeBatchService`（`ShapeBatchService`）已在 `BusinessServiceExtensions.AddPPABusiness` 中注册。Ribbon「格式化」组提供「全稿三线表」；选中表格的三线表仍用「三线表」按钮。`FormatCurrentSlideTables` / `FormatSelectedTables` 等仍可供代码调用。形状批量删除/复制/改尺寸暂无 Ribbon。
- **默认内容**：首次创建或解析失败重写时，使用 `PPA.Core.Configuration.PPAConfig.GetDefaultXmlContent()` 中的模板；各专题文档中的 XML 片段若与本地文件不一致，以你磁盘上的 `PPAConfig.xml` 为准，片段仅说明字段含义。

## 贡献流程（简版）

1. 先在 Issue 或 PR 描述中说明动机与范围
2. 保持单一职责改动，避免无关重构
3. 至少本地 `dotnet build` 通过；改 Ribbon 需在 PowerPoint/WPS 双测
4. PR 写清：改了什么、为什么、如何验证

## 专题文档

- [项目完善规划.md](项目完善规划.md)
- [功能扩展研究报告.md](功能扩展研究报告.md)
- [三线表相关功能梳理.md](三线表相关功能梳理.md)
- [PPT主题颜色参考.md](PPT主题颜色参考.md)
- [Application变量命名规范.md](Application变量命名规范.md)
- `Office UI Help Files/`（图标/控件参考）
