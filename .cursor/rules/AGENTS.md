# PPA-Universal 代理规则（精简）

## 1. 沟通与修改

- 回复简洁，结论优先；说明改动时用 **问题 → 方案 → 状态**。
- 只改任务所需文件，避免无关重构与「顺手整理」。

## 2. 仓库分层（须遵守）

| 层级   | 路径                               | 职责                                                                                 |
| ------ | ---------------------------------- | ------------------------------------------------------------------------------------ |
| 核心   | `src/Core/PPA.Core`                | 抽象接口、配置 `PPAConfig`、日志 `ILogger`（`PPA.Logging`）、`CoreServiceExtensions`   |
| 业务   | `src/Core/PPA.Business`            | 业务接口与实现、`BusinessServiceExtensions`；**不**直接引用 NetOffice / WPS COM      |
| 适配   | `src/Adapters/PPA.Adapter.*`       | 平台相关实现；差异收口在此                                                           |
| 宿主   | `src/Hosts/PPA.Universal`          | 平台检测、DI 引导、`UniversalIntegration`                                            |
| 插件壳 | `src/Hosts/PPA.Universal.ComAddIn` | COM、`PPARibbon.xml`、`RibbonCallbacks`、Ribbon 图标资源                             |

**与宿主相关的其他路径**：`ico/`（Ribbon PNG 源）、`build/`（如 `rebuild-register.bat`）、`docs/`（设计与专题说明，入口 **`docs/README.md`**）。

## 3. 依赖注入

- 注册：`AddPPACore`、`AddPPABusiness`、各 Adapter 的 `Add*Adapter`，在 `UniversalBootstrapper` 中组合（以代码为准）。
- 解析：宿主侧优先 `UniversalIntegration.GetService<T>()` / `ServiceProvider`；**新类型**能构造注入则注入，避免再堆静态单例。

## 4. 日志与异常

- 使用 **`ILogger`**：业务与适配层关键步骤 Info/Warning，异常 **`LogError(..., ex)`**。
- Ribbon 回调中：`try/catch` 内记录日志并给用户简短反馈（与现有 `RibbonCallbacks` 风格一致）。

## 5. COM / 互操作

- 业务只面对 **`IApplicationContext`、`IShapeContext` 等抽象**，不经 `Globals` 之类旧入口。
- **释放**：`Shape`、`Selection`、`ShapeRange` 等 COM 对象用 `using` 或明确释放；批量操作：**收集 → 处理 → 释放**。
- **双宿主**：PowerPoint 与 WPS 行为差异只在 Adapter 收口；版本相关能力在 Adapter 或调用处判断并降级/提示。

## 6. Ribbon 与撤销

- UI 定义：`Resources/PPARibbon.xml`；回调：`RibbonCallbacks.cs`，经 `PPAUniversalComAddIn` 转发。
- 会改文档状态的操作：在可撤销路径上用 **`CreateUndoScope`**（或项目内等价的 `IUndoService` 用法）。

## 7. 配置与文档

- **`PPAConfig.xml`**（默认 `%LOCALAPPDATA%\PPA.Universal\PPAConfig.xml`，`PPAConfig.LoadOrCreate`）：顶层以当前代码为准，含 **`Table`、`Text`、`Chart`**（`Chart` 内含 `RegularFont` / `TitleFont` / `LegendFont`）、**`GlassCard`、`Logging`**。默认 XML 模板见 `PPAConfig.GetDefaultXmlContent()`。
- 新增或变更 **`PPAConfig`** 字段：在代码/XML 样例或 `docs/` 中说明含义与默认值。
- 引入新服务、改分层或 Ribbon 契约：**同步更新** `docs/`（见 **`docs/README.md`** 索引）。

## 8. 本地验证

- 常规：`dotnet build src/PPA.Layered.sln -c Release`（或仅构建 `PPA.Universal.ComAddIn` 工程）。
- 改 COM 加载项并注册：优先用 `build\rebuild-register.bat`；若出现 DLL 复制失败，先关闭 PowerPoint/WPS 再构建。

## 9. 日期与时间

- 与 **`date.mdc`**（本目录，始终应用）一致：文档或模板中的日期/时间用占位符（如 **`[当前日期]`**），**不编造**具体日历时间。
