# PPA 分层架构

本目录包含 PPA 项目的新分层架构实现。

## 架构概览

```
src/
├── Core/                           # 共享核心层
│   ├── PPA.Core/                   # 基础设施（抽象接口、DI、日志）
│   └── PPA.Business/               # 业务逻辑（平台无关）
│
├── Adapters/                       # 平台适配器层
│   ├── PPA.Adapter.PowerPoint/     # PowerPoint 适配器（NetOffice）
│   └── PPA.Adapter.WPS/            # WPS 适配器（COM dynamic）
│
└── Hosts/                          # 版本入口层
    ├── PPA.Universal/              # 通用版入口（自动检测平台）
    └── PPA.Universal.ComAddIn/     # COM 加载项入口（供 PowerPoint/WPS 调用）
```

## 精简说明

- 仅保留 `Core`、`Adapters` 以及 `Hosts/PPA.Universal`，确保双平台共用骨干
- 原有 `Entry/`、`Hosts/PPA.PowerPoint`、`Hosts/PPA.WPS`、`Legacy/`、`Tests/` 等目录已移除
- 若需恢复旧版入口或调试工程，可在历史提交中查找对应目录

## 项目说明

### PPA.Core

- **职责**：平台无关的基础设施
- **内容**：
  - `Abstraction/` - 平台无关接口（IApplicationContext、IShapeContext 等）
  - `Logging/` - 日志抽象和基础实现
  - `DI/` - 依赖注入扩展
  - `Exceptions/` - 自定义异常

### PPA.Business

- **职责**：平台无关的业务逻辑
- **内容**：
  - `Abstractions/` - 业务服务接口
  - `Services/` - 业务服务实现
  - `DI/` - 业务服务注册

### PPA.Adapter.PowerPoint

- **职责**：PowerPoint 平台的具体实现
- **内容**：
  - 实现 PPA.Core 中定义的所有接口
  - 使用 NetOffice 库与 PowerPoint 交互
  - `DI/` - PowerPoint 适配器服务注册

### PPA.Adapter.WPS

- **职责**：WPS 平台的具体实现
- **内容**：
  - 使用 COM dynamic 方式与 WPS 交互
  - `WPSHelper.cs` - WPS COM 互操作辅助
  - `WPSContext.cs` 等上下文实现
  - `DI/` - WPS 适配器服务注册

### PPA.Universal

- **职责**：唯一保留的入口层，负责自动检测平台并加载适配器
- **内容**：
  - `Platform/PlatformDetector.cs` - 运行时平台检测
  - `Platform/AdapterFactory.cs` - 适配器动态加载工厂
  - `UniversalBootstrapper.cs` - 通用版 DI 容器初始化
  - `PPAUniversal.cs` - 通用版主入口类
  - `Integration/` - 集成帮助类
- **说明**：原有 PowerPoint/WPS 专版入口已移除，统一通过通用版加载

### PPA.Universal.ComAddIn

- **职责**：向 PowerPoint/WPS 提供一个最小化 COM Add-in 宿主
- **内容**：
  - `PPAUniversalComAddIn.cs` - 实现 `IDTExtensibility2`，在 `OnConnection` 中调用 `UniversalIntegration.Initialize`
  - `Properties/AssemblyInfo.cs` - 定义 COM 可见性与 GUID
- **用途**：通过 `regasm` 注册后，可在 PowerPoint/WPS 的“COM 加载项”面板勾选启用，减少手动集成

## 构建

```powershell
# 使用构建脚本
.\build\build-layered.ps1 -Configuration Release -Clean -Restore

# 或直接使用 dotnet CLI
dotnet build src\PPA.Layered.sln -c Release
```

## COM 加载项注册

```powershell
# 1. 先完成构建
dotnet build src\PPA.Layered.sln -c Debug

# 2. 注册 COM 加载项（默认选择 64 位 regasm，若 Office 为 32 位可传 -RegasmPath）
pwsh .\tools\register-com-addin.ps1 -Action Register -Configuration Debug

# 3. 如需卸载
pwsh .\tools\register-com-addin.ps1 -Action Unregister -Configuration Debug
```

> 注册成功后，打开 PowerPoint/WPS → 选项 → 加载项 → 管理 “COM 加载项” → 转到，即可看到 `PPA.Universal.ComAddIn` 条目，勾选启用即可。

## 完成状态

1. ~~**阶段一**：架构重构~~ ✅ 已完成
2. ~~**阶段二**：实现 WPS 适配器~~ ✅ 已完成
3. ~~**阶段三**：实现通用版~~ ✅ 已完成
4. ~~**阶段四**：测试与发布~~ ✅ 已完成

## 测试与发布

```powershell
# 运行测试
.\build\test.ps1

# 发布所有版本
.\build\publish.ps1 -Version All -RunTests

# 发布单个版本
.\build\publish.ps1 -Version PowerPoint
.\build\publish.ps1 -Version WPS
.\build\publish.ps1 -Version Universal
```

## 使用示例

### 通用版（自动检测平台）

```csharp
using PPA.Universal;
using PPA.Universal.Integration;

// 方式1：使用静态集成类
UniversalIntegration.InitializeAuto();  // 自动检测运行中的应用
var context = UniversalIntegration.Context;
var platform = UniversalIntegration.Platform;  // PowerPoint 或 WPS

// 方式2：使用实例
var ppa = new PPAUniversal();
ppa.StartupAuto();
// 或指定应用程序对象
ppa.Startup(applicationObject);
```

### 示例：基于幻灯片左对齐选中形状

```csharp
using PPA.Universal.Integration;

UniversalIntegration.AlignSelectionLeftToSlide();
```

## WPS 三线表兼容说明

三线表格式（表头上/下边框 + 末行下边框）在 **PowerPoint** 与 **WPS** 上的行为存在差异，当前实现做了尽量统一，但仍有已知限制：

- **实现方式**

  - 业务层统一通过 `TableFormatService` 对行应用 `RowStyle`：清空所有边框 → 设置字体/对齐 → 仅绘制三条横线。
  - 单元格背景在核心层只调用 `ICellContext.SetBackgroundVisible(false)`，不再设置具体颜色。
  - 在 WPS 适配层中，通过操作 `Cell.Shape.Fill`（可见性/透明度）尽量“盖掉”原有底纹。

- **三类表在 WPS 下的表现目标**

  - **表 1：已有填充背景** → 尽量清理到“近似透明”，只保留三线和文字。
  - **表 2：无填充** → 保持透明底 + 三线（推荐场景）。
  - **表 3：套用表格样式** → 尽量弱化样式底色，但受限于 WPS 内部样式引擎，某些主题/区域仍可能残留轻微底纹。

- **已知限制（WPS 平台）**

  - WPS 未公开与 PowerPoint `No Style, No Grid` 等价的清样式 API，当前实现只能在 `Shape.Fill` 层“覆盖”，无法彻底关闭表格样式引擎。
  - 对部分主题/样式，标题行或合并单元格区域的底色可能与纯透明场景略有差异（但不会影响三线结构和数据可读性）。

- **使用建议**
  - 在 WPS 中使用三线表前，**优先选择简洁/无底纹的表格样式**，或直接从“无填充”表格开始再点击“三线表”。
  - 若对配色有严格要求，推荐在 PowerPoint 中完成最终排版，再在 WPS 中仅做轻量查看/演示。
  - 如遇到极端主题导致底纹无法接受，可在幻灯片模板层调整页面背景颜色来弱化差异.
