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
    ├── PPA.PowerPoint/             # PowerPoint 专版入口
    ├── PPA.WPS/                    # WPS 专版入口
    └── PPA.Universal/              # 通用版入口（自动检测平台）
```

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

### PPA.PowerPoint

- **职责**：PowerPoint 版本的入口点
- **内容**：
  - `AddInBootstrapper.cs` - DI 容器初始化
  - `Integration/` - 与现有代码的集成帮助

### PPA.WPS

- **职责**：WPS 版本的入口点
- **内容**：
  - `WPSAddInBootstrapper.cs` - DI 容器初始化
  - `WPSAddIn.cs` - WPS 插件主类
  - `FeatureCompatibility.cs` - 功能兼容性检查

### PPA.Universal

- **职责**：通用版入口，自动检测平台
- **内容**：
  - `Platform/PlatformDetector.cs` - 运行时平台检测
  - `Platform/AdapterFactory.cs` - 适配器动态加载工厂
  - `UniversalBootstrapper.cs` - 通用版 DI 容器初始化
  - `PPAUniversal.cs` - 通用版主入口类
  - `Integration/` - 集成帮助类

## 构建

```powershell
# 使用构建脚本
.\build\build-layered.ps1 -Configuration Release -Clean -Restore

# 或直接使用 dotnet CLI
dotnet build src\PPA.Layered.sln -c Release
```

## 与现有代码集成

在现有的 `ThisAddIn.cs` 中添加：

```csharp
using PPA.PowerPoint.Integration;

// 在 Startup 方法中
private void ThisAddIn_Startup(object sender, EventArgs e)
{
    // ... 现有初始化代码 ...

    // 初始化新架构
    LegacyIntegration.Initialize(NetApp, NativeApp);
}

// 在 Shutdown 方法中
private void ThisAddIn_Shutdown(object sender, EventArgs e)
{
    // 清理新架构资源
    LegacyIntegration.Cleanup();

    // ... 现有清理代码 ...
}
```

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

### PowerPoint 专版

```csharp
using PPA.PowerPoint.Integration;

LegacyIntegration.Initialize(netApp, nativeApp);
var context = LegacyIntegration.Context;
```

### WPS 专版

```csharp
using PPA.WPS;

var addIn = new WPSAddIn();
addIn.StartupAuto();  // 自动连接 WPS
var context = addIn.Context;
```
