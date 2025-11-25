# PPA 项目清理与集成指南

本文档说明新架构与现有 PPA 项目的关系。

## ✅ 集成状态：已完成

命名空间冲突已解决：

- `PPA.Core.Logging` → `PPA.Logging`（新架构）
- `PPA.Core.Abstraction.Infrastructure.ILogger`（旧项目，保持不变）

### 冲突原因

- `PPA.Core.Logging.ILogger` 与 `PPA.Core.Abstraction.Infrastructure.ILogger` 冲突
- `PPA.Core.Logging.LogLevel` 与 `PPA.Core.Abstraction.Infrastructure.LogLevel` 冲突

### 解决方案（已完成）

1. **重命名新架构命名空间**

   - 将 `PPA.Core.Logging` 改为 `PPA.Logging` 或其他不冲突的命名空间

2. **或者迁移现有代码**
   - 将现有 `PPA.Core.Abstraction.Infrastructure.ILogger` 改为使用新架构的接口

## 🚀 当前构建方式

### 新架构（dotnet CLI）

```powershell
dotnet build src\PPA.Layered.sln -c Debug
```

### VSTO 入口项目（Visual Studio）

```
1. 打开 Visual Studio
2. 打开 src\Entry\PPA.PowerPoint.VSTO\PPA.PowerPoint.VSTO.csproj
3. 重新生成解决方案
```

### 现有 PPA 项目（Visual Studio）

```
1. 打开 Visual Studio
2. 打开 PPA\PPA.sln
3. 右键解决方案 → 重新生成解决方案
```

## 项目结构

```
PPA-Universal/
├── PPA/                        # 现有 VSTO 项目（可继续使用）
│
└── src/                        # 新架构项目
    ├── Core/
    │   ├── PPA.Core/           # 核心抽象
    │   └── PPA.Business/       # 业务逻辑
    ├── Adapters/
    │   ├── PPA.Adapter.PowerPoint/
    │   └── PPA.Adapter.WPS/
    ├── Hosts/
    │   ├── PPA.PowerPoint/     # PowerPoint 类库
    │   ├── PPA.WPS/            # WPS 类库
    │   └── PPA.Universal/      # 通用版
    ├── Legacy/
    │   └── PPA.Legacy/         # 服务桥接层
    ├── Entry/
    │   └── PPA.PowerPoint.VSTO/ # 新 VSTO 入口
    └── Tests/
        └── PPA.Tests/          # 单元测试
```

## ⚠️ 重要说明

两个项目目前是**独立的**：

- PPA 项目：现有功能，仅支持 PowerPoint
- 新架构：支持 PowerPoint + WPS，用于新功能开发

## 可删除的文件

### 1. Core/Abstraction/Presentation/ (已被 PPA.Core.Abstraction 替代)

以下文件的功能已迁移到新架构，但**暂时保留**以支持现有代码：

- `ApplicationType.cs` - 被 `PPA.Core.Abstraction.PlatformType` 替代
- `FeatureSupportLevel.cs` - 被 `PPA.Core.Abstraction.Feature` 替代

**保留**（新架构未实现 Ribbon 相关）：

- `IRibbonXmlProvider.cs`
- `IRibbonIconProvider.cs`
- `IRibbonCommandRouter.cs`

### 2. Core/Abstraction/Infrastructure/ (部分替代)

**保留**（现有代码仍在使用）：

- `IApplicationProvider.cs` - 被 `IApplicationContext` 替代，但现有 UI 代码依赖
- `ILogger.cs` - 被 `PPA.Core.Logging.ILogger` 替代，但命名空间不同
- `LogLevel.cs` - 被 `PPA.Core.Logging.LogLevel` 替代

### 3. Core/Abstraction/Business/ (保留)

这些接口仍被 `Manipulation/` 下的实现类使用：

- `ITableFormatHelper.cs` 等 - 保留

## 建议的清理步骤

### 阶段一：验证集成（当前）

```
✅ 已完成：
- ThisAddIn.cs 引用新架构
- 项目引用新项目
```

### 阶段二：逐步迁移（后续）

1. **迁移 Business 服务**

   - 将 `Manipulation/` 下的 Helper 类改为使用 `PPA.Business` 服务
   - 修改 DI 注册

2. **迁移 UI 代码**

   - 将 `CustomRibbon` 改为使用新架构的 `IApplicationContext`

3. **清理旧代码**
   - 删除不再使用的接口和类

## 不可删除的文件

以下文件仍然是必要的：

### Core/

- `ApplicationProvider.cs` - UI 代码依赖
- `Profiler.cs` - 日志系统
- `ExHandler.cs` - 异常处理
- `ResourceManager.cs` - 多语言资源
- `Logging/` - 日志适配器

### Manipulation/

- 所有文件 - 业务逻辑实现

### UI/

- 所有文件 - 用户界面

### Utilities/

- 所有文件 - 工具类

### Shape/

- 所有文件 - 形状操作

## 当前状态

```
PPA 项目现在同时引用：
├── 旧架构（Core/Abstraction/）- 现有功能
└── 新架构（src/）           - 新功能开发

这允许渐进式迁移而不影响现有功能。
```
