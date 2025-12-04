# PPA-Universal (PPA 分层架构)

PPA-Universal 是 PPA (PowerPoint Assistant) 项目的新一代分层架构实现。它通过统一的抽象层和适配器模式，实现了对 **Microsoft PowerPoint** 和 **WPS 演示** 的双平台支持，允许开发者编写一次代码，即可在两个平台上运行。

## 📖 文档中心

本项目的主要文档位于源码目录中，请参考：

- 👉 **[核心架构与开发文档](src/README.md)**  
  包含详细的架构设计（Core/Adapters/Hosts）、模块职责说明、详细的构建与测试指南。

- 📄 **[项目分析评估报告](docs/project_analysis_report.md)**  
  项目的技术栈分析与架构评估。

- 🤝 **[贡献指南](CONTRIBUTING.md)**  
  参与贡献代码的规则与建议。

## 🚀 快速上手

### 构建项目

项目提供了 PowerShell 脚本以简化构建流程：

```powershell
# 执行完整构建（清理、还原、编译）
.\build\build-layered.ps1 -Configuration Release -Clean -Restore

# 或者使用 .NET CLI
dotnet build src\PPA.Layered.sln -c Release
```

### 安装/注册插件

若要将 PPA 作为 COM 加载项安装到 Office/WPS 中：

```powershell
# 1. 构建 Debug 版本
dotnet build src\PPA.Layered.sln -c Debug

# 2. 注册 COM 组件
pwsh .\tools\register-com-addin.ps1 -Action Register -Configuration Debug
```

> 注册成功后，在 PowerPoint 或 WPS 的“COM 加载项”设置中即可看到并启用本插件。

## 🏗️ 目录结构简述

- **`src/`**: 核心源代码
  - **`Core/`**: 抽象接口与业务逻辑（平台无关）
  - **`Adapters/`**: 针对 PowerPoint 和 WPS 的具体实现
  - **`Hosts/`**: 应用程序入口与启动器
- **`build/`**: 自动化构建脚本
- **`tools/`**: 辅助工具脚本（如 COM 注册工具）
- **`docs/`**: 项目文档存放处

---

_更多详细信息，请务必阅读 [src/README.md](src/README.md)。_
