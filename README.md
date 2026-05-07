# PPA-Universal

面向 PowerPoint / WPS 的 COM 加载项项目。

## 给使用者（快速开始）

1. 在 Windows 上安装 .NET SDK（可执行 `dotnet`）。
2. 双击执行：`build\rebuild-register.bat`
3. 打开 PowerPoint 或 WPS，在 COM 加载项中勾选 `PPA.Universal.ComAddIn`。

> 脚本会自动提权、清理并构建 Release、执行 x64/x86 RegAsm 注册。

## 常见问题

- 构建时 DLL 被占用：关闭 PowerPoint/WPS 后重试（脚本已尝试自动结束进程）。
- 注册后看不到加载项：重启 Office/WPS；确认本机启用了 COM 加载项。

## 开发者入口

开发者与贡献者文档请看：
- [docs/README.md](docs/README.md)

该文档包含架构、目录职责、开发流程、贡献约定与专题文档索引。
