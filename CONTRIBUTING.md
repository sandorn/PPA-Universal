# 贡献指南

感谢参与 **PPA-Universal**。提交前请阅读 **[docs/README.md](docs/README.md)** 与仓库 **`.cursor/rules/AGENTS.md`**（分层、异常、COM 与日志约定）。

## 流程

1. 在 Issue 中说明问题或方案（功能请先对齐 [功能扩展研究报告](docs/功能扩展研究报告.md) 中的范围，避免与路线图严重冲突）。
2. Fork 后建分支，小步提交；PR 描述写清「动机、改动点、如何验证」。
3. 至少本地 `dotnet build src/PPA.Layered.sln -c Release` 通过；若改 Ribbon，请在 PowerPoint 与 WPS 各测一遍图标与主要按钮。

## 代码约定（摘要）

- 业务逻辑放在 **PPA.Business**，通过接口 + DI；平台相关只在 **Adapter**。
- 关键路径用项目统一的异常/日志模式（见 `AGENTS.md`）；避免无捕获的 COM 泄漏。
- 尽量保持 PR 单一职责，避免大范围格式化与无关重命名。

## 报告 Bug

请附：版本（Office/WPS、位数）、复现步骤、预期与实际行为、相关日志片段（若有）。
