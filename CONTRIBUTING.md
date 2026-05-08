# 贡献指南

感谢参与 **PPA-Universal**。提交前请阅读 **[docs/README.md](docs/README.md)** 与仓库 **`.cursor/rules/AGENTS.md`**（分层、异常、COM 与日志约定）。

## 流程

1. 在 Issue 中说明问题或方案（功能请先对齐 [功能扩展研究报告](docs/功能扩展研究报告.md) 中的范围，避免与路线图严重冲突）。
2. Fork 后建分支，小步提交；PR 描述写清「动机、改动点、如何验证」。
3. 至少本地 `dotnet build src/PPA.Layered.sln -c Release` 与 `dotnet test src/PPA.Layered.sln -c Release` 通过；若改 Ribbon，请按 [docs/Ribbon-manual-regression.md](docs/Ribbon-manual-regression.md) 在 PowerPoint 与 WPS 各测一遍。

## 代码约定（摘要）

- 业务逻辑放在 **PPA.Business**，通过接口 + DI；平台相关只在 **Adapter**。
- 关键路径用项目统一的异常/日志模式（见 `AGENTS.md`）；避免无捕获的 COM 泄漏。
- **代码格式**：以仓库根 `.editorconfig` 为准（例如 C# 使用 Tab）；主干已与 `dotnet format src/PPA.Layered.sln` 全库对齐。提交前建议对本分支执行一次相同命令，减少仅因缩进/换行产生的 diff。
- 尽量保持 PR 单一职责，避免无关重命名与「顺手」大范围逻辑改动。

## 报告 Bug

请附：版本（Office/WPS、位数）、复现步骤、预期与实际行为、相关日志片段（若有）。
