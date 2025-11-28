# PowerPoint 三线表竖线残留问题报告

- **报告日期**: [当前日期]
- **负责人**: [项目名称] 团队
- **问题编号**: PPT-THREELINE-01

## 问题概述

文件：D:\CODES\PPA-Universal\src\Core\PPA.Business\Services\TableFormatService.cs
行号：168
在 PowerPoint 平台使用“选中表格 → 三线表”功能时，表格的竖向边框无法被彻底清除，即使业务层已经调用 `SetBorder(BorderEdge.All, BorderStyle.None)` 重置边框。WPS 平台表现正常，说明该问题为 PowerPoint 特有的渲染或样式残留。

## 影响范围

- 平台：PowerPoint（WPS 正常）
- 功能：Ribbon「三线表」按钮
- 现象：表头+数据区仍能设置上/下粗线，但内列竖线保持旧样式，无法满足三线表的视觉要求。

## 复现步骤

1. 在 PowerPoint 中插入带竖线的表格（默认表格或内置样式）。
2. 选中表格或多个表格。
3. 点击 PPA Ribbon → 表格组 → 「三线表」。
4. 观察结果：横向三线生效，但竖向边界未消失。

## 已有尝试

- 在 `TableFormatService.SetRowStyle` 中调用 `cell.SetBorder(BorderEdge.All, BorderStyle.None)` 清除旧样式。
- 在首行/末行分别调用 `SetHeaderBottomBorder`、`SetLastRowBottomBorder` 强化横向边界。
- 日志显示流程执行正常，PowerPoint 仍保留竖线，说明问题可能出在 PowerPoint 对表格样式或主题的缓存机制。

## 后续计划

1. 研究 PowerPoint COM 接口是否需要调用 `Table.ApplyStyle` 或 `Shape.Line.Visible` 等额外操作才能去除默认竖线。
2. 检查是否存在列级边框集合（如 `Table.Columns[i].Borders`）需要单独清理。
3. 如确认为 PowerPoint 限制，考虑提供“移除竖线”备选项或使用 VBA 方式批量删除。

> 当前功能速度、成功率均满足要求；本报告用于追踪 PowerPoint 竖线残留的后续处理进度。
