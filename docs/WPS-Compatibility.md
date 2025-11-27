# PPA WPS 兼容性文档

> 版本：1.0  
> 更新日期：2024-11

## 概述

PPA (PowerPoint Automation) 插件支持 Microsoft PowerPoint 和 WPS 演示两个平台。本文档详细说明 WPS 平台的兼容性状况和已知限制。

## 兼容性总览

| 功能模块   | 兼容性      | 性能 | 备注                     |
| ---------- | ----------- | ---- | ------------------------ |
| 文本格式化 | ✅ 完全支持 | 正常 | 所有属性可用             |
| 图表格式化 | ✅ 完全支持 | 正常 | 需使用实际字体名         |
| 形状对齐   | ✅ 完全支持 | 正常 | 位置/大小/参考线全部可用 |
| 表格格式化 | ⚠️ 部分支持 | 较慢 | 背景清除不支持           |

## 详细兼容性

### 1. 文本格式化 (TextFormatHelper)

**状态：完全兼容**

| 属性                                                  | 支持情况 |
| ----------------------------------------------------- | -------- |
| TextFrame.MarginTop/Left/Right/Bottom                 | ✅       |
| TextFrame.TextRange.Font.Name                         | ✅       |
| TextFrame.TextRange.Font.NameFarEast                  | ✅       |
| TextFrame.TextRange.Font.Size                         | ✅       |
| TextFrame.TextRange.Font.Bold                         | ✅       |
| TextFrame.TextRange.Font.Color                        | ✅       |
| TextFrame.TextRange.ParagraphFormat.Alignment         | ✅       |
| TextFrame.TextRange.ParagraphFormat.SpaceBefore/After | ✅       |

**注意事项：**

- 线条 (Type=9) 没有文本，访问 TextRange 会抛出异常
- 图表 (Type=3) 的文本需通过 Chart API 访问
- 表格 (HasTable=True) 的文本需通过 Table.Cell 访问

### 2. 图表格式化 (ChartFormatHelper)

**状态：完全兼容**

| 属性                       | 支持情况 |
| -------------------------- | -------- |
| Chart.HasTitle             | ✅       |
| Chart.ChartTitle.Font.Name | ✅       |
| Chart.ChartTitle.Font.Size | ✅       |
| Chart.HasLegend            | ✅       |
| Chart.Legend.Font.Name     | ✅       |
| Chart.Legend.Font.Size     | ✅       |
| Chart.SeriesCollection     | ✅       |
| Chart.Axes                 | ✅       |

**重要限制：**

```csharp
// ❌ 错误：ChartFont 不支持主题字体占位符
chart.ChartTitle.Font.Name = "+mn-lt";

// ✅ 正确：使用实际字体名称
chart.ChartTitle.Font.Name = "微软雅黑";
```

### 3. 形状对齐 (AlignHelper)

**状态：完全兼容**

| 属性/操作           | 支持情况 |
| ------------------- | -------- |
| Shape.Left/Top      | ✅       |
| Shape.Width/Height  | ✅       |
| Presentation.Guides | ✅       |
| Shape.Fill          | ✅       |
| Shape.Line          | ✅       |

所有对齐、吸附、拉伸、等大小操作均正常工作。

### 4. 表格格式化 (TableFormatService)

**状态：部分兼容**

| 属性/操作                              | 支持情况 | 说明           |
| -------------------------------------- | -------- | -------------- |
| Cell.Shape.TextFrame.TextRange.Font.\* | ✅       | 字体设置正常   |
| Cell.Borders[*].Visible                | ✅       | 边框可见性     |
| Cell.Borders[*].Weight                 | ✅       | 边框粗细       |
| Cell.Borders[*].ForeColor              | ✅       | 边框颜色       |
| Table.FirstRow/FirstCol 等             | ✅       | 表格设置       |
| **Cell.Shape.Fill**                    | ❌       | 背景色无法修改 |
| **ExecuteMso("TableStyleClearTable")** | ❌       | 清除样式无效   |
| **Application.ScreenUpdating**         | ❌       | 不存在此属性   |

**性能说明：**

由于 WPS 不支持以下优化机制，表格格式化性能约为 PowerPoint 的 1/10：

1. 无法暂停屏幕重绘 (ScreenUpdating)
2. 无法使用 WM_SETREDRAW 消息
3. 无法使用 LockWindowUpdate API
4. 每个单元格操作都会触发界面刷新

**建议：** 对于大型表格或需要高性能的场景，建议使用 Microsoft PowerPoint。

## WPS COM API 限制汇总

### 不支持的属性

| 属性                       | 说明                   |
| -------------------------- | ---------------------- |
| Application.ScreenUpdating | WPS 不存在此属性       |
| Cell.Shape.Fill.\*         | 表格单元格背景无法修改 |

### 不支持的方法

| 方法                                                       | 说明           |
| ---------------------------------------------------------- | -------------- |
| Application.CommandBars.ExecuteMso("TableStyleClearTable") | 无效果         |
| Table.ApplyStyle()                                         | 可能不完全支持 |

### 不支持的 Windows API

| API              | 说明             |
| ---------------- | ---------------- |
| WM_SETREDRAW     | WPS 不响应此消息 |
| LockWindowUpdate | 对 WPS 窗口无效  |

## 代码适配策略

### 1. 能力探测模式

对于可能不支持的操作，使用"首次失败后跳过"策略：

```csharp
private bool _canSetFillColor = true;

private void SetCellBackground(Cell cell, int color)
{
    if (!_canSetFillColor) return;

    try
    {
        cell.Shape.Fill.ForeColor.RGB = color;
    }
    catch
    {
        _canSetFillColor = false; // 首次失败后跳过所有后续尝试
    }
}
```

### 2. 字体设置适配

```csharp
// 文本框/表格单元格：可使用主题字体
textRange.Font.Name = "+mn-lt";      // ✅ OK
textRange.Font.NameFarEast = "+mn-ea"; // ✅ OK

// 图表：必须使用实际字体名
chartTitle.Font.Name = "微软雅黑";    // ✅ OK
chartTitle.Font.Name = "+mn-lt";     // ❌ 显示为 "+mn-lt"
```

### 3. 形状类型检测

```csharp
int shapeType = shape.Type;
bool hasTable = false;
bool hasChart = false;

try { hasTable = shape.HasTable == MsoTriState.msoTrue; } catch { }
try { hasChart = shape.HasChart == MsoTriState.msoTrue; } catch { }

// 根据形状类型选择处理方式
if (hasTable) { /* 通过 Table.Cell 访问 */ }
else if (hasChart) { /* 通过 Chart API 访问 */ }
else if (shapeType == 9) { /* 线条，跳过文本操作 */ }
else { /* 普通文本框 */ }
```

## 测试工具

项目包含 WPS 兼容性测试工具：`src/Tests/PPA.WPS.Debug`

运行方式：

```bash
cd src/Tests/PPA.WPS.Debug
dotnet run
```

测试菜单：

1. 表格格式化测试
2. 文本格式化测试
3. 图表格式化测试
4. 形状对齐测试

## 版本历史

| 版本 | 日期    | 变更                         |
| ---- | ------- | ---------------------------- |
| 1.0  | 2024-11 | 初始版本，完成基础兼容性测试 |

## 参考资料

- WPS 开发文档：https://open.wps.cn/
- NetOffice 文档：https://netoffice.io/
- PowerPoint COM 参考：https://docs.microsoft.com/office/vba/api/overview/powerpoint
