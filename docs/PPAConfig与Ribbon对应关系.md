# PPAConfig.xml 与 Ribbon 操作对应关系

配置文件默认路径：`%LOCALAPPDATA%\PPA.Universal\PPAConfig.xml`（由 `UniversalBootstrapper` 在初始化时 `LoadOrCreate`）。默认 XML 模板见 `PPA.Core.Configuration.PPAConfig.GetDefaultXmlContent()`。

更新日期：[当前日期]

## 1. 配置节总览

| 配置节 | 作用 |
| ------ | ---- |
| `Defaults` | 幻灯片宽高兜底（宿主读不到 `PageSetup` 或与演示文稿断开时使用）；影响对齐「相对幻灯片」、形状创建默认矩形居中、适配器 `IPresentationContext.SlideWidth/Height` 回退值。 |
| `Table` | 三线表样式、表格边框/数字格式、数据行与表头字体等。 |
| `Text` | 文本框批量格式化（`Text.Font`）及段落/项目符号等（供后续扩展；当前 Ribbon「文本字体」仅使用 `Text.Font`）。 |
| `Chart` | 图表内常规/标题/图例字体（`ChartBatchService.FormatChartFont`）。 |
| `GlassCard` | 毛玻璃卡片几何、渐变与卡片内文字样式。 |
| `Duplicate` | 矩阵复制 / 线性复制对话框初始值（行数、列数、间距、份数、方向）。 |
| `Logging` | 文件日志开关、保留策略、级别、单文件大小等。 |

## 2. Ribbon 按钮 → 配置映射

下列仅列出**与配置文件直接相关**的项；纯几何对齐、分布、吸附等不读写配置。

| Ribbon（PPARibbon.xml） | 回调 / 服务 | 使用的配置 |
| ------------------------- | ------------ | ---------- |
| **矩阵复制** | `DuplicateCopyDialogs` + `IShapeDuplicateService` | `Duplicate` |
| **线性复制** | 同上 | `Duplicate` |
| **三线表** | `ITableFormatService.FormatTableAsThreeLine` | `Table` |
| **全稿三线表** | `ITableBatchService.FormatAllTables` → `ITableFormatService` | `Table` |
| **表格字体** | `ITableFormatService.FormatTableFont`（未传入字体时） | **第 1 行**：`Table.HeaderRowFont`（缺省则同数据行）；**第 2 行起**：`Table.DataRowFont`（缺省则同表头），与三线表表头加粗一致。 |
| **文本字体** | `ITextBatchService.FormatTextBoxFont` | **`Text.Font`**（经 `FontConfig.ToFontStyle()`） |
| **图表字体** | `IChartBatchService.FormatChartFont` | `Chart.RegularFont` / `TitleFont` / `LegendFont` |
| **查找替换** | `ITextBatchService.ReplaceText` | 无专用节（行为由业务实现决定） |
| **毛玻璃卡片** | `IGlassCardService.CreateGlassCard` | `GlassCard`（含 `TextStyle`） |
| **创建矩形**（仅幻灯片选中时默认位置） | `IShapeCreationService` | `Defaults` + 当前演示文稿尺寸 |
| **对齐参考：幻灯片** | `IAlignmentService` 计算参考边 | 当前演示文稿尺寸；缺失时用 **`Defaults`** |

其余按钮（等宽/等高、延伸、裁切、隐藏等）当前不读取上述 XML 节。

## 3. 近期代码约定（避免写死）

- **文本框字体**：须来自 `PPAConfig.Text.Font`，不得在 Ribbon 层写死字号/字体名（已通过 `FontConfig.ToFontStyle()` 统一）。
- **幻灯片尺寸兜底**：统一使用 `Defaults.SlideWidthFallback` / `SlideHeightFallback`；业务层与 `PowerPointPresentationContext` / `WPSPresentationContext` 在 COM 读数失败时与此保持一致。
- **字体节点主题色**：`ParseFontElement` 同时支持 `ThemeColorIndex` 与 `ThemeColor`（如 `Accent2`），与默认模板中 `Text.Font` 的写法一致。

## 4. `Logging` 与 Ribbon

`Logging` **无对应 Ribbon 按钮**；仅影响启动时是否注册文件日志及滚动策略（见 `UniversalBootstrapper`）。

## 5. 默认值与兜底（避免与模板 XML 漂移）

- **磁盘模板**：`PPAConfig.GetDefaultXmlContent()` 为首次创建或解析失败重写时的权威内容。
- **代码兜底**：`PpaConfigTemplateFallbacks`（与 `GetDefaultXmlContent()` 对齐）集中提供幻灯片宽高、Ribbon「文本字体」在 `Text` 缺失时、图表标题/图例在 `Chart` 缺失时的 `FontStyle`；`AdapterFactory`、对齐与形状创建中的 960×540 亦引用同一常量。
- **极端失败**：`LoadOrCreate` 在重写文件仍无法解析时，会再尝试 **内存解析** 默认 XML（`LoadFromDefaultXmlString`），最后才退回空的 `PPAConfig()`。
- **仍存在的合理分支**：`FontConfig.ToFontStyle()` 在「空节点」时用 **15pt** 作为字号兜底，与模板中 **表格数据行** 字号一致；**文本框** 专用兜底请用 `PpaConfigTemplateFallbacks.TextBoxRibbonFontStyle()`（16pt、加粗、Accent2）。

## 6. 相关源码索引

- 配置模型与加载：`src/Core/PPA.Core/Configuration/PPAConfig.cs`（含 `PpaConfigTemplateFallbacks`）
- Ribbon 回调：`src/Hosts/PPA.Universal.ComAddIn/RibbonCallbacks.cs`
- 宿主加载配置与适配器：`src/Hosts/PPA.Universal/UniversalBootstrapper.cs`、`Platform/AdapterFactory.cs`
