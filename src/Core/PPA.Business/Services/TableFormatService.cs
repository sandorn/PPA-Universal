using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
    /// <summary>
    /// 表格格式化服务实现
    /// </summary>
    public class TableFormatService : ITableFormatService
    {
        private readonly ILogger _logger;
        private readonly ITableOperations _tableOps;

        public TableFormatService(ILogger logger, ITableOperations tableOps)
        {
            _logger = logger ?? NullLogger.Instance;
            _tableOps = tableOps;
        }

        public void FormatTable(ITableContext table, TableFormatOptions options = null)
        {
            if (table == null)
            {
                _logger.LogWarning("表格对象为空，跳过格式化");
                return;
            }

            options = options ?? new TableFormatOptions();
            _logger.LogInformation($"开始格式化表格，行数: {table.RowCount}, 列数: {table.ColumnCount}");

            try
            {
                // 应用表格样式
                if (options.ApplyTableStyle && !string.IsNullOrEmpty(options.TableStyleId))
                {
                    table.ApplyStyle(options.TableStyleId);
                }

                // 设置表格选项
                if (options.Settings != null)
                {
                    table.SetTableOptions(
                        options.Settings.FirstRow,
                        options.Settings.FirstCol,
                        options.Settings.LastRow,
                        options.Settings.LastCol,
                        options.Settings.HorizBanding,
                        options.Settings.VertBanding
                    );
                }

                // 格式化表头
                if (options.FormatHeader && options.HeaderStyle != null && table.RowCount > 0)
                {
                    SetRowStyle(table, 1, options.HeaderStyle);
                }

                // 格式化数据行
                if (options.FormatDataRows && options.DataRowStyle != null)
                {
                    for (int row = 2; row <= table.RowCount; row++)
                    {
                        var style = (options.AlternateRowStyle != null && row % 2 == 0)
                            ? options.AlternateRowStyle
                            : options.DataRowStyle;
                        SetRowStyle(table, row, style);
                    }
                }

                // 再次强化表头下边框，避免部分平台被数据行顶部边框覆盖
                if (options.FormatHeader && options.HeaderStyle != null && table.RowCount > 0)
                {
                    SetHeaderBottomBorder(table, options.HeaderStyle);
                }

                // 末行下边框（使用表头边框样式）
                if (options.FormatHeader && options.HeaderStyle != null && table.RowCount > 1)
                {
                    SetLastRowBottomBorder(table, options.HeaderStyle);
                }

                _logger.LogInformation("表格格式化完成");
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"格式化表格时出错: {ex.Message}", ex);
                throw;
            }
        }

        public void FormatTableWithDefaults(ITableContext table)
        {
            // 使用默认配置格式化（与旧实现类似的风格）
            var options = new TableFormatOptions
            {
                FormatHeader = true,
                FormatDataRows = true,
                ApplyBorders = true,
                ApplyFont = true,
                // WPS 不支持 ApplyStyle，背景清除由调用方处理
                ApplyTableStyle = false,
                Settings = new TableSettings
                {
                    FirstRow = true,
                    FirstCol = false,
                    LastRow = false,
                    LastCol = false,
                    HorizBanding = false,
                    VertBanding = false
                },
                HeaderStyle = new RowStyle
                {
                    HideBackground = true, // 隐藏背景，使用主题色
                    FontName = "等线",
                    FontNameFarEast = "等线",
                    FontSize = 12,
                    Bold = true,
                    ThemeColorIndex = 13, // 深色1 (dk1)
                    Alignment = TextAlignment.Center,
                    TopBorder = BorderStyle.SolidTheme(5, 1.75f),    // Accent1
                    BottomBorder = BorderStyle.SolidTheme(5, 1.75f), // Accent1
                    LeftBorder = BorderStyle.None,   // 隐藏左边框
                    RightBorder = BorderStyle.None   // 隐藏右边框
                },
                DataRowStyle = new RowStyle
                {
                    HideBackground = true, // 隐藏背景
                    FontName = "等线",
                    FontNameFarEast = "等线",
                    FontSize = 11,
                    Bold = false,
                    ThemeColorIndex = 13, // 深色1 (dk1)
                    TopBorder = BorderStyle.SolidTheme(6, 1.0f),     // Accent2
                    LeftBorder = BorderStyle.None,   // 隐藏左边框
                    RightBorder = BorderStyle.None   // 隐藏右边框
                }
            };

            FormatTable(table, options);
        }

        public void FormatTableAsThreeLine(ITableContext table)
        {
            if (table == null)
            {
                _logger.LogWarning("表格对象为空，无法应用三线表格式");
                return;
            }

            var options = CreateThreeLineOptions();
            _logger.LogInformation("开始应用三线表格式");
            FormatTable(table, options);
        }

        public void SetRowStyle(ITableContext table, int rowIndex, RowStyle style)
        {
            if (table == null || style == null || rowIndex < 1 || rowIndex > table.RowCount)
                return;

            for (int col = 1; col <= table.ColumnCount; col++)
            {
                var cell = table.GetCell(rowIndex, col);
                if (cell == null) continue;

                // 清空现有边框，避免平台保留旧样式（尤其是 PowerPoint 竖线）
                cell.SetBorder(BorderEdge.All, BorderStyle.None);
                // TODO: PPT-THREELINE-01 PowerPoint 仍可能残留内置竖线，需要后续做平台特定清理

                // 设置背景
                if (style.HideBackground)
                {
                    cell.SetBackgroundVisible(false);
                }
                else if (style.BackgroundColor.HasValue)
                {
                    cell.SetBackground(style.BackgroundColor.Value);
                }

                // 设置字体
                var fontStyle = style.ToFontStyle();
                if (!string.IsNullOrEmpty(fontStyle.Name) || fontStyle.Size > 0)
                {
                    cell.SetFont(fontStyle);
                }

                // 设置对齐
                if (style.Alignment.HasValue)
                {
                    cell.SetAlignment(style.Alignment.Value);
                }

                // 设置边框
                if (style.TopBorder.HasValue)
                {
                    cell.SetBorder(BorderEdge.Top, style.TopBorder.Value);
                }
                if (style.BottomBorder.HasValue)
                {
                    cell.SetBorder(BorderEdge.Bottom, style.BottomBorder.Value);
                }
                if (style.LeftBorder.HasValue)
                {
                    cell.SetBorder(BorderEdge.Left, style.LeftBorder.Value);
                }
                if (style.RightBorder.HasValue)
                {
                    cell.SetBorder(BorderEdge.Right, style.RightBorder.Value);
                }
            }
        }

        /// <summary>
        /// 强化表头下边框
        /// </summary>
        private void SetHeaderBottomBorder(ITableContext table, RowStyle headerStyle)
        {
            if (table == null || table.RowCount == 0) return;
            var borderStyle = headerStyle.BottomBorder ?? headerStyle.TopBorder;
            if (!borderStyle.HasValue) return;

            for (int col = 1; col <= table.ColumnCount; col++)
            {
                var cell = table.GetCell(1, col);
                cell?.SetBorder(BorderEdge.Bottom, borderStyle.Value);
            }
        }

        /// <summary>
        /// 设置末行下边框
        /// </summary>
        private void SetLastRowBottomBorder(ITableContext table, RowStyle headerStyle)
        {
            if (table == null || table.RowCount == 0) return;

            int lastRow = table.RowCount;
            var borderStyle = headerStyle.BottomBorder ?? headerStyle.TopBorder;
            if (!borderStyle.HasValue) return;

            for (int col = 1; col <= table.ColumnCount; col++)
            {
                var cell = table.GetCell(lastRow, col);
                if (cell != null)
                {
                    cell.SetBorder(BorderEdge.Bottom, borderStyle.Value);
                }
            }
        }

        public void DistributeColumnWidths(ITableContext table)
        {
            if (table == null || table.ColumnCount == 0) return;

            // 计算总宽度并平均分配
            float totalWidth = 0;
            for (int col = 1; col <= table.ColumnCount; col++)
            {
                totalWidth += table.GetColumnWidth(col);
            }

            float avgWidth = totalWidth / table.ColumnCount;
            for (int col = 1; col <= table.ColumnCount; col++)
            {
                table.SetColumnWidth(col, avgWidth);
            }

            _logger.LogInformation($"列宽已均匀分布，每列宽度: {avgWidth}");
        }

        public void DistributeRowHeights(ITableContext table)
        {
            if (table == null || table.RowCount == 0) return;

            // 计算总高度并平均分配
            float totalHeight = 0;
            for (int row = 1; row <= table.RowCount; row++)
            {
                totalHeight += table.GetRowHeight(row);
            }

            float avgHeight = totalHeight / table.RowCount;
            for (int row = 1; row <= table.RowCount; row++)
            {
                table.SetRowHeight(row, avgHeight);
            }

            _logger.LogInformation($"行高已均匀分布，每行高度: {avgHeight}");
        }

        private TableFormatOptions CreateThreeLineOptions()
        {
            var headerBorder = BorderStyle.SolidTheme(13, 1.75f);
            return new TableFormatOptions
            {
                FormatHeader = true,
                FormatDataRows = true,
                ApplyBorders = true,
                ApplyFont = true,
                ApplyTableStyle = false,
                Settings = new TableSettings
                {
                    FirstRow = true,
                    FirstCol = false,
                    LastRow = false,
                    LastCol = false,
                    HorizBanding = false,
                    VertBanding = false
                },
                HeaderStyle = new RowStyle
                {
                    HideBackground = true,
                    FontName = "等线",
                    FontNameFarEast = "等线",
                    FontSize = 12,
                    Bold = true,
                    Alignment = TextAlignment.Center,
                    TopBorder = headerBorder,
                    BottomBorder = headerBorder,
                    LeftBorder = BorderStyle.None,
                    RightBorder = BorderStyle.None
                },
                DataRowStyle = new RowStyle
                {
                    HideBackground = true,
                    FontName = "等线",
                    FontNameFarEast = "等线",
                    FontSize = 11,
                    Bold = false,
                    Alignment = TextAlignment.Center,
                    TopBorder = BorderStyle.None,
                    BottomBorder = BorderStyle.None,
                    LeftBorder = BorderStyle.None,
                    RightBorder = BorderStyle.None
                }
            };
        }
    }
}
