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
                HeaderStyle = new RowStyle
                {
                    HideBackground = true, // 隐藏背景，使用主题色
                    FontName = "等线",
                    FontNameFarEast = "等线",
                    FontSize = 12,
                    Bold = true,
                    ThemeColorIndex = 13, // 深色1 (dk1)
                    Alignment = TextAlignment.Center,
                    Border = BorderStyle.Solid(0x2B579A, 1.5f) // 蓝色边框
                },
                DataRowStyle = new RowStyle
                {
                    HideBackground = true, // 隐藏背景
                    FontName = "等线",
                    FontNameFarEast = "等线",
                    FontSize = 11,
                    Bold = false,
                    ThemeColorIndex = 13, // 深色1 (dk1)
                    Border = BorderStyle.Solid(0xD0D0D0, 0.75f) // 浅灰边框
                }
            };

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
                if (style.Border.HasValue)
                {
                    cell.SetBorder(BorderEdge.All, style.Border.Value);
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
    }
}
