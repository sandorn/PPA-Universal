using PPA.Business.Abstractions;  
using PPA.Core.Abstraction;  
using PPA.Core.Configuration;  
using PPA.Logging;  
using System;  
  
namespace PPA.Business.Services  
{  
    /// <summary>  
    /// 表格格式化服务实现  
    /// </summary>  
    public class TableFormatService : ITableFormatService  
    {  
        private readonly ILogger _logger;  
        private readonly ITableOperations _tableOps;  
        private readonly PPAConfig _config;  
  
        public TableFormatService(ILogger logger, ITableOperations tableOps, PPAConfig config)  
        {  
            _logger = logger ?? NullLogger.Instance;  
            _tableOps = tableOps;  
            _config = config;  
        }  
  
        public void FormatTable(ITableContext table, TableFormatOptions options = null)  
        {  
            if (table == null)  
            {  
                _logger.LogWarning("表格对象为空，跳过格式化");  
                return;  
            }  
  
            options ??= new TableFormatOptions();  
            _logger.LogInformation($"开始格式化表格，行数: {table.RowCount}, 列数: {table.ColumnCount}");  
  
            try  
            {  
                // 应用表格样式（PowerPoint 上可使用 No Style，WPS 暂不生效）  
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
  
        public void FormatTableAsThreeLine(ITableContext table)  
        {  
            if (table == null)  
            {  
                _logger.LogWarning("表格对象为空，无法应用三线表格式");  
                return;  
            }  
 
            var options = CreateThreeLineOptions();  
            var hasTableConfig = _config?.Table != null;  
            _logger.LogInformation($"开始应用三线表格式（来自配置），是否存在 Table 配置: {hasTableConfig}");   

            // WPS 中不再套用任何表格样式，仅依靠 ClearMenu + 三线表逻辑；
            // PowerPoint 中仍然可以通过 StyleId 套用一个“无样式、无网格”的基础样式。
            if (IsWpsTable(table))
            {
                options.ApplyTableStyle = false;
                options.TableStyleId = null;
            }

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
  
                // 1. 清空现有边框  
                cell.SetBorder(BorderEdge.All, BorderStyle.None);  
  
                // 2. 清理背景色：统一隐藏背景，不再根据行样式配置颜色  
                cell.SetBackgroundVisible(false);  
  
                // 3. 设置字体  
                var fontStyle = style.ToFontStyle();  
                if (!string.IsNullOrEmpty(fontStyle.Name) || fontStyle.Size > 0)  
                {  
                    cell.SetFont(fontStyle);  
                }  
  
                // 4. 设置对齐  
                if (style.Alignment.HasValue)  
                {  
                    cell.SetAlignment(style.Alignment.Value);  
                }  
  
                // 5. 设置边框  
                if (style.TopBorder.HasValue) cell.SetBorder(BorderEdge.Top, style.TopBorder.Value);  
                if (style.BottomBorder.HasValue) cell.SetBorder(BorderEdge.Bottom, style.BottomBorder.Value);  
                if (style.LeftBorder.HasValue) cell.SetBorder(BorderEdge.Left, style.LeftBorder.Value);  
                if (style.RightBorder.HasValue) cell.SetBorder(BorderEdge.Right, style.RightBorder.Value);  
            }  
        }  
  
        private bool IsWpsTable(ITableContext table)  
        {  
            if (table == null) return false;  
  
            try  
            {  
                var native = table.NativeTable;  
                if (native == null) return false;  
  
                var type = native.GetType();  
                var typeName = type.FullName ?? type.Name ?? string.Empty;  
  
                if (string.IsNullOrEmpty(typeName)) return false;  
  
                return typeName.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) >= 0  
                       || typeName.IndexOf("WPP", StringComparison.OrdinalIgnoreCase) >= 0  
                       || typeName.IndexOf("Kingsoft", StringComparison.OrdinalIgnoreCase) >= 0;  
            }  
            catch  
            {  
                return false;  
            }  
        }  
  
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
            _logger.LogInformation($"列宽已均匀分布: {avgWidth}");  
        }  
  
        public void DistributeRowHeights(ITableContext table)  
        {  
            if (table == null || table.RowCount == 0) return;  
  
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
            _logger.LogInformation($"行高已均匀分布: {avgHeight}");  
        }  
  
        private TableFormatOptions CreateThreeLineOptions()   
        {   
            var tableConfig = _config?.Table;  

            if (tableConfig == null)  
            {  
                // 正常情况下不应发生：PPAConfig.LoadOrCreate 已保证解析失败时重写默认配置
                _logger.LogError("CreateThreeLineOptions: Table 配置为空，已跳过三线表参数构建。请检查 PPAConfig.xml 是否被手动破坏。");  
                return new TableFormatOptions();  
            }  

            // 从配置映射 TableFormatOptions  
            var headerFont = tableConfig.HeaderRowFont;  
            var dataFont = tableConfig.DataRowFont;  
  
            // 表格设置（可为空，使用默认 false）  
            var settingsConfig = tableConfig.TableSettings;  
            var settings = settingsConfig == null  
                ? new TableSettings()  
                : new TableSettings  
                {  
                    FirstRow = settingsConfig.FirstRow,  
                    FirstCol = settingsConfig.FirstCol,  
                    LastRow = settingsConfig.LastRow,  
                    LastCol = settingsConfig.LastCol,  
                    HorizBanding = settingsConfig.HorizBanding,  
                    VertBanding = settingsConfig.VertBanding   
                };  
  
            // 主题色：优先用配置的 *ColorIndex，没有就留空让底层用默认  
            BorderStyle? headerBorder = null;  
            if (tableConfig.HeaderRowBorderColorIndex.HasValue)  
            {  
                headerBorder = BorderStyle.SolidTheme(  
                    tableConfig.HeaderRowBorderColorIndex.Value,  
                    tableConfig.HeaderRowBorderWidth);  
            }  
  
            BorderStyle? dataTopBorder = null;  
            if (tableConfig.DataRowBorderColorIndex.HasValue)  
            {  
                dataTopBorder = BorderStyle.SolidTheme(  
                    tableConfig.DataRowBorderColorIndex.Value,  
                    tableConfig.DataRowBorderWidth);  
            }  
  
            BorderStyle? finalRowBorder = null;  
            if (tableConfig.FinalRowBorderColorIndex.HasValue)  
            {  
                finalRowBorder = BorderStyle.SolidTheme(  
                    tableConfig.FinalRowBorderColorIndex.Value,  
                    tableConfig.FinalRowBorderWidth);  
            }  
  
            // 三线表目前不绘制竖向边框，只保留三条横线
            var headerStyle = new RowStyle  
            {  
                FontName = headerFont?.Name,  
                FontNameFarEast = headerFont?.NameFarEast,  
                FontSize = headerFont?.Size,  
                Bold = headerFont?.Bold,  
                ThemeColorIndex = headerFont?.ThemeColorIndex,  
                Alignment = TextAlignment.Center,  
                TopBorder = headerBorder,  
                BottomBorder = headerBorder,  
                LeftBorder = BorderStyle.None,  
                RightBorder = BorderStyle.None   
            };  

            var dataRowStyle = new RowStyle  
            {  
                FontName = dataFont?.Name,  
                FontNameFarEast = dataFont?.NameFarEast,  
                FontSize = dataFont?.Size,  
                Bold = dataFont?.Bold,  
                ThemeColorIndex = dataFont?.ThemeColorIndex,  
                Alignment = TextAlignment.Center,  
                TopBorder = dataTopBorder,  
                BottomBorder = null,  
                LeftBorder = BorderStyle.None,  
                RightBorder = BorderStyle.None   
            };  
  
            var options = new TableFormatOptions  
            {  
                FormatHeader = true,  
                FormatDataRows = true,  
                ApplyBorders = true,  
                ApplyFont = true,  
                // 默认根据配置中的 StyleId 决定是否套用表格样式；
                // 具体是否执行由调用方（如 WPS 分支）再做覆盖。
                ApplyTableStyle = !string.IsNullOrEmpty(tableConfig.StyleId),   
                TableStyleId = tableConfig.StyleId,   
                Settings = settings,   
                HeaderStyle = headerStyle,  
                DataRowStyle = dataRowStyle,  
                AutoNumberFormat = tableConfig.AutoNumberFormat,  
                DecimalPlaces = tableConfig.DecimalPlaces,  
                NegativeTextColor = tableConfig.NegativeTextColor,
                FinalRowBottomBorder = finalRowBorder ?? headerBorder  
            };  
            return options;  
        }  
    }  
}  
