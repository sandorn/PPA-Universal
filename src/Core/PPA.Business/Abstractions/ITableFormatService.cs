using PPA.Core.Abstraction;

namespace PPA.Business.Abstractions
{
    /// <summary>
    /// 表格格式化服务接口（平台无关）
    /// </summary>
    public interface ITableFormatService
    {
        /// <summary>
        /// 格式化单个表格
        /// </summary>
        /// <param name="table">表格上下文</param>
        /// <param name="options">格式化选项</param>
        void FormatTable(ITableContext table, TableFormatOptions options = null);

        /// <summary>
        /// 格式化表格（使用默认配置）
        /// </summary>
        void FormatTableWithDefaults(ITableContext table);

        /// <summary>
        /// 设置表格行样式
        /// </summary>
        /// <param name="table">表格上下文</param>
        /// <param name="rowIndex">行索引（1-based）</param>
        /// <param name="style">行样式</param>
        void SetRowStyle(ITableContext table, int rowIndex, RowStyle style);

        /// <summary>
        /// 设置表格列宽度均匀分布
        /// </summary>
        void DistributeColumnWidths(ITableContext table);

        /// <summary>
        /// 设置表格行高均匀分布
        /// </summary>
        void DistributeRowHeights(ITableContext table);
    }

    /// <summary>
    /// 表格格式化选项
    /// </summary>
    public class TableFormatOptions
    {
        /// <summary>是否格式化表头</summary>
        public bool FormatHeader { get; set; } = true;

        /// <summary>是否格式化数据行</summary>
        public bool FormatDataRows { get; set; } = true;

        /// <summary>是否应用边框样式</summary>
        public bool ApplyBorders { get; set; } = true;

        /// <summary>是否应用字体样式</summary>
        public bool ApplyFont { get; set; } = true;

        /// <summary>表头样式</summary>
        public RowStyle HeaderStyle { get; set; }

        /// <summary>数据行样式</summary>
        public RowStyle DataRowStyle { get; set; }

        /// <summary>交替行样式（可选）</summary>
        public RowStyle AlternateRowStyle { get; set; }
    }

    /// <summary>
    /// 行样式
    /// </summary>
    public class RowStyle
    {
        /// <summary>背景色（RGB）</summary>
        public int? BackgroundColor { get; set; }

        /// <summary>是否隐藏背景</summary>
        public bool HideBackground { get; set; }

        /// <summary>字体名称（西文）</summary>
        public string FontName { get; set; }

        /// <summary>字体名称（中文/远东）</summary>
        public string FontNameFarEast { get; set; }

        /// <summary>字体大小</summary>
        public float? FontSize { get; set; }

        /// <summary>字体颜色（RGB）</summary>
        public int? FontColor { get; set; }

        /// <summary>主题颜色索引</summary>
        public int? ThemeColorIndex { get; set; }

        /// <summary>是否加粗</summary>
        public bool? Bold { get; set; }

        /// <summary>文本对齐方式</summary>
        public TextAlignment? Alignment { get; set; }

        /// <summary>边框样式</summary>
        public BorderStyle? Border { get; set; }

        /// <summary>
        /// 转换为 FontStyle
        /// </summary>
        public FontStyle ToFontStyle()
        {
            return new FontStyle
            {
                Name = FontName,
                NameFarEast = FontNameFarEast,
                Size = FontSize ?? 11,
                Bold = Bold ?? false,
                ColorRgb = FontColor,
                ThemeColorIndex = ThemeColorIndex
            };
        }
    }
}
