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
        /// 将表格格式化为标准三线表样式
        /// </summary>
        /// <param name="table">表格上下文</param>
        void FormatTableAsThreeLine(ITableContext table);

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

        /// <summary>是否应用表格样式</summary>
        public bool ApplyTableStyle { get; set; } = true;

        /// <summary>表格样式 ID（GUID 格式）</summary>
        public string TableStyleId { get; set; }

        /// <summary>表格设置</summary>
        public TableSettings Settings { get; set; }

        /// <summary>表头样式</summary>
        public RowStyle HeaderStyle { get; set; }

        /// <summary>数据行样式</summary>
        public RowStyle DataRowStyle { get; set; }

        /// <summary>交替行样式（可选）</summary>
        public RowStyle AlternateRowStyle { get; set; }

        /// <summary>是否启用数字格式化</summary>
        public bool AutoNumberFormat { get; set; } = true;

        /// <summary>数字格式化保留的小数位数</summary>
        public int DecimalPlaces { get; set; } = 0;

        /// <summary>负数文本颜色（RGB）</summary>
        public int NegativeTextColor { get; set; } = 255; // 红色

        /// <summary>末行下边框样式（如未指定则回退到表头样式）</summary>
        public BorderStyle? FinalRowBottomBorder { get; set; }
    }

    /// <summary>
    /// 表格设置
    /// </summary>
    public class TableSettings
    {
        /// <summary>首行特殊格式</summary>
        public bool FirstRow { get; set; } = true;

        /// <summary>首列特殊格式</summary>
        public bool FirstCol { get; set; } = false;

        /// <summary>末行特殊格式</summary>
        public bool LastRow { get; set; } = false;

        /// <summary>末列特殊格式</summary>
        public bool LastCol { get; set; } = false;

        /// <summary>水平条带</summary>
        public bool HorizBanding { get; set; } = false;

        /// <summary>垂直条带</summary>
        public bool VertBanding { get; set; } = false;
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

        /// <summary>上边框样式</summary>
        public BorderStyle? TopBorder { get; set; }

        /// <summary>下边框样式</summary>
        public BorderStyle? BottomBorder { get; set; }

        /// <summary>左边框样式</summary>
        public BorderStyle? LeftBorder { get; set; }

        /// <summary>右边框样式</summary>
        public BorderStyle? RightBorder { get; set; }

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
