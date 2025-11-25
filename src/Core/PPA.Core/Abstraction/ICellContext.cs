namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 单元格上下文接口
    /// </summary>
    public interface ICellContext
    {
        /// <summary>行索引（1-based）</summary>
        int Row { get; }

        /// <summary>列索引（1-based）</summary>
        int Column { get; }

        /// <summary>单元格文本</summary>
        string Text { get; set; }

        /// <summary>是否合并单元格</summary>
        bool IsMerged { get; }

        /// <summary>设置背景色</summary>
        /// <param name="colorRgb">RGB 颜色值</param>
        void SetBackground(int colorRgb);

        /// <summary>设置背景可见性</summary>
        void SetBackgroundVisible(bool visible);

        /// <summary>获取背景色</summary>
        int GetBackground();

        /// <summary>设置边框样式</summary>
        /// <param name="edge">边框边缘</param>
        /// <param name="style">边框样式</param>
        void SetBorder(BorderEdge edge, BorderStyle style);

        /// <summary>设置字体样式</summary>
        void SetFont(FontStyle style);

        /// <summary>设置文本对齐方式</summary>
        void SetAlignment(TextAlignment alignment);

        /// <summary>获取原生单元格对象</summary>
        object NativeCell { get; }
    }

    /// <summary>
    /// 字体样式
    /// </summary>
    public class FontStyle
    {
        /// <summary>字体名称（西文）</summary>
        public string Name { get; set; }

        /// <summary>字体名称（中文/远东）</summary>
        public string NameFarEast { get; set; }

        /// <summary>字体大小</summary>
        public float Size { get; set; } = 11;

        /// <summary>是否粗体</summary>
        public bool Bold { get; set; }

        /// <summary>是否斜体</summary>
        public bool Italic { get; set; }

        /// <summary>字体颜色（RGB）</summary>
        public int? ColorRgb { get; set; }

        /// <summary>主题颜色索引（可选）</summary>
        public int? ThemeColorIndex { get; set; }
    }

    /// <summary>
    /// 文本对齐方式
    /// </summary>
    public enum TextAlignment
    {
        /// <summary>左对齐</summary>
        Left = 1,

        /// <summary>居中</summary>
        Center = 2,

        /// <summary>右对齐</summary>
        Right = 3,

        /// <summary>两端对齐</summary>
        Justify = 4
    }

    /// <summary>
    /// 边框边缘枚举
    /// </summary>
    public enum BorderEdge
    {
        /// <summary>左边框</summary>
        Left = 1,

        /// <summary>上边框</summary>
        Top = 2,

        /// <summary>右边框</summary>
        Right = 3,

        /// <summary>下边框</summary>
        Bottom = 4,

        /// <summary>所有边框</summary>
        All = 5
    }

    /// <summary>
    /// 边框样式
    /// </summary>
    public struct BorderStyle
    {
        /// <summary>边框宽度</summary>
        public float Weight { get; set; }

        /// <summary>边框颜色（RGB）</summary>
        public int Color { get; set; }

        /// <summary>主题颜色索引（优先于 RGB）</summary>
        public int? ThemeColorIndex { get; set; }

        /// <summary>边框可见性</summary>
        public bool Visible { get; set; }

        /// <summary>边框线型</summary>
        public BorderLineStyle LineStyle { get; set; }

        public static BorderStyle None => new BorderStyle { Visible = false };

        /// <summary>使用 RGB 颜色创建实线边框</summary>
        public static BorderStyle Solid(int color, float weight = 1.0f)
        {
            return new BorderStyle
            {
                Visible = true,
                Color = color,
                Weight = weight,
                LineStyle = BorderLineStyle.Solid
            };
        }

        /// <summary>使用主题色创建实线边框</summary>
        public static BorderStyle SolidTheme(int themeColorIndex, float weight = 1.0f)
        {
            return new BorderStyle
            {
                Visible = true,
                ThemeColorIndex = themeColorIndex,
                Weight = weight,
                LineStyle = BorderLineStyle.Solid
            };
        }
    }

    /// <summary>
    /// 边框线型枚举
    /// </summary>
    public enum BorderLineStyle
    {
        /// <summary>实线</summary>
        Solid = 1,

        /// <summary>虚线</summary>
        Dash = 2,

        /// <summary>点线</summary>
        Dot = 3,

        /// <summary>点划线</summary>
        DashDot = 4,

        /// <summary>双点划线</summary>
        DashDotDot = 5
    }
}
