namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 形状上下文接口
    /// </summary>
    public interface IShapeContext
    {
        /// <summary>形状名称</summary>
        string Name { get; set; }

        /// <summary>形状类型</summary>
        ShapeType ShapeType { get; }

        /// <summary>形状位置和尺寸</summary>
        ShapeRect Bounds { get; set; }

        /// <summary>旋转角度</summary>
        float Rotation { get; set; }

        /// <summary>是否为表格</summary>
        bool IsTable { get; }

        /// <summary>是否为图表</summary>
        bool IsChart { get; }

        /// <summary>是否有文本框架</summary>
        bool HasTextFrame { get; }

        /// <summary>获取表格接口（如果是表格）</summary>
        ITableContext Table { get; }

        /// <summary>获取原生形状对象</summary>
        object NativeShape { get; }
    }

    /// <summary>
    /// 形状类型枚举
    /// </summary>
    public enum ShapeType
    {
        /// <summary>自动形状</summary>
        AutoShape = 1,

        /// <summary>文本框</summary>
        TextBox = 17,

        /// <summary>表格</summary>
        Table = 19,

        /// <summary>图表</summary>
        Chart = 3,

        /// <summary>图片</summary>
        Picture = 13,

        /// <summary>组合</summary>
        Group = 6,

        /// <summary>占位符</summary>
        Placeholder = 14,

        /// <summary>媒体</summary>
        Media = 16,

        /// <summary>其他</summary>
        Other = 0
    }

    /// <summary>
    /// 形状位置和尺寸结构
    /// </summary>
    public struct ShapeRect
    {
        /// <summary>左边距</summary>
        public float Left { get; set; }

        /// <summary>上边距</summary>
        public float Top { get; set; }

        /// <summary>宽度</summary>
        public float Width { get; set; }

        /// <summary>高度</summary>
        public float Height { get; set; }

        /// <summary>右边界</summary>
        public float Right => Left + Width;

        /// <summary>下边界</summary>
        public float Bottom => Top + Height;

        /// <summary>中心 X 坐标</summary>
        public float CenterX => Left + Width / 2;

        /// <summary>中心 Y 坐标</summary>
        public float CenterY => Top + Height / 2;

        public ShapeRect(float left, float top, float width, float height)
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }
    }
}
