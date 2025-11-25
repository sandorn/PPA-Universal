namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 平台支持的功能特性枚举
    /// </summary>
    public enum Feature
    {
        /// <summary>基础表格操作</summary>
        TableBasic = 1,

        /// <summary>表格边框高级样式</summary>
        TableAdvancedBorder = 2,

        /// <summary>图表操作</summary>
        Chart = 10,

        /// <summary>图表高级格式化</summary>
        ChartAdvanced = 11,

        /// <summary>形状对齐</summary>
        ShapeAlignment = 20,

        /// <summary>形状批量操作</summary>
        ShapeBatch = 21,

        /// <summary>文本高级格式化</summary>
        TextAdvanced = 30,

        /// <summary>撤销/重做支持</summary>
        UndoRedo = 40,

        /// <summary>快捷键支持</summary>
        Shortcuts = 50
    }
}
