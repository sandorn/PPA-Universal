using System.Collections.Generic;

namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 选择上下文接口
    /// </summary>
    public interface ISelectionContext
    {
        /// <summary>选择类型</summary>
        SelectionType Type { get; }

        /// <summary>是否有选中内容</summary>
        bool HasSelection { get; }

        /// <summary>选中的形状数量</summary>
        int ShapeCount { get; }

        /// <summary>获取选中的形状</summary>
        IEnumerable<IShapeContext> SelectedShapes { get; }

        /// <summary>获取原生选择对象</summary>
        object NativeSelection { get; }
    }

    /// <summary>
    /// 选择类型枚举
    /// </summary>
    public enum SelectionType
    {
        /// <summary>无选择</summary>
        None = 0,

        /// <summary>幻灯片选择</summary>
        Slides = 1,

        /// <summary>形状选择</summary>
        Shapes = 2,

        /// <summary>文本选择</summary>
        Text = 3
    }
}
