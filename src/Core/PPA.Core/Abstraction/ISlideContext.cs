using System.Collections.Generic;

namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 幻灯片上下文接口
    /// </summary>
    public interface ISlideContext
    {
        /// <summary>幻灯片索引（1-based）</summary>
        int SlideIndex { get; }

        /// <summary>幻灯片编号</summary>
        int SlideNumber { get; }

        /// <summary>幻灯片中的形状数量</summary>
        int ShapeCount { get; }

        /// <summary>获取所有形状</summary>
        IEnumerable<IShapeContext> Shapes { get; }

        /// <summary>获取原生幻灯片对象</summary>
        object NativeSlide { get; }
    }
}
