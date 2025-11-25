using System.Collections.Generic;

namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 幻灯片操作接口（平台无关）
    /// </summary>
    public interface ISlideOperations
    {
        /// <summary>获取幻灯片中的所有形状</summary>
        IEnumerable<object> GetShapes(object slide);

        /// <summary>获取幻灯片中的形状数量</summary>
        int GetShapeCount(object slide);

        /// <summary>在幻灯片中添加形状</summary>
        object AddShape(object slide, ShapeType type, ShapeRect bounds);

        /// <summary>在幻灯片中添加表格</summary>
        object AddTable(object slide, int rows, int columns, ShapeRect bounds);

        /// <summary>复制幻灯片</summary>
        object DuplicateSlide(object slide);

        /// <summary>删除幻灯片</summary>
        void DeleteSlide(object slide);

        /// <summary>获取幻灯片索引</summary>
        int GetSlideIndex(object slide);

        /// <summary>移动幻灯片到指定位置</summary>
        void MoveSlide(object slide, int newIndex);
    }
}
