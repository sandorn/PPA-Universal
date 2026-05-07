using System.Collections.Generic;
using PPA.Core.Abstraction;

namespace PPA.Business.Abstractions
{
    /// <summary>
    /// 形状复制服务接口
    /// </summary>
    public interface IShapeDuplicateService
    {
        /// <summary>
        /// 矩阵复制：按行列网格排列复制形状
        /// </summary>
        /// <param name="shapes">要复制的形状集合</param>
        /// <param name="rows">行数</param>
        /// <param name="columns">列数</param>
        /// <param name="rowSpacing">行间距（像素）</param>
        /// <param name="columnSpacing">列间距（像素）</param>
        /// <returns>创建的复制形状集合</returns>
        IEnumerable<IShapeContext> MatrixCopy(IEnumerable<IShapeContext> shapes, int rows, int columns, float rowSpacing, float columnSpacing);

        /// <summary>
        /// 线性复制：按指定方向和间距复制形状
        /// </summary>
        /// <param name="shapes">要复制的形状集合</param>
        /// <param name="count">复制数量</param>
        /// <param name="spacing">间距（像素）</param>
        /// <param name="direction">复制方向</param>
        /// <returns>创建的复制形状集合</returns>
        IEnumerable<IShapeContext> LinearCopy(IEnumerable<IShapeContext> shapes, int count, float spacing, LinearCopyDirection direction);
    }

    /// <summary>
    /// 线性复制方向
    /// </summary>
    public enum LinearCopyDirection
    {
        /// <summary>水平方向（向右）</summary>
        Horizontal,
        /// <summary>垂直方向（向下）</summary>
        Vertical
    }
}

