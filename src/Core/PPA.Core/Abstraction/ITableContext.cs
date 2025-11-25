namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 表格上下文接口
    /// </summary>
    public interface ITableContext
    {
        /// <summary>行数</summary>
        int RowCount { get; }

        /// <summary>列数</summary>
        int ColumnCount { get; }

        /// <summary>获取单元格</summary>
        /// <param name="row">行索引（1-based）</param>
        /// <param name="column">列索引（1-based）</param>
        ICellContext GetCell(int row, int column);

        /// <summary>获取行高</summary>
        /// <param name="row">行索引（1-based）</param>
        float GetRowHeight(int row);

        /// <summary>设置行高</summary>
        /// <param name="row">行索引（1-based）</param>
        /// <param name="height">高度</param>
        void SetRowHeight(int row, float height);

        /// <summary>获取列宽</summary>
        /// <param name="column">列索引（1-based）</param>
        float GetColumnWidth(int column);

        /// <summary>设置列宽</summary>
        /// <param name="column">列索引（1-based）</param>
        /// <param name="width">宽度</param>
        void SetColumnWidth(int column, float width);

        /// <summary>获取原生表格对象</summary>
        object NativeTable { get; }
    }
}
