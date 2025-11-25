namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 表格操作接口（平台无关）
    /// </summary>
    public interface ITableOperations
    {
        /// <summary>获取表格行数</summary>
        int GetRowCount(object table);

        /// <summary>获取表格列数</summary>
        int GetColumnCount(object table);

        /// <summary>获取单元格</summary>
        object GetCell(object table, int row, int col);

        /// <summary>设置单元格文本</summary>
        void SetCellText(object cell, string text);

        /// <summary>获取单元格文本</summary>
        string GetCellText(object cell);

        /// <summary>设置单元格背景色</summary>
        void SetCellBackground(object cell, int colorRgb);

        /// <summary>获取单元格背景色</summary>
        int GetCellBackground(object cell);

        /// <summary>设置边框样式</summary>
        void SetBorderStyle(object cell, BorderEdge edge, BorderStyle style);

        /// <summary>获取行高</summary>
        float GetRowHeight(object table, int row);

        /// <summary>设置行高</summary>
        void SetRowHeight(object table, int row, float height);

        /// <summary>获取列宽</summary>
        float GetColumnWidth(object table, int column);

        /// <summary>设置列宽</summary>
        void SetColumnWidth(object table, int column, float width);

        /// <summary>合并单元格</summary>
        void MergeCells(object table, int startRow, int startCol, int endRow, int endCol);

        /// <summary>拆分单元格</summary>
        void SplitCell(object cell, int rows, int cols);
    }
}
