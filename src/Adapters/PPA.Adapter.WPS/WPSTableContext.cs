using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 表格上下文实现
    /// </summary>
    public class WPSTableContext : ITableContext
    {
        private readonly dynamic _table;

        public WPSTableContext(dynamic table)
        {
            _table = table;
        }

        public int RowCount
        {
            get
            {
                try { return _table?.Rows?.Count ?? 0; }
                catch { return 0; }
            }
        }

        public int ColumnCount
        {
            get
            {
                try { return _table?.Columns?.Count ?? 0; }
                catch { return 0; }
            }
        }

        public ICellContext GetCell(int row, int column)
        {
            try
            {
                if (row < 1 || row > RowCount || column < 1 || column > ColumnCount)
                    return null;

                dynamic cell = _table.Cell(row, column);
                return cell != null ? new WPSCellContext(cell, row, column) : null;
            }
            catch
            {
                return null;
            }
        }

        public float GetRowHeight(int row)
        {
            try
            {
                if (row < 1 || row > RowCount) return 0;
                return (float)(_table.Rows[row].Height ?? 0);
            }
            catch
            {
                return 0;
            }
        }

        public void SetRowHeight(int row, float height)
        {
            try
            {
                if (row < 1 || row > RowCount) return;
                _table.Rows[row].Height = height;
            }
            catch { }
        }

        public float GetColumnWidth(int column)
        {
            try
            {
                if (column < 1 || column > ColumnCount) return 0;
                return (float)(_table.Columns[column].Width ?? 0);
            }
            catch
            {
                return 0;
            }
        }

        public void SetColumnWidth(int column, float width)
        {
            try
            {
                if (column < 1 || column > ColumnCount) return;
                _table.Columns[column].Width = width;
            }
            catch { }
        }

        public void ApplyStyle(string styleId)
        {
            try
            {
                if (string.IsNullOrEmpty(styleId) || _table == null) return;
                _table.ApplyStyle(styleId, false);
            }
            catch { }
        }

        public void SetTableOptions(bool firstRow, bool firstCol, bool lastRow, bool lastCol, bool horizBanding, bool vertBanding)
        {
            try
            {
                if (_table == null) return;
                _table.FirstRow = firstRow;
                _table.FirstCol = firstCol;
                _table.LastRow = lastRow;
                _table.LastCol = lastCol;
                _table.HorizBanding = horizBanding;
                _table.VertBanding = vertBanding;
            }
            catch { }
        }

        public object NativeTable => _table;

        /// <summary>
        /// 获取 WPS Table 动态对象
        /// </summary>
        public dynamic Table => _table;
    }
}
