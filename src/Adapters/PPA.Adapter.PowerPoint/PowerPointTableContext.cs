using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 表格上下文实现
    /// </summary>
    public class PowerPointTableContext : ITableContext
    {
        private readonly NETOP.Table _table;

        public PowerPointTableContext(NETOP.Table table)
        {
            _table = table;
        }

        public int RowCount => _table?.Rows?.Count ?? 0;

        public int ColumnCount => _table?.Columns?.Count ?? 0;

        public ICellContext GetCell(int row, int column)
        {
            try
            {
                if (row < 1 || row > RowCount || column < 1 || column > ColumnCount)
                    return null;

                var cell = _table.Cell(row, column);
                return cell != null ? new PowerPointCellContext(cell, row, column) : null;
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
                return _table.Rows[row].Height;
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
                return _table.Columns[column].Width;
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
        /// 获取原生 Table 对象
        /// </summary>
        public NETOP.Table NetTable => _table;
    }
}
