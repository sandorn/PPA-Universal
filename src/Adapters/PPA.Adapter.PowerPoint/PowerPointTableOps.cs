using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 表格操作实现
    /// </summary>
    public class PowerPointTableOps : ITableOperations
    {
        public int GetRowCount(object table)
        {
            var netTable = table as NETOP.Table;
            return netTable?.Rows?.Count ?? 0;
        }

        public int GetColumnCount(object table)
        {
            var netTable = table as NETOP.Table;
            return netTable?.Columns?.Count ?? 0;
        }

        public object GetCell(object table, int row, int col)
        {
            var netTable = table as NETOP.Table;
            if (netTable == null) return null;

            try
            {
                return netTable.Cell(row, col);
            }
            catch
            {
                return null;
            }
        }

        public void SetCellText(object cell, string text)
        {
            var netCell = cell as NETOP.Cell;
            if (netCell == null) return;

            try
            {
                var textRange = netCell.Shape?.TextFrame?.TextRange;
                if (textRange != null)
                {
                    textRange.Text = text ?? string.Empty;
                }
            }
            catch { }
        }

        public string GetCellText(object cell)
        {
            var netCell = cell as NETOP.Cell;
            if (netCell == null) return string.Empty;

            try
            {
                return netCell.Shape?.TextFrame?.TextRange?.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public void SetCellBackground(object cell, int colorRgb)
        {
            var netCell = cell as NETOP.Cell;
            if (netCell == null) return;

            try
            {
                var fill = netCell.Shape?.Fill;
                if (fill == null) return;

                fill.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
                fill.Solid();
                fill.ForeColor.RGB = colorRgb;
            }
            catch { }
        }

        public int GetCellBackground(object cell)
        {
            var netCell = cell as NETOP.Cell;
            if (netCell == null) return 0;

            try
            {
                return netCell.Shape?.Fill?.ForeColor?.RGB ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        public void SetBorderStyle(object cell, BorderEdge edge, BorderStyle style)
        {
            var netCell = cell as NETOP.Cell;
            if (netCell == null) return;

            try
            {
                var borders = netCell.Borders;
                if (borders == null) return;

                if (edge == BorderEdge.All)
                {
                    SetSingleBorder(borders[NETOP.Enums.PpBorderType.ppBorderLeft], style);
                    SetSingleBorder(borders[NETOP.Enums.PpBorderType.ppBorderTop], style);
                    SetSingleBorder(borders[NETOP.Enums.PpBorderType.ppBorderRight], style);
                    SetSingleBorder(borders[NETOP.Enums.PpBorderType.ppBorderBottom], style);
                }
                else
                {
                    var borderType = ConvertEdgeToType(edge);
                    SetSingleBorder(borders[borderType], style);
                }
            }
            catch { }
        }

        private void SetSingleBorder(NETOP.LineFormat border, BorderStyle style)
        {
            if (border == null) return;

            if (!style.Visible)
            {
                border.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoFalse;
                return;
            }

            border.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
            border.Weight = style.Weight;
            border.ForeColor.RGB = style.Color;
        }

        private NETOP.Enums.PpBorderType ConvertEdgeToType(BorderEdge edge)
        {
            switch (edge)
            {
                case BorderEdge.Left: return NETOP.Enums.PpBorderType.ppBorderLeft;
                case BorderEdge.Top: return NETOP.Enums.PpBorderType.ppBorderTop;
                case BorderEdge.Right: return NETOP.Enums.PpBorderType.ppBorderRight;
                case BorderEdge.Bottom: return NETOP.Enums.PpBorderType.ppBorderBottom;
                default: return NETOP.Enums.PpBorderType.ppBorderLeft;
            }
        }

        public float GetRowHeight(object table, int row)
        {
            var netTable = table as NETOP.Table;
            if (netTable == null) return 0;

            try
            {
                return netTable.Rows[row].Height;
            }
            catch
            {
                return 0;
            }
        }

        public void SetRowHeight(object table, int row, float height)
        {
            var netTable = table as NETOP.Table;
            if (netTable == null) return;

            try
            {
                netTable.Rows[row].Height = height;
            }
            catch { }
        }

        public float GetColumnWidth(object table, int column)
        {
            var netTable = table as NETOP.Table;
            if (netTable == null) return 0;

            try
            {
                return netTable.Columns[column].Width;
            }
            catch
            {
                return 0;
            }
        }

        public void SetColumnWidth(object table, int column, float width)
        {
            var netTable = table as NETOP.Table;
            if (netTable == null) return;

            try
            {
                netTable.Columns[column].Width = width;
            }
            catch { }
        }

        public void MergeCells(object table, int startRow, int startCol, int endRow, int endCol)
        {
            var netTable = table as NETOP.Table;
            if (netTable == null) return;

            try
            {
                netTable.Cell(startRow, startCol).Merge(netTable.Cell(endRow, endCol));
            }
            catch { }
        }

        public void SplitCell(object cell, int rows, int cols)
        {
            var netCell = cell as NETOP.Cell;
            if (netCell == null) return;

            try
            {
                netCell.Split(rows, cols);
            }
            catch { }
        }
    }
}
