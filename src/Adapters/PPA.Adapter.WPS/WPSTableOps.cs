using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 表格操作实现
    /// </summary>
    public class WPSTableOps : ITableOperations
    {
        public int GetRowCount(object table)
        {
            dynamic dynTable = table;
            try { return dynTable?.Rows?.Count ?? 0; }
            catch { return 0; }
        }

        public int GetColumnCount(object table)
        {
            dynamic dynTable = table;
            try { return dynTable?.Columns?.Count ?? 0; }
            catch { return 0; }
        }

        public object GetCell(object table, int row, int col)
        {
            dynamic dynTable = table;
            if (dynTable == null) return null;

            try { return dynTable.Cell(row, col); }
            catch { return null; }
        }

        public void SetCellText(object cell, string text)
        {
            dynamic dynCell = cell;
            if (dynCell == null) return;

            try
            {
                dynamic textRange = dynCell.Shape?.TextFrame?.TextRange;
                if (textRange != null)
                {
                    textRange.Text = text ?? string.Empty;
                }
            }
            catch { }
        }

        public string GetCellText(object cell)
        {
            dynamic dynCell = cell;
            if (dynCell == null) return string.Empty;

            try
            {
                return dynCell.Shape?.TextFrame?.TextRange?.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public void SetCellBackground(object cell, int colorRgb)
        {
            dynamic dynCell = cell;
            if (dynCell == null) return;

            try
            {
                dynamic fill = dynCell.Shape?.Fill;
                if (fill == null) return;

                fill.Visible = WPSHelper.TriState.True;
                fill.Solid();
                fill.ForeColor.RGB = colorRgb;
            }
            catch { }
        }

        public int GetCellBackground(object cell)
        {
            dynamic dynCell = cell;
            if (dynCell == null) return 0;

            try { return dynCell.Shape?.Fill?.ForeColor?.RGB ?? 0; }
            catch { return 0; }
        }

        public void SetBorderStyle(object cell, BorderEdge edge, BorderStyle style)
        {
            dynamic dynCell = cell;
            if (dynCell == null) return;

            try
            {
                dynamic borders = dynCell.Borders;
                if (borders == null) return;

                if (edge == BorderEdge.All)
                {
                    SetSingleBorder(borders[WPSHelper.BorderType.Left], style);
                    SetSingleBorder(borders[WPSHelper.BorderType.Top], style);
                    SetSingleBorder(borders[WPSHelper.BorderType.Right], style);
                    SetSingleBorder(borders[WPSHelper.BorderType.Bottom], style);
                }
                else
                {
                    int borderType = ConvertEdgeToType(edge);
                    SetSingleBorder(borders[borderType], style);
                }
            }
            catch { }
        }

        private void SetSingleBorder(dynamic border, BorderStyle style)
        {
            if (border == null) return;

            try
            {
                if (!style.Visible)
                {
                    border.Visible = WPSHelper.TriState.False;
                    return;
                }

                border.Visible = WPSHelper.TriState.True;
                border.Weight = style.Weight;
                border.ForeColor.RGB = style.Color;
            }
            catch { }
        }

        private int ConvertEdgeToType(BorderEdge edge)
        {
            switch (edge)
            {
                case BorderEdge.Left: return WPSHelper.BorderType.Left;
                case BorderEdge.Top: return WPSHelper.BorderType.Top;
                case BorderEdge.Right: return WPSHelper.BorderType.Right;
                case BorderEdge.Bottom: return WPSHelper.BorderType.Bottom;
                default: return WPSHelper.BorderType.Left;
            }
        }

        public float GetRowHeight(object table, int row)
        {
            dynamic dynTable = table;
            try { return (float)(dynTable?.Rows[row].Height ?? 0); }
            catch { return 0; }
        }

        public void SetRowHeight(object table, int row, float height)
        {
            dynamic dynTable = table;
            try { if (dynTable != null) dynTable.Rows[row].Height = height; }
            catch { }
        }

        public float GetColumnWidth(object table, int column)
        {
            dynamic dynTable = table;
            try { return (float)(dynTable?.Columns[column].Width ?? 0); }
            catch { return 0; }
        }

        public void SetColumnWidth(object table, int column, float width)
        {
            dynamic dynTable = table;
            try { if (dynTable != null) dynTable.Columns[column].Width = width; }
            catch { }
        }

        public void MergeCells(object table, int startRow, int startCol, int endRow, int endCol)
        {
            dynamic dynTable = table;
            if (dynTable == null) return;

            try
            {
                dynamic startCell = dynTable.Cell(startRow, startCol);
                dynamic endCell = dynTable.Cell(endRow, endCol);
                startCell?.Merge(endCell);
            }
            catch { }
        }

        public void SplitCell(object cell, int rows, int cols)
        {
            dynamic dynCell = cell;
            try { dynCell?.Split(rows, cols); }
            catch { }
        }
    }
}
