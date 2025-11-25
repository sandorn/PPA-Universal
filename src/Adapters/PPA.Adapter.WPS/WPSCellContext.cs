using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 单元格上下文实现
    /// </summary>
    public class WPSCellContext : ICellContext
    {
        private readonly dynamic _cell;
        private readonly int _row;
        private readonly int _column;

        public WPSCellContext(dynamic cell, int row, int column)
        {
            _cell = cell;
            _row = row;
            _column = column;
        }

        public int Row => _row;

        public int Column => _column;

        public string Text
        {
            get
            {
                try
                {
                    return _cell?.Shape?.TextFrame?.TextRange?.Text ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }
            set
            {
                try
                {
                    dynamic textRange = _cell?.Shape?.TextFrame?.TextRange;
                    if (textRange != null)
                    {
                        textRange.Text = value ?? string.Empty;
                    }
                }
                catch { }
            }
        }

        public bool IsMerged
        {
            get
            {
                // WPS 可能不直接提供 IsMerged 属性
                // 暂时返回 false
                return false;
            }
        }

        public void SetBackground(int colorRgb)
        {
            try
            {
                dynamic fill = _cell?.Shape?.Fill;
                if (fill == null) return;

                fill.Visible = WPSHelper.TriState.True;
                fill.Solid();
                fill.ForeColor.RGB = colorRgb;
            }
            catch { }
        }

        public int GetBackground()
        {
            try
            {
                return _cell?.Shape?.Fill?.ForeColor?.RGB ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        public void SetBorder(BorderEdge edge, BorderStyle style)
        {
            try
            {
                dynamic borders = _cell?.Borders;
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

                switch (style.LineStyle)
                {
                    case BorderLineStyle.Dash:
                        border.DashStyle = WPSHelper.LineStyle.Dash;
                        break;
                    case BorderLineStyle.Dot:
                        border.DashStyle = WPSHelper.LineStyle.RoundDot;
                        break;
                    case BorderLineStyle.DashDot:
                        border.DashStyle = WPSHelper.LineStyle.DashDot;
                        break;
                    case BorderLineStyle.DashDotDot:
                        border.DashStyle = WPSHelper.LineStyle.DashDotDot;
                        break;
                    default:
                        border.DashStyle = WPSHelper.LineStyle.Solid;
                        break;
                }
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

        public void SetBackgroundVisible(bool visible)
        {
            try
            {
                dynamic fill = _cell?.Shape?.Fill;
                if (fill == null) return;

                fill.Visible = visible ? WPSHelper.TriState.True : WPSHelper.TriState.False;
            }
            catch { }
        }

        public void SetFont(FontStyle style)
        {
            if (style == null) return;

            try
            {
                dynamic textRange = _cell?.Shape?.TextFrame?.TextRange;
                if (textRange == null) return;

                dynamic font = textRange.Font;
                if (font == null) return;

                if (!string.IsNullOrEmpty(style.Name))
                    font.Name = style.Name;

                if (!string.IsNullOrEmpty(style.NameFarEast))
                    font.NameFarEast = style.NameFarEast;

                if (style.Size > 0)
                    font.Size = style.Size;

                font.Bold = style.Bold ? WPSHelper.TriState.True : WPSHelper.TriState.False;
                font.Italic = style.Italic ? WPSHelper.TriState.True : WPSHelper.TriState.False;

                if (style.ColorRgb.HasValue)
                {
                    font.Color.RGB = style.ColorRgb.Value;
                }
                else if (style.ThemeColorIndex.HasValue)
                {
                    font.Color.ObjectThemeColor = style.ThemeColorIndex.Value;
                }
            }
            catch { }
        }

        public void SetAlignment(TextAlignment alignment)
        {
            try
            {
                dynamic textRange = _cell?.Shape?.TextFrame?.TextRange;
                if (textRange == null) return;

                int ppAlign = alignment switch
                {
                    TextAlignment.Left => 1,    // ppAlignLeft
                    TextAlignment.Center => 2,  // ppAlignCenter
                    TextAlignment.Right => 3,   // ppAlignRight
                    TextAlignment.Justify => 4, // ppAlignJustify
                    _ => 1
                };

                textRange.ParagraphFormat.Alignment = ppAlign;
            }
            catch { }
        }

        public object NativeCell => _cell;

        /// <summary>
        /// 获取 WPS Cell 动态对象
        /// </summary>
        public dynamic Cell => _cell;
    }
}
