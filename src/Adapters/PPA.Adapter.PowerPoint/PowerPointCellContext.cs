using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 单元格上下文实现
    /// </summary>
    public class PowerPointCellContext : ICellContext
    {
        private readonly NETOP.Cell _cell;
        private readonly int _row;
        private readonly int _column;

        public PowerPointCellContext(NETOP.Cell cell, int row, int column)
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
                    var textRange = _cell?.Shape?.TextFrame?.TextRange;
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
                try
                {
                    // NetOffice Cell 没有直接的 IsMerged 属性，需要通过其他方式判断
                    // 暂时返回 false，后续可根据需要完善
                    return false;
                }
                catch
                {
                    return false;
                }
            }
        }

        public void SetBackground(int colorRgb)
        {
            try
            {
                var fill = _cell?.Shape?.Fill;
                if (fill == null) return;

                fill.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
                fill.Solid();
                fill.ForeColor.RGB = colorRgb;
            }
            catch { }
        }

        public void SetBackgroundVisible(bool visible)
        {
            try
            {
                var fill = _cell?.Shape?.Fill;
                if (fill == null) return;

                fill.Visible = visible
                    ? NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                    : NetOffice.OfficeApi.Enums.MsoTriState.msoFalse;
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
                var borders = _cell?.Borders;
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

            switch (style.LineStyle)
            {
                case BorderLineStyle.Dash:
                    border.DashStyle = NetOffice.OfficeApi.Enums.MsoLineDashStyle.msoLineDash;
                    break;
                case BorderLineStyle.Dot:
                    border.DashStyle = NetOffice.OfficeApi.Enums.MsoLineDashStyle.msoLineRoundDot;
                    break;
                case BorderLineStyle.DashDot:
                    border.DashStyle = NetOffice.OfficeApi.Enums.MsoLineDashStyle.msoLineDashDot;
                    break;
                case BorderLineStyle.DashDotDot:
                    border.DashStyle = NetOffice.OfficeApi.Enums.MsoLineDashStyle.msoLineDashDotDot;
                    break;
                default:
                    border.DashStyle = NetOffice.OfficeApi.Enums.MsoLineDashStyle.msoLineSolid;
                    break;
            }
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

        public void SetFont(FontStyle style)
        {
            if (style == null) return;

            try
            {
                var textRange = _cell?.Shape?.TextFrame?.TextRange;
                if (textRange == null) return;

                var font = textRange.Font;
                if (font == null) return;

                if (!string.IsNullOrEmpty(style.Name))
                    font.Name = style.Name;

                if (!string.IsNullOrEmpty(style.NameFarEast))
                    font.NameFarEast = style.NameFarEast;

                if (style.Size > 0)
                    font.Size = style.Size;

                font.Bold = style.Bold
                    ? NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                    : NetOffice.OfficeApi.Enums.MsoTriState.msoFalse;

                font.Italic = style.Italic
                    ? NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                    : NetOffice.OfficeApi.Enums.MsoTriState.msoFalse;

                if (style.ColorRgb.HasValue)
                {
                    font.Color.RGB = style.ColorRgb.Value;
                }
                else if (style.ThemeColorIndex.HasValue)
                {
                    font.Color.ObjectThemeColor = (NetOffice.OfficeApi.Enums.MsoThemeColorIndex)style.ThemeColorIndex.Value;
                }
            }
            catch { }
        }

        public void SetAlignment(TextAlignment alignment)
        {
            try
            {
                var textRange = _cell?.Shape?.TextFrame?.TextRange;
                if (textRange == null) return;

                var ppAlign = alignment switch
                {
                    TextAlignment.Left => NETOP.Enums.PpParagraphAlignment.ppAlignLeft,
                    TextAlignment.Center => NETOP.Enums.PpParagraphAlignment.ppAlignCenter,
                    TextAlignment.Right => NETOP.Enums.PpParagraphAlignment.ppAlignRight,
                    TextAlignment.Justify => NETOP.Enums.PpParagraphAlignment.ppAlignJustify,
                    _ => NETOP.Enums.PpParagraphAlignment.ppAlignLeft
                };

                textRange.ParagraphFormat.Alignment = ppAlign;
            }
            catch { }
        }

        public object NativeCell => _cell;

        /// <summary>
        /// 获取原生 Cell 对象
        /// </summary>
        public NETOP.Cell NetCell => _cell;
    }
}
