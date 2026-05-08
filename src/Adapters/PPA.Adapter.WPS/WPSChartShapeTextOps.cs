using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
	/// <summary>
	/// WPS 图表字体设置（dynamic COM）。
	/// </summary>
	public class WPSChartShapeTextOps : IChartShapeTextOperations
	{
		public void ApplyChartFonts(object nativeShape, FontStyle primaryFont, FontStyle legendFont)
		{
			dynamic native = nativeShape;
			if (native?.Chart == null) return;

			dynamic chart = native.Chart;

			if (chart.HasTitle && chart.ChartTitle != null)
				ApplyFontToWps(chart.ChartTitle.Font, primaryFont);

			if (chart.HasLegend && chart.Legend != null)
				ApplyFontToWps(chart.Legend.Font, legendFont);

			try
			{
				if (chart.Axes != null)
				{
					foreach (dynamic axis in chart.Axes)
					{
						if (axis?.TickLabels?.Font != null)
							ApplyFontToWps(axis.TickLabels.Font, primaryFont);
					}
				}
			}
			catch { }

			try
			{
				if (chart.SeriesCollection != null)
				{
					foreach (dynamic series in chart.SeriesCollection)
					{
						if (series?.DataLabels != null)
						{
							var labels = series.DataLabels;
							if (labels.Font != null)
								ApplyFontToWps(labels.Font, primaryFont);
						}
					}
				}
			}
			catch { }
		}

		private static void ApplyFontToWps(dynamic font, FontStyle fontStyle)
		{
			if (font == null || fontStyle == null) return;

			try
			{
				if (!string.IsNullOrEmpty(fontStyle.Name))
					font.Name = fontStyle.Name;
				if (!string.IsNullOrEmpty(fontStyle.NameFarEast))
					font.NameFarEast = fontStyle.NameFarEast;
				if (fontStyle.Size > 0)
					font.Size = fontStyle.Size;
				font.Bold = fontStyle.Bold ? -1 : 0;
				if (fontStyle.ColorRgb.HasValue && font.Fill?.ForeColor != null)
					font.Fill.ForeColor.RGB = fontStyle.ColorRgb.Value;
				if (fontStyle.ThemeColorIndex.HasValue && font.Fill?.ForeColor != null)
					font.Fill.ForeColor.ObjectThemeColor = fontStyle.ThemeColorIndex.Value;
			}
			catch { }
		}
	}
}
