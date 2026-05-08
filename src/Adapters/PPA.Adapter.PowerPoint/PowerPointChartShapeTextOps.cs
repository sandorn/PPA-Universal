using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;
using OFFICE = NetOffice.OfficeApi.Enums;

namespace PPA.Adapter.PowerPoint
{
	/// <summary>
	/// PowerPoint 图表上字体设置（NetOffice + dynamic 兼容多类 Chart 字体对象）。
	/// </summary>
	public class PowerPointChartShapeTextOps : IChartShapeTextOperations
	{
		public void ApplyChartFonts(object nativeShape, FontStyle primaryFont, FontStyle legendFont)
		{
			var netShape = nativeShape as NETOP.Shape;
			if (netShape?.Chart == null) return;

			var chart = netShape.Chart;

			if (chart.HasTitle && chart.ChartTitle != null)
				ApplyFontToPowerPoint(chart.ChartTitle.Font, primaryFont);

			if (chart.HasLegend && chart.Legend != null)
				ApplyFontToPowerPoint(chart.Legend.Font, legendFont);

			try
			{
				var axes = chart.Axes();
				if (axes != null)
				{
					dynamic axesCollection = axes;
					if (axesCollection != null)
					{
						int count = axesCollection.Count ?? 0;
						for (int i = 1; i <= count; i++)
						{
							try
							{
								dynamic axis = axesCollection[i];
								if (axis?.TickLabels?.Font != null)
									ApplyFontToPowerPoint(axis.TickLabels.Font, primaryFont);
							}
							catch { /* 单个轴 */ }
						}
					}
				}
			}
			catch { /* 无坐标轴 */ }

			try
			{
				var seriesCollection = chart.SeriesCollection();
				if (seriesCollection != null)
				{
					dynamic seriesColl = seriesCollection;
					if (seriesColl != null)
					{
						int count = seriesColl.Count ?? 0;
						for (int i = 1; i <= count; i++)
						{
							try
							{
								dynamic series = seriesColl[i];
								if (series?.DataLabels != null)
								{
									var labels = series.DataLabels;
									if (labels.Font != null)
										ApplyFontToPowerPoint(labels.Font, primaryFont);
								}
							}
							catch { /* 单个系列 */ }
						}
					}
				}
			}
			catch { /* 无数据标签 */ }
		}

		private static void ApplyFontToPowerPoint(object font, FontStyle fontStyle)
		{
			if (font == null || fontStyle == null) return;

			try
			{
				dynamic f = font;

				if (!string.IsNullOrEmpty(fontStyle.Name))
				{
					try { f.Name = fontStyle.Name; } catch { }
				}
				if (!string.IsNullOrEmpty(fontStyle.NameFarEast))
				{
					try { f.NameFarEast = fontStyle.NameFarEast; } catch { }
				}
				if (fontStyle.Size > 0)
				{
					try { f.Size = fontStyle.Size; } catch { }
				}

				try
				{
					f.Bold = fontStyle.Bold ? OFFICE.MsoTriState.msoTrue : OFFICE.MsoTriState.msoFalse;
				}
				catch
				{
					try { f.Bold = fontStyle.Bold; } catch { }
				}

				if (fontStyle.ColorRgb.HasValue)
				{
					try { f.Fill.ForeColor.RGB = fontStyle.ColorRgb.Value; }
					catch { try { f.Color.RGB = fontStyle.ColorRgb.Value; } catch { } }
				}
				if (fontStyle.ThemeColorIndex.HasValue)
				{
					try
					{
						f.Fill.ForeColor.ObjectThemeColor =
							(OFFICE.MsoThemeColorIndex)fontStyle.ThemeColorIndex.Value;
					}
					catch { try { f.Color.ObjectThemeColor = fontStyle.ThemeColorIndex.Value; } catch { } }
				}
			}
			catch { /* 忽略单项 */ }
		}
	}
}
