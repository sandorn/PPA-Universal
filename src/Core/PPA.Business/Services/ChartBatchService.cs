using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using PPA.Logging;
using NetOffice.PowerPointApi;

namespace PPA.Business.Services
{
	/// <summary>
	/// 图表批量操作服务实现
	/// </summary>
	public class ChartBatchService : IChartBatchService
	{
		private readonly ILogger _logger;
		private readonly PPAConfig _config;

		public ChartBatchService(ILogger logger, PPAConfig config)
		{
			_logger = logger ?? NullLogger.Instance;
			_config = config;
		}

		public void FormatSelectedCharts(IApplicationContext context)
		{
			if (context?.Selection == null)
			{
				_logger.LogWarning("无法获取选择对象");
				return;
			}

			var shapes = context.Selection.SelectedShapes?.ToList();
			if (shapes == null || shapes.Count == 0)
			{
				_logger.LogWarning("没有选中任何形状");
				return;
			}

			var chartShapes = shapes.Where(s => s?.IsChart == true).ToList();
			if (chartShapes.Count == 0)
			{
				_logger.LogWarning("选中的形状中没有图表");
				return;
			}

			_logger.LogInformation($"格式化 {chartShapes.Count} 个图表");

			// 使用默认字体样式
			FormatChartFont(chartShapes, null);
		}

		public void FormatCurrentSlideCharts(IApplicationContext context)
		{
			if (context?.ActiveWindow?.ActiveSlide == null)
			{
				_logger.LogWarning("无法获取当前幻灯片");
				return;
			}

			var slide = context.ActiveWindow.ActiveSlide;
			var chartShapes = slide.Shapes?.Where(s => s?.IsChart == true).ToList();

			if (chartShapes == null || chartShapes.Count == 0)
			{
				_logger.LogWarning("当前幻灯片中没有图表");
				return;
			}

			_logger.LogInformation($"格式化当前幻灯片 {chartShapes.Count} 个图表");

			// 使用默认字体样式
			FormatChartFont(chartShapes, null);
		}

		public void FormatChartFont(IEnumerable<IShapeContext> shapes, FontStyle fontStyle = null)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count == 0)
			{
				_logger.LogWarning("没有选中任何形状");
				return;
			}

			if (fontStyle == null)
			{
				fontStyle = BuildDefaultChartTitleFontStyle();
			}

			var legendFontStyle = BuildDefaultChartLegendFontStyle();

			_logger.LogInformation($"格式化 {shapeList.Count} 个图表的字体");

			foreach (var shape in shapeList)
			{
				if (shape?.IsChart != true) continue;

				try
				{
					var platform = GetPlatform(shape);
					if (platform == PlatformType.PowerPoint)
					{
						FormatChartFontPowerPoint(shape, fontStyle, legendFontStyle);
					}
					else if (platform == PlatformType.WPS)
					{
						FormatChartFontWPS(shape, fontStyle, legendFontStyle);
					}

					_logger.LogInformation($"格式化图表字体: {shape.Name}");
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"格式化图表字体失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation("图表字体格式化完成");
		}

		private void FormatChartFontPowerPoint(IShapeContext shape, FontStyle titleFontStyle, FontStyle legendFontStyle)
		{
			try
			{
				var netShape = shape.NativeShape as Shape;
				if (netShape?.Chart == null) return;

				var chart = netShape.Chart;

				// 格式化图表标题
				if (chart.HasTitle && chart.ChartTitle != null)
				{
					var titleFont = chart.ChartTitle.Font;
					ApplyFontToPowerPoint(titleFont, titleFontStyle);
				}

				// 格式化图例
				if (chart.HasLegend && chart.Legend != null)
				{
					var legendFont = chart.Legend.Font;
					ApplyFontToPowerPoint(legendFont, legendFontStyle);
				}

				// 格式化坐标轴
				try
				{
					var axes = chart.Axes();
					if (axes != null)
					{
						// 使用 dynamic 访问集合
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
									{
										ApplyFontToPowerPoint(axis.TickLabels.Font, titleFontStyle);
									}
								}
								catch
								{
									// 忽略单个轴的错误
								}
							}
						}
					}
				}
				catch
				{
					// 某些图表类型可能不支持坐标轴，忽略错误
				}

				// 格式化数据标签
				try
				{
					var seriesCollection = chart.SeriesCollection();
					if (seriesCollection != null)
					{
						// 使用 dynamic 访问集合
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
										{
											ApplyFontToPowerPoint(labels.Font, titleFontStyle);
										}
									}
								}
								catch
								{
									// 忽略单个系列的错误
								}
							}
						}
					}
				}
				catch
				{
					// 某些图表类型可能不支持数据标签，忽略错误
				}
			}
			catch (System.Exception ex)
			{
				_logger.LogError($"PowerPoint 图表字体格式化失败: {ex.Message}", ex);
				throw;
			}
		}

		private void FormatChartFontWPS(IShapeContext shape, FontStyle titleFontStyle, FontStyle legendFontStyle)
		{
			try
			{
				dynamic nativeShape = shape.NativeShape;
				if (nativeShape?.Chart == null) return;

				dynamic chart = nativeShape.Chart;

				// 格式化图表标题
				if (chart.HasTitle && chart.ChartTitle != null)
				{
					var titleFont = chart.ChartTitle.Font;
					ApplyFontToWPS(titleFont, titleFontStyle);
				}

				// 格式化图例
				if (chart.HasLegend && chart.Legend != null)
				{
					var legendFont = chart.Legend.Font;
					ApplyFontToWPS(legendFont, legendFontStyle);
				}

				// 格式化坐标轴
				try
				{
					if (chart.Axes != null)
					{
						foreach (dynamic axis in chart.Axes)
						{
							if (axis?.TickLabels?.Font != null)
							{
								ApplyFontToWPS(axis.TickLabels.Font, titleFontStyle);
							}
						}
					}
				}
				catch
				{
					// 某些图表类型可能不支持坐标轴，忽略错误
				}

				// 格式化数据标签
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
								{
									ApplyFontToWPS(labels.Font, titleFontStyle);
								}
							}
						}
					}
				}
				catch
				{
					// 某些图表类型可能不支持数据标签，忽略错误
				}
			}
			catch (System.Exception ex)
			{
				_logger.LogError($"WPS 图表字体格式化失败: {ex.Message}", ex);
				throw;
			}
		}

		private void ApplyFontToPowerPoint(object font, FontStyle fontStyle)
		{
			try
			{
				if (font == null) return;

				// 图表字体对象在不同图表元素上类型不一致（并不总是 Font2）。
				// 用 dynamic 兼容 ChartTitle/Legend/AxisTickLabels/DataLabels 等场景。
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
					f.Bold = fontStyle.Bold
						? NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
						: NetOffice.OfficeApi.Enums.MsoTriState.msoFalse;
				}
				catch
				{
					try { f.Bold = fontStyle.Bold; } catch { }
				}

				if (fontStyle.ColorRgb.HasValue)
				{
					try
					{
						f.Fill.ForeColor.RGB = fontStyle.ColorRgb.Value;
					}
					catch
					{
						try { f.Color.RGB = fontStyle.ColorRgb.Value; } catch { }
					}
				}
				if (fontStyle.ThemeColorIndex.HasValue)
				{
					try
					{
						f.Fill.ForeColor.ObjectThemeColor =
							(NetOffice.OfficeApi.Enums.MsoThemeColorIndex)fontStyle.ThemeColorIndex.Value;
					}
					catch
					{
						try { f.Color.ObjectThemeColor = fontStyle.ThemeColorIndex.Value; } catch { }
					}
				}
			}
			catch
			{
				// 忽略字体设置错误
			}
		}

		private FontStyle BuildDefaultChartTitleFontStyle()
		{
			var cfg = _config?.Chart?.TitleFont;
			return new FontStyle
			{
				Name = string.IsNullOrWhiteSpace(cfg?.Name) ? "+mn-lt" : cfg.Name,
				NameFarEast = string.IsNullOrWhiteSpace(cfg?.NameFarEast) ? "+mn-ea" : cfg.NameFarEast,
				Size = (cfg?.Size ?? 0) > 0 ? cfg.Size : 11,
				Bold = cfg?.Bold ?? false,
				ThemeColorIndex = cfg?.ThemeColorIndex
			};
		}

		private FontStyle BuildDefaultChartLegendFontStyle()
		{
			var cfg = _config?.Chart?.LegendFont;
			return new FontStyle
			{
				Name = string.IsNullOrWhiteSpace(cfg?.Name) ? "+mn-lt" : cfg.Name,
				NameFarEast = string.IsNullOrWhiteSpace(cfg?.NameFarEast) ? "+mn-ea" : cfg.NameFarEast,
				Size = (cfg?.Size ?? 0) > 0 ? cfg.Size : 8,
				Bold = cfg?.Bold ?? false,
				ThemeColorIndex = cfg?.ThemeColorIndex
			};
		}

		private void ApplyFontToWPS(dynamic font, FontStyle fontStyle)
		{
			try
			{
				if (font == null) return;

				if (!string.IsNullOrEmpty(fontStyle.Name))
				{
					font.Name = fontStyle.Name;
				}
				if (!string.IsNullOrEmpty(fontStyle.NameFarEast))
				{
					font.NameFarEast = fontStyle.NameFarEast;
				}
				if (fontStyle.Size > 0)
				{
					font.Size = fontStyle.Size;
				}
				font.Bold = fontStyle.Bold ? -1 : 0;
				if (fontStyle.ColorRgb.HasValue && font.Fill?.ForeColor != null)
				{
					font.Fill.ForeColor.RGB = fontStyle.ColorRgb.Value;
				}
				if (fontStyle.ThemeColorIndex.HasValue && font.Fill?.ForeColor != null)
				{
					font.Fill.ForeColor.ObjectThemeColor = fontStyle.ThemeColorIndex.Value;
				}
			}
			catch
			{
				// 忽略字体设置错误
			}
		}

		private PlatformType GetPlatform(IShapeContext shape)
		{
			if (shape?.NativeShape == null) return PlatformType.Unknown;

			var typeName = shape.NativeShape.GetType().FullName;
			if (typeName?.Contains("NetOffice.PowerPointApi") == true)
			{
				return PlatformType.PowerPoint;
			}
			else if (typeName?.Contains("WPS") == true || typeName?.Contains("Kingsoft") == true)
			{
				return PlatformType.WPS;
			}

			return PlatformType.Unknown;
		}
	}
}

