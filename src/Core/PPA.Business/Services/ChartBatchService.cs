using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 图表批量操作服务实现
	/// </summary>
	public class ChartBatchService : IChartBatchService
	{
		private readonly ILogger _logger;
		private readonly PPAConfig _config;
		private readonly IChartShapeTextOperations _chartTextOps;

		public ChartBatchService(ILogger logger, PPAConfig config, IChartShapeTextOperations chartTextOps)
		{
			_logger = logger ?? NullLogger.Instance;
			_config = config;
			_chartTextOps = chartTextOps;
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

			var primaryFont = fontStyle ?? BuildDefaultChartTitleFontStyle();
			var legendFontStyle = BuildDefaultChartLegendFontStyle();

			_logger.LogInformation($"格式化 {shapeList.Count} 个图表的字体");

			foreach (var shape in shapeList)
			{
				if (shape?.IsChart != true || shape.NativeShape == null) continue;

				try
				{
					_chartTextOps.ApplyChartFonts(shape.NativeShape, primaryFont, legendFontStyle);
					_logger.LogInformation($"格式化图表字体: {shape.Name}");
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"格式化图表字体失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation("图表字体格式化完成");
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
	}
}
