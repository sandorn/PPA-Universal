using System;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using PPA.Logging;

namespace PPA.Adapter.WPS
{
	/// <summary>
	/// WPS 平台的毛玻璃卡片渲染器（退化版实现）
	/// </summary>
	public class WPSGlassCardRenderer : IGlassCardRenderer
	{
		private readonly ILogger _logger;
		private readonly IIdMsoCommandExecutor _idMsoExecutor;

		public WPSGlassCardRenderer(ILogger logger = null, IIdMsoCommandExecutor idMsoExecutor = null)
		{
			_logger = logger ?? NullLogger.Instance;
			_idMsoExecutor = idMsoExecutor;
		}

		public void RenderGlassCard(IApplicationContext app, ShapeRect rect, GlassCardConfig config)
		{
			try
			{
				// 1. 获取原生幻灯片对象 (dynamic)
				var slideContext = app?.ActiveWindow?.ActiveSlide;
				if (slideContext == null)
				{
					_logger.LogWarning("无法获取当前幻灯片上下文，跳过渲染");
					return;
				}

				dynamic slide = slideContext.NativeSlide;
				if (slide == null)
				{
					_logger.LogWarning("NativeSlide 为空");
					return;
				}

				_logger.LogInformation("开始在 WPS 上创建形状");

				// 2. 创建圆角矩形 (msoShapeRoundedRectangle = 5)
				// WPS dynamic 调用 AddShape
				dynamic shapes = slide.Shapes;
				dynamic shape = shapes.AddShape(5, rect.Left, rect.Top, rect.Width, rect.Height);

				// 设置名称
				try { shape.Name = "PPA_GlassCard_" + DateTime.Now.Ticks; } catch { }

				// 3. 初步设置样式（半透明白色填充 + 边框）
				// 边框：为避免出现黑色/深色边框，这里在 WPS 上统一使用白色细边
				if (config.BorderWidth > 0)
				{
					shape.Line.Visible = -1; // msoTrue
					shape.Line.Weight = config.BorderWidth;
					shape.Line.ForeColor.RGB = 16777215; // White
				}
				else
				{
					shape.Line.Visible = 0; // msoFalse
				}

				// 填充：WPS 对渐变停靠点支持不稳定，这里简化为白色半透明填充，保证有明显的透明感
				try
				{
					dynamic fill = shape.Fill;
					fill.Visible = -1; // msoTrue
					try { fill.Solid(); } catch { }
					fill.ForeColor.RGB = 16777215; // White
												   // 使用略高透明度，让背景内容更明显透出
					fill.Transparency = 0.35f;
				}
				catch
				{
					// 渐变应用失败时，退回到简单半透明填充
					shape.Fill.Visible = -1; // msoTrue
					try { shape.Fill.Solid(); } catch { }
					shape.Fill.ForeColor.RGB = 16777215; // White
					shape.Fill.Transparency = 0.4f; // 稍明显的透明度
				}

				// 阴影：右下角柔和阴影（增强强度，确保在 WPS 中可见）
				try
				{
					dynamic shadow = shape.Shadow;
					shadow.Style = 1; // msoShadowStyleOuterShadow
					shadow.Type = 1; // msoShadow1
					shadow.Visible = -1;  // True
					shadow.ForeColor.RGB = 0x000000;
					shadow.OffsetX = 10f;
					shadow.OffsetY = 10f;
					shadow.Transparency = 0.8f;
					var blurRadius = config.BlurRadius > 0 ? config.BlurRadius : 10f;
					shadow.Blur = blurRadius; // 模糊 10 磅
					shadow.Size = 100f;

					// 验证设置是否生效
					Console.WriteLine($"阴影类型: {shadow.Type}");
					Console.WriteLine($"透明度: {shadow.Transparency}");
					Console.WriteLine($"X偏移: {shadow.OffsetX}");
					Console.WriteLine($"Y偏移: {shadow.OffsetY}");
					Console.WriteLine($"模糊半径: {shadow.Blur}");
				}
				catch { }

				_logger.LogInformation("成功创建 WPS 卡片形状");
			}
			catch (Exception ex)
			{
				_logger.LogError($"WPS 渲染毛玻璃卡片失败: {ex.Message}", ex);
			}
		}
	}
}
