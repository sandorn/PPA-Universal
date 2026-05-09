using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 形状创建服务实现
	/// </summary>
	public class ShapeCreationService : IShapeCreationService
	{
		private readonly ILogger _logger;
		private readonly IShapeOperations _shapeOps;
		private readonly IApplicationContext _context;
		private readonly PPAConfig _config;

		public ShapeCreationService(ILogger logger, IShapeOperations shapeOps, IApplicationContext context, PPAConfig config)
		{
			_logger = logger ?? NullLogger.Instance;
			_shapeOps = shapeOps;
			_context = context;
			_config = config;
		}

		public IEnumerable<IShapeContext> CreateRectanglesAtShapes(IEnumerable<IShapeContext> shapes)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count == 0)
			{
				_logger.LogWarning("没有选中任何形状");
				return Enumerable.Empty<IShapeContext>();
			}

			_logger.LogInformation($"在 {shapeList.Count} 个形状位置创建矩形");

			var createdShapes = new List<IShapeContext>();

			foreach (var shape in shapeList)
			{
				if (shape?.NativeShape == null) continue;

				// 获取形状所在的幻灯片
				var slide = GetShapeSlide(shape);
				if (slide == null) continue;

				// 获取形状的边界
				var bounds = shape.Bounds;

				// 创建矩形
				var rectangle = _shapeOps.CreateRectangle(slide, bounds);
				if (rectangle != null)
				{
					// 创建 IShapeContext 包装
					// 注意：这里需要根据平台创建对应的 ShapeContext
					// 暂时返回 null，需要在 Adapter 层实现
					_logger.LogInformation($"在位置 ({bounds.Left}, {bounds.Top}) 创建矩形");
				}
			}

			_logger.LogInformation($"创建了 {createdShapes.Count} 个矩形");
			return createdShapes;
		}

		public IEnumerable<IShapeContext> CreateRectanglesOnSlides(IEnumerable<ISlideContext> slides, ShapeRect? bounds = null)
		{
			var slideList = slides?.ToList();
			if (slideList == null || slideList.Count == 0)
			{
				_logger.LogWarning("没有选中任何幻灯片");
				return Enumerable.Empty<IShapeContext>();
			}

			var d = _config?.Defaults;
			var slideW = _context?.ActivePresentation?.SlideWidth
				?? (d != null && d.SlideWidthFallback > 0 ? d.SlideWidthFallback : PpaConfigTemplateFallbacks.SlideWidthFallback);
			var slideH = _context?.ActivePresentation?.SlideHeight
				?? (d != null && d.SlideHeightFallback > 0 ? d.SlideHeightFallback : PpaConfigTemplateFallbacks.SlideHeightFallback);
			const float rw = 200f;
			const float rh = 100f;
			var defaultBounds = bounds ?? new ShapeRect
			{
				Left = (slideW - rw) / 2f,
				Top = (slideH - rh) / 2f,
				Width = rw,
				Height = rh
			};

			_logger.LogInformation($"在 {slideList.Count} 个幻灯片上创建矩形");

			var createdShapes = new List<IShapeContext>();

			foreach (var slide in slideList)
			{
				if (slide == null) continue;

				var rectangle = _shapeOps.CreateRectangle(slide, defaultBounds);
				if (rectangle != null)
				{
					_logger.LogInformation($"在幻灯片 {slide.SlideIndex} 上创建矩形");
				}
			}

			_logger.LogInformation($"创建了 {createdShapes.Count} 个矩形");
			return createdShapes;
		}

		private ISlideContext GetShapeSlide(IShapeContext shape)
		{
			// 通过 ApplicationContext 获取当前活动的幻灯片
			// 或者通过 shape.NativeShape.Parent 获取
			try
			{
				var context = _context;
				if (context?.ActiveWindow?.ActiveSlide != null)
				{
					return context.ActiveWindow.ActiveSlide;
				}

				// 尝试从原生形状获取父幻灯片
				if (shape?.NativeShape != null)
				{
					dynamic nativeShape = shape.NativeShape;
					dynamic parent = nativeShape?.Parent;
					if (parent != null)
					{
						// 这里需要根据平台创建对应的 SlideContext
						// 暂时返回 null
					}
				}
			}
			catch
			{
				// 忽略错误
			}

			return null;
		}
	}
}

