using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 裁除服务实现
	/// </summary>
	public class CropService : ICropService
	{
		private readonly ILogger _logger;
		private readonly IApplicationContext _context;

		public CropService(ILogger logger, IApplicationContext context)
		{
			_logger = logger ?? NullLogger.Instance;
			_context = context;
		}

		public void CropShapesToSlide(IEnumerable<IShapeContext> shapes, ISlideContext slide)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count == 0)
			{
				_logger.LogWarning("没有选中任何形状");
				return;
			}

			if (slide == null)
			{
				_logger.LogWarning("幻灯片为空");
				return;
			}

			// 获取幻灯片尺寸
			var presentation = _context?.ActivePresentation;
			if (presentation == null)
			{
				_logger.LogWarning("无法获取演示文稿");
				return;
			}

			float slideWidth = presentation.SlideWidth;
			float slideHeight = presentation.SlideHeight;

			_logger.LogInformation($"裁除 {shapeList.Count} 个形状到幻灯片边界 ({slideWidth} x {slideHeight})");

			int croppedCount = 0;

			foreach (var shape in shapeList)
			{
				var bounds = shape.Bounds;
				bool needsCrop = false;
				float newLeft = bounds.Left;
				float newTop = bounds.Top;
				float newWidth = bounds.Width;
				float newHeight = bounds.Height;

				// 检查并调整左边界
				if (bounds.Left < 0)
				{
					newWidth += bounds.Left; // 减少宽度
					newLeft = 0;
					needsCrop = true;
				}

				// 检查并调整上边界
				if (bounds.Top < 0)
				{
					newHeight += bounds.Top; // 减少高度
					newTop = 0;
					needsCrop = true;
				}

				// 检查并调整右边界
				if (bounds.Right > slideWidth)
				{
					newWidth = slideWidth - newLeft;
					needsCrop = true;
				}

				// 检查并调整下边界
				if (bounds.Bottom > slideHeight)
				{
					newHeight = slideHeight - newTop;
					needsCrop = true;
				}

				// 确保宽度和高度为正数
				if (newWidth < 0) newWidth = 0;
				if (newHeight < 0) newHeight = 0;

				if (needsCrop)
				{
					shape.Bounds = new ShapeRect(newLeft, newTop, newWidth, newHeight);
					croppedCount++;
				}
			}

			_logger.LogInformation($"已裁除 {croppedCount} 个形状");
		}

		public void CropAllShapesToSlide(ISlideContext slide)
		{
			if (slide == null)
			{
				_logger.LogWarning("幻灯片为空");
				return;
			}

			var shapes = slide.Shapes?.ToList();
			if (shapes == null || shapes.Count == 0)
			{
				_logger.LogWarning("幻灯片中没有形状");
				return;
			}

			CropShapesToSlide(shapes, slide);
		}
	}
}

