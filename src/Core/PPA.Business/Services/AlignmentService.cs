using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Business.Geometry;
using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 对齐服务实现
	/// </summary>
	public class AlignmentService : IAlignmentService
	{
		private readonly ILogger _logger;
		private readonly IShapeOperations _shapeOps;
		private readonly IApplicationContext _context;
		private readonly PPAConfig _config;

		public AlignmentService(ILogger logger, IShapeOperations shapeOps, IApplicationContext context, PPAConfig config)
		{
			_logger = logger ?? NullLogger.Instance;
			_shapeOps = shapeOps;
			_context = context;
			_config = config;
		}

		public void Align(IEnumerable<IShapeContext> shapes, AlignmentType alignment, AlignmentReference reference)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count == 0)
			{
				_logger.LogWarning("没有选中任何形状");
				return;
			}

			_logger.LogInformation($"执行对齐: {alignment}, 参考: {reference}, 形状数: {shapeList.Count}");

			// 计算参考边界
			float refValue = CalculateReferenceValue(shapeList, alignment, reference);

			// 对每个形状应用对齐
			foreach (var shape in shapeList)
			{
				var bounds = shape.Bounds;
				var newBounds = ShapeAlignmentMath.ApplyAlignment(bounds, alignment, refValue);
				shape.Bounds = newBounds;
			}

			_logger.LogInformation("对齐完成");
		}

		private float CalculateReferenceValue(List<IShapeContext> shapes, AlignmentType alignment, AlignmentReference reference)
		{
			switch (reference)
			{
				case AlignmentReference.Slide:
					return CalculateSlideReference(alignment, shapes);
				case AlignmentReference.FirstObject:
					return CalculateSingleShapeReference(shapes.FirstOrDefault(), alignment);
				case AlignmentReference.LastObject:
					return CalculateSingleShapeReference(shapes.LastOrDefault(), alignment);
				case AlignmentReference.SelectedObjects:
				default:
					return CalculateSelectionReference(shapes, alignment);
			}
		}

		private float CalculateSlideReference(AlignmentType alignment, List<IShapeContext> shapes)
		{
			var d = _config?.Defaults;
			float slideWidth = _context?.ActivePresentation?.SlideWidth
				?? (d != null && d.SlideWidthFallback > 0 ? d.SlideWidthFallback : PpaConfigTemplateFallbacks.SlideWidthFallback);
			float slideHeight = _context?.ActivePresentation?.SlideHeight
				?? (d != null && d.SlideHeightFallback > 0 ? d.SlideHeightFallback : PpaConfigTemplateFallbacks.SlideHeightFallback);
			var rects = shapes.Select(s => s.Bounds).ToList();
			return ShapeAlignmentMath.CalculateSlideReference(alignment, slideWidth, slideHeight, rects);
		}

		private static float CalculateSingleShapeReference(IShapeContext shape, AlignmentType alignment)
		{
			if (shape == null)
				return 0f;
			return ShapeAlignmentMath.CalculateSingleShapeReference(shape.Bounds, alignment);
		}

		private static float CalculateSelectionReference(List<IShapeContext> shapes, AlignmentType alignment)
		{
			var rects = shapes.Select(s => s.Bounds).ToList();
			return ShapeAlignmentMath.CalculateSelectionReference(rects, alignment);
		}

		public void Distribute(IEnumerable<IShapeContext> shapes, DistributionType distribution)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count < 3)
			{
				_logger.LogWarning("分布操作至少需要3个形状");
				return;
			}

			_logger.LogInformation($"执行分布: {distribution}, 形状数: {shapeList.Count}");

			if (distribution == DistributionType.Horizontal)
			{
				DistributeHorizontally(shapeList);
			}
			else
			{
				DistributeVertically(shapeList);
			}

			_logger.LogInformation("分布完成");
		}

		private static void DistributeHorizontally(List<IShapeContext> shapes)
		{
			var sorted = shapes.OrderBy(s => s.Bounds.Left).ToList();
			var rects = sorted.Select(s => s.Bounds).ToList();
			var newLefts = ShapeAlignmentMath.ComputeHorizontalDistributedLefts(rects);
			for (var i = 0; i < sorted.Count; i++)
			{
				var bounds = sorted[i].Bounds;
				sorted[i].Bounds = new ShapeRect(newLefts[i], bounds.Top, bounds.Width, bounds.Height);
			}
		}

		private static void DistributeVertically(List<IShapeContext> shapes)
		{
			var sorted = shapes.OrderBy(s => s.Bounds.Top).ToList();
			var rects = sorted.Select(s => s.Bounds).ToList();
			var newTops = ShapeAlignmentMath.ComputeVerticalDistributedTops(rects);
			for (var i = 0; i < sorted.Count; i++)
			{
				var bounds = sorted[i].Bounds;
				sorted[i].Bounds = new ShapeRect(bounds.Left, newTops[i], bounds.Width, bounds.Height);
			}
		}

		public void SetEqualWidth(IEnumerable<IShapeContext> shapes)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count < 2) return;

			var maxWidth = shapeList.Max(s => s.Bounds.Width);
			foreach (var shape in shapeList)
			{
				var bounds = shape.Bounds;
				shape.Bounds = new ShapeRect(bounds.Left, bounds.Top, maxWidth, bounds.Height);
			}

			_logger.LogInformation($"已设置等宽: {maxWidth}");
		}

		public void SetEqualHeight(IEnumerable<IShapeContext> shapes)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count < 2) return;

			var maxHeight = shapeList.Max(s => s.Bounds.Height);
			foreach (var shape in shapeList)
			{
				var bounds = shape.Bounds;
				shape.Bounds = new ShapeRect(bounds.Left, bounds.Top, bounds.Width, maxHeight);
			}

			_logger.LogInformation($"已设置等高: {maxHeight}");
		}

		public void SetEqualSize(IEnumerable<IShapeContext> shapes)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count < 2) return;

			var maxWidth = shapeList.Max(s => s.Bounds.Width);
			var maxHeight = shapeList.Max(s => s.Bounds.Height);
			foreach (var shape in shapeList)
			{
				var bounds = shape.Bounds;
				shape.Bounds = new ShapeRect(bounds.Left, bounds.Top, maxWidth, maxHeight);
			}

			_logger.LogInformation($"已设置等大小: {maxWidth} x {maxHeight}");
		}

		public void SwapPositions(IShapeContext shape1, IShapeContext shape2)
		{
			if (shape1 == null || shape2 == null) return;

			var bounds1 = shape1.Bounds;
			var bounds2 = shape2.Bounds;
			var (na, nb) = ShapeAlignmentMath.SwapCenters(bounds1, bounds2);
			shape1.Bounds = na;
			shape2.Bounds = nb;

			_logger.LogInformation("已交换两个形状的位置");
		}

		public void SnapToShape(IEnumerable<IShapeContext> shapes, SnapDirection snapDirection)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count < 2)
			{
				_logger.LogWarning("吸附操作至少需要2个形状");
				return;
			}

			var referenceShape = shapeList.First();
			var otherShapes = shapeList.Skip(1).ToList();

			_logger.LogInformation($"执行吸附: {snapDirection}, 基准形状: 1, 其他形状数: {otherShapes.Count}");

			var refBounds = referenceShape.Bounds;

			// 以第一个形状为基准，其余形状整体贴靠在基准一侧（与基准对边共线）：
			// 左：他形左 = 基准左 − (他形右 − 他形左)，他形整体在基准左侧，他形右 = 基准左
			// 右：他形左 = 基准右，他形整体在基准右侧
			// 上：他形上 = 基准上 − (他形下 − 他形上)，他形整体在基准上方，他形下 = 基准上
			// 下：他形上 = 基准下，他形整体在基准下方
			foreach (var shape in otherShapes)
			{
				var b = shape.Bounds;
				shape.Bounds = ShapeAlignmentMath.SnapOtherToReference(b, refBounds, snapDirection);
			}

			_logger.LogInformation("吸附完成");
		}

		public void ExtendAlignment(IEnumerable<IShapeContext> shapes, ExtendDirection extendDirection)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count < 2)
			{
				_logger.LogWarning("延伸对齐操作至少需要2个形状");
				return;
			}

			_logger.LogInformation($"执行延伸对齐: {extendDirection}, 形状数: {shapeList.Count}");

			float targetValue = 0f;

			switch (extendDirection)
			{
				case ExtendDirection.Left:
					targetValue = shapeList.Min(s => s.Bounds.Left);
					break;
				case ExtendDirection.Right:
					targetValue = shapeList.Max(s => s.Bounds.Right);
					break;
				case ExtendDirection.Top:
					targetValue = shapeList.Min(s => s.Bounds.Top);
					break;
				case ExtendDirection.Bottom:
					targetValue = shapeList.Max(s => s.Bounds.Bottom);
					break;
			}

			foreach (var shape in shapeList)
			{
				var bounds = shape.Bounds;
				ShapeRect newBounds;

				switch (extendDirection)
				{
					case ExtendDirection.Left:
						newBounds = new ShapeRect(targetValue, bounds.Top, bounds.Right - targetValue, bounds.Height);
						break;
					case ExtendDirection.Right:
						newBounds = new ShapeRect(bounds.Left, bounds.Top, targetValue - bounds.Left, bounds.Height);
						break;
					case ExtendDirection.Top:
						newBounds = new ShapeRect(bounds.Left, targetValue, bounds.Width, bounds.Bottom - targetValue);
						break;
					case ExtendDirection.Bottom:
						newBounds = new ShapeRect(bounds.Left, bounds.Top, bounds.Width, targetValue - bounds.Top);
						break;
					default:
						newBounds = bounds;
						break;
				}

				shape.Bounds = newBounds;
			}

			_logger.LogInformation("延伸对齐完成");
		}

		public void SwapPositionsAndSize(IShapeContext shape1, IShapeContext shape2)
		{
			if (shape1 == null || shape2 == null) return;

			var bounds1 = shape1.Bounds;
			var bounds2 = shape2.Bounds;

			// 交换位置和大小
			shape1.Bounds = new ShapeRect(bounds2.Left, bounds2.Top, bounds2.Width, bounds2.Height);
			shape2.Bounds = new ShapeRect(bounds1.Left, bounds1.Top, bounds1.Width, bounds1.Height);

			_logger.LogInformation("已交换两个形状的位置和大小");
		}
	}
}
