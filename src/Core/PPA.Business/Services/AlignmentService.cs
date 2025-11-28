using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
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

        public AlignmentService(ILogger logger, IShapeOperations shapeOps, IApplicationContext context)
        {
            _logger = logger ?? NullLogger.Instance;
            _shapeOps = shapeOps;
            _context = context;
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
                var newBounds = ApplyAlignment(bounds, alignment, refValue);
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
            // 获取幻灯片尺寸
            float slideWidth = _context?.ActivePresentation?.SlideWidth ?? 960f;
            float slideHeight = _context?.ActivePresentation?.SlideHeight ?? 540f;

            switch (alignment)
            {
                case AlignmentType.Left:
                    return 0f;
                case AlignmentType.Top:
                    return 0f;
                case AlignmentType.Right:
                    return slideWidth;
                case AlignmentType.Bottom:
                    return slideHeight;
                case AlignmentType.CenterHorizontal:
                    return slideWidth / 2;
                case AlignmentType.CenterVertical:
                    return slideHeight / 2;
                default:
                    return CalculateSelectionReference(shapes, alignment);
            }
        }

        private float CalculateSingleShapeReference(IShapeContext shape, AlignmentType alignment)
        {
            if (shape == null)
            {
                return 0f;
            }

            var bounds = shape.Bounds;
            switch (alignment)
            {
                case AlignmentType.Left:
                    return bounds.Left;
                case AlignmentType.Right:
                    return bounds.Right;
                case AlignmentType.Top:
                    return bounds.Top;
                case AlignmentType.Bottom:
                    return bounds.Bottom;
                case AlignmentType.CenterHorizontal:
                    return bounds.CenterX;
                case AlignmentType.CenterVertical:
                    return bounds.CenterY;
                default:
                    return 0f;
            }
        }

        private float CalculateSelectionReference(List<IShapeContext> shapes, AlignmentType alignment)
        {
            switch (alignment)
            {
                case AlignmentType.Left:
                    return shapes.Min(s => s.Bounds.Left);
                case AlignmentType.Right:
                    return shapes.Max(s => s.Bounds.Right);
                case AlignmentType.Top:
                    return shapes.Min(s => s.Bounds.Top);
                case AlignmentType.Bottom:
                    return shapes.Max(s => s.Bounds.Bottom);
                case AlignmentType.CenterHorizontal:
                    var minX = shapes.Min(s => s.Bounds.Left);
                    var maxX = shapes.Max(s => s.Bounds.Right);
                    return (minX + maxX) / 2;
                case AlignmentType.CenterVertical:
                    var minY = shapes.Min(s => s.Bounds.Top);
                    var maxY = shapes.Max(s => s.Bounds.Bottom);
                    return (minY + maxY) / 2;
                default:
                    return 0f;
            }
        }

        private ShapeRect ApplyAlignment(ShapeRect bounds, AlignmentType alignment, float refValue)
        {
            switch (alignment)
            {
                case AlignmentType.Left:
                    return new ShapeRect(refValue, bounds.Top, bounds.Width, bounds.Height);
                case AlignmentType.Right:
                    return new ShapeRect(refValue - bounds.Width, bounds.Top, bounds.Width, bounds.Height);
                case AlignmentType.Top:
                    return new ShapeRect(bounds.Left, refValue, bounds.Width, bounds.Height);
                case AlignmentType.Bottom:
                    return new ShapeRect(bounds.Left, refValue - bounds.Height, bounds.Width, bounds.Height);
                case AlignmentType.CenterHorizontal:
                    return new ShapeRect(refValue - bounds.Width / 2, bounds.Top, bounds.Width, bounds.Height);
                case AlignmentType.CenterVertical:
                    return new ShapeRect(bounds.Left, refValue - bounds.Height / 2, bounds.Width, bounds.Height);
                default:
                    return bounds;
            }
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

        private void DistributeHorizontally(List<IShapeContext> shapes)
        {
            var sorted = shapes.OrderBy(s => s.Bounds.Left).ToList();
            var leftmost = sorted.First().Bounds.Left;
            var rightmost = sorted.Last().Bounds.Right;
            var totalShapeWidth = shapes.Sum(s => s.Bounds.Width);
            var totalSpace = rightmost - leftmost - totalShapeWidth;
            var gap = totalSpace / (shapes.Count - 1);

            float currentX = leftmost;
            foreach (var shape in sorted)
            {
                var bounds = shape.Bounds;
                shape.Bounds = new ShapeRect(currentX, bounds.Top, bounds.Width, bounds.Height);
                currentX += bounds.Width + gap;
            }
        }

        private void DistributeVertically(List<IShapeContext> shapes)
        {
            var sorted = shapes.OrderBy(s => s.Bounds.Top).ToList();
            var topmost = sorted.First().Bounds.Top;
            var bottommost = sorted.Last().Bounds.Bottom;
            var totalShapeHeight = shapes.Sum(s => s.Bounds.Height);
            var totalSpace = bottommost - topmost - totalShapeHeight;
            var gap = totalSpace / (shapes.Count - 1);

            float currentY = topmost;
            foreach (var shape in sorted)
            {
                var bounds = shape.Bounds;
                shape.Bounds = new ShapeRect(bounds.Left, currentY, bounds.Width, bounds.Height);
                currentY += bounds.Height + gap;
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

            // 计算中心点
            var center1X = bounds1.CenterX;
            var center1Y = bounds1.CenterY;
            var center2X = bounds2.CenterX;
            var center2Y = bounds2.CenterY;

            // 交换位置（保持各自的大小，只交换中心点）
            shape1.Bounds = new ShapeRect(
                center2X - bounds1.Width / 2,
                center2Y - bounds1.Height / 2,
                bounds1.Width,
                bounds1.Height);

            shape2.Bounds = new ShapeRect(
                center1X - bounds2.Width / 2,
                center1Y - bounds2.Height / 2,
                bounds2.Width,
                bounds2.Height);

            _logger.LogInformation("已交换两个形状的位置");
        }
    }
}
