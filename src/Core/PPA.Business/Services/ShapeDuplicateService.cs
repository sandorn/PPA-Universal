using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 形状复制服务实现
	/// </summary>
	public class ShapeDuplicateService : IShapeDuplicateService
	{
		private readonly ILogger _logger;
		private readonly IShapeOperations _shapeOps;
		private readonly IApplicationContext _context;

		public ShapeDuplicateService(ILogger logger, IShapeOperations shapeOps, IApplicationContext context)
		{
			_logger = logger ?? NullLogger.Instance;
			_shapeOps = shapeOps;
			_context = context;
		}

		public IEnumerable<IShapeContext> MatrixCopy(IEnumerable<IShapeContext> shapes, int rows, int columns, float rowSpacing, float columnSpacing)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count == 0)
			{
				_logger.LogWarning("没有选中任何形状");
				return Enumerable.Empty<IShapeContext>();
			}

			if (rows < 1 || columns < 1)
			{
				_logger.LogWarning("行数和列数必须大于0");
				return Enumerable.Empty<IShapeContext>();
			}

			_logger.LogInformation($"矩阵复制: {shapeList.Count} 个形状, {rows}行 x {columns}列, 行间距: {rowSpacing}, 列间距: {columnSpacing}");

			var allCopies = new List<IShapeContext>();
			var slide = GetCurrentSlide();

			foreach (var originalShape in shapeList)
			{
				if (originalShape?.NativeShape == null) continue;

				var originalBounds = originalShape.Bounds;
				var shapeWidth = originalBounds.Width;
				var shapeHeight = originalBounds.Height;

				// 计算起始位置（原始形状的位置）
				float startX = originalBounds.Left;
				float startY = originalBounds.Top;

				for (int row = 0; row < rows; row++)
				{
					for (int col = 0; col < columns; col++)
					{
						// 跳过原始位置（第一行第一列）
						if (row == 0 && col == 0) continue;

						// 计算新位置
						float newX = startX + col * (shapeWidth + columnSpacing);
						float newY = startY + row * (shapeHeight + rowSpacing);

						// 复制形状
						var copiedShape = _shapeOps.CopyShape(originalShape.NativeShape);
						if (copiedShape != null && slide != null)
						{
							// 设置新位置
							var newBounds = new ShapeRect(newX, newY, shapeWidth, shapeHeight);
							_shapeOps.SetBounds(copiedShape, newBounds);

							_logger.LogInformation($"复制形状到位置 ({newX}, {newY})");
						}
					}
				}
			}

			_logger.LogInformation($"矩阵复制完成，共创建 {allCopies.Count} 个副本");
			return allCopies;
		}

		public IEnumerable<IShapeContext> LinearCopy(IEnumerable<IShapeContext> shapes, int count, float spacing, LinearCopyDirection direction)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count == 0)
			{
				_logger.LogWarning("没有选中任何形状");
				return Enumerable.Empty<IShapeContext>();
			}

			if (count < 1)
			{
				_logger.LogWarning("复制数量必须大于0");
				return Enumerable.Empty<IShapeContext>();
			}

			_logger.LogInformation($"线性复制: {shapeList.Count} 个形状, 数量: {count}, 间距: {spacing}, 方向: {direction}");

			var allCopies = new List<IShapeContext>();
			var slide = GetCurrentSlide();

			foreach (var originalShape in shapeList)
			{
				if (originalShape?.NativeShape == null) continue;

				var originalBounds = originalShape.Bounds;
				var shapeWidth = originalBounds.Width;
				var shapeHeight = originalBounds.Height;

				float startX = originalBounds.Left;
				float startY = originalBounds.Top;

				for (int i = 1; i <= count; i++)
				{
					float newX = startX;
					float newY = startY;

					if (direction == LinearCopyDirection.Horizontal)
					{
						// 水平方向：向右复制
						newX = startX + i * (shapeWidth + spacing);
					}
					else
					{
						// 垂直方向：向下复制
						newY = startY + i * (shapeHeight + spacing);
					}

					// 复制形状
					var copiedShape = _shapeOps.CopyShape(originalShape.NativeShape);
					if (copiedShape != null && slide != null)
					{
						// 设置新位置
						var newBounds = new ShapeRect(newX, newY, shapeWidth, shapeHeight);
						_shapeOps.SetBounds(copiedShape, newBounds);

						_logger.LogInformation($"复制形状到位置 ({newX}, {newY})");
					}
				}
			}

			_logger.LogInformation($"线性复制完成，共创建 {allCopies.Count} 个副本");
			return allCopies;
		}

		private ISlideContext GetCurrentSlide()
		{
			try
			{
				return _context?.ActiveWindow?.ActiveSlide;
			}
			catch
			{
				return null;
			}
		}
	}
}

