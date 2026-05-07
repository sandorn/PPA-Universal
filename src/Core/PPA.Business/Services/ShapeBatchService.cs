using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 形状批量操作：删除、复制、统一尺寸（基于当前选择）。
	/// </summary>
	public class ShapeBatchService : IShapeBatchService
	{
		private readonly ILogger _logger;
		private readonly IShapeOperations _shapeOps;

		public ShapeBatchService(ILogger logger, IShapeOperations shapeOps)
		{
			_logger = logger ?? NullLogger.Instance;
			_shapeOps = shapeOps;
		}

		public void DeleteSelectedShapes(IApplicationContext context)
		{
			var natives = CollectSelectedNativeShapes(context);
			if (natives.Count == 0)
			{
				_logger.LogWarning("没有可删除的选中形状");
				return;
			}

			foreach (var s in natives)
			{
				try
				{
					_shapeOps.DeleteShape(s);
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"删除形状失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation($"已删除 {natives.Count} 个形状");
		}

		public void DuplicateSelectedShapes(IApplicationContext context)
		{
			var natives = CollectSelectedNativeShapes(context);
			if (natives.Count == 0)
			{
				_logger.LogWarning("没有可复制的选中形状");
				return;
			}

			int ok = 0;
			foreach (var s in natives)
			{
				try
				{
					var copy = _shapeOps.CopyShape(s);
					if (copy != null)
						ok++;
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"复制形状失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation($"已复制 {ok}/{natives.Count} 个形状");
		}

		public void ResizeSelectedShapes(IApplicationContext context, float width, float height)
		{
			if (width <= 0 || height <= 0)
			{
				_logger.LogWarning("宽度和高度必须大于 0");
				return;
			}

			var shapes = GetSelectedShapeContexts(context);
			if (shapes.Count == 0)
			{
				_logger.LogWarning("没有选中形状，跳过调整大小");
				return;
			}

			foreach (var shape in shapes)
			{
				if (shape?.NativeShape == null) continue;
				try
				{
					var b = _shapeOps.GetBounds(shape.NativeShape);
					_shapeOps.SetBounds(shape.NativeShape, new ShapeRect(b.Left, b.Top, width, height));
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"调整形状大小失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation($"已将 {shapes.Count} 个形状调整为 {width}x{height}");
		}

		private static List<IShapeContext> GetSelectedShapeContexts(IApplicationContext context)
		{
			var sel = context?.Selection;
			if (sel == null || sel.Type != SelectionType.Shapes || sel.ShapeCount == 0)
				return new List<IShapeContext>();
			return sel.SelectedShapes?.Where(s => s?.NativeShape != null).ToList() ?? new List<IShapeContext>();
		}

		private List<object> CollectSelectedNativeShapes(IApplicationContext context)
		{
			return GetSelectedShapeContexts(context).Select(s => s.NativeShape).ToList();
		}
	}
}
