using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 批量表格：对范围内所有表格应用与「三线表」相同的配置化格式（<see cref="ITableFormatService.FormatTableAsThreeLine"/>）。
	/// </summary>
	public class TableBatchService : ITableBatchService
	{
		private readonly ILogger _logger;
		private readonly ITableFormatService _tableFormat;
		private readonly IShapeOperations _shapeOps;

		public TableBatchService(ILogger logger, ITableFormatService tableFormat, IShapeOperations shapeOps)
		{
			_logger = logger ?? NullLogger.Instance;
			_tableFormat = tableFormat;
			_shapeOps = shapeOps;
		}

		public void FormatAllTables(IApplicationContext context)
		{
			if (context?.ActivePresentation == null)
			{
				_logger.LogWarning("无法获取演示文稿，跳过批量格式化表格");
				return;
			}

			var pres = context.ActivePresentation;
			int n = 0;
			for (int i = 1; i <= pres.SlideCount; i++)
			{
				try
				{
					var slide = pres.GetSlide(i);
					n += FormatTablesOnSlide(slide);
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"批量格式化：幻灯片 {i} 处理失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation($"全稿批量三线表格式已处理 {n} 个表格");
		}

		public void FormatSelectedTables(IApplicationContext context)
		{
			var shapes = GetSelectedShapes(context);
			if (shapes == null || shapes.Count == 0)
			{
				_logger.LogWarning("没有选中形状，跳过批量格式化表格");
				return;
			}

			int n = 0;
			foreach (var shape in EnumerateShapesDepthFirst(shapes))
			{
				if (shape.IsTable && shape.Table != null)
				{
					_tableFormat.FormatTableAsThreeLine(shape.Table);
					n++;
				}
			}

			_logger.LogInformation($"已对已选范围内 {n} 个表格应用三线表格式");
		}

		public void FormatCurrentSlideTables(IApplicationContext context)
		{
			var slide = context?.ActiveWindow?.ActiveSlide;
			if (slide == null)
			{
				_logger.LogWarning("无法获取当前幻灯片，跳过批量格式化表格");
				return;
			}

			int n = FormatTablesOnSlide(slide);
			_logger.LogInformation($"当前页已处理 {n} 个表格（三线表格式）");
		}

		private int FormatTablesOnSlide(ISlideContext slide)
		{
			int n = 0;
			foreach (var shape in EnumerateShapesDepthFirst(slide.Shapes))
			{
				if (shape.IsTable && shape.Table != null)
				{
					_tableFormat.FormatTableAsThreeLine(shape.Table);
					n++;
				}
			}

			return n;
		}

		private static List<IShapeContext> GetSelectedShapes(IApplicationContext context)
		{
			var sel = context?.Selection;
			if (sel == null || sel.Type != SelectionType.Shapes || sel.ShapeCount == 0)
				return null;
			return sel.SelectedShapes?.ToList();
		}

		private IEnumerable<IShapeContext> EnumerateShapesDepthFirst(IEnumerable<IShapeContext> roots)
		{
			if (roots == null) yield break;
			foreach (var root in roots)
			{
				foreach (var s in EnumerateShapeRecursive(root))
					yield return s;
			}
		}

		private IEnumerable<IShapeContext> EnumerateShapeRecursive(IShapeContext shape)
		{
			if (shape?.NativeShape == null || _shapeOps == null)
				yield break;

			if (_shapeOps.IsGroup(shape.NativeShape))
			{
				foreach (var child in _shapeOps.GetGroupChildShapes(shape.NativeShape))
				{
					var wrapped = _shapeOps.WrapShape(child);
					if (wrapped != null)
					{
						foreach (var inner in EnumerateShapeRecursive(wrapped))
							yield return inner;
					}
				}

				yield break;
			}

			yield return shape;
		}
	}
}
