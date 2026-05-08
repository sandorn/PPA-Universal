using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 文本批量操作服务实现
	/// </summary>
	public class TextBatchService : ITextBatchService
	{
		private readonly ILogger _logger;
		private readonly IShapeOperations _shapeOps;
		private readonly ITextShapeTextOperations _textOps;

		public TextBatchService(ILogger logger, IShapeOperations shapeOps, ITextShapeTextOperations textOps)
		{
			_logger = logger ?? NullLogger.Instance;
			_shapeOps = shapeOps;
			_textOps = textOps;
		}

		public void FormatSelectedText(IApplicationContext context, TextFormatOptions options)
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

			_logger.LogInformation($"格式化 {shapes.Count} 个形状的文本");

			foreach (var shape in shapes)
			{
				if (shape?.HasTextFrame != true) continue;

				try
				{
					FormatShapeText(shape, options);
					_logger.LogInformation($"格式化形状文本: {shape.Name}");
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"格式化形状文本失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation("文本格式化完成");
		}

		public void ReplaceText(IApplicationContext context, string find, string replace)
		{
			if (context?.ActivePresentation == null)
			{
				_logger.LogWarning("无法获取演示文稿");
				return;
			}

			if (string.IsNullOrEmpty(find))
			{
				_logger.LogWarning("查找文本不能为空");
				return;
			}

			replace ??= string.Empty;
			_logger.LogInformation($"批量替换文本: '{find}' -> '{replace}'");

			var pres = context.ActivePresentation;
			int n = pres.SlideCount;
			for (int i = 1; i <= n; i++)
			{
				try
				{
					var slide = pres.GetSlide(i);
					if (slide?.Shapes == null) continue;
					foreach (var shape in slide.Shapes)
						ReplaceInShapeRecursive(shape, find, replace);
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"替换文本时处理幻灯片 {i} 失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation("文本替换完成");
		}

		public void FormatTextBoxFont(IEnumerable<IShapeContext> shapes, FontStyle fontStyle = null)
		{
			var shapeList = shapes?.ToList();
			if (shapeList == null || shapeList.Count == 0)
			{
				_logger.LogWarning("没有选中任何形状");
				return;
			}

			if (fontStyle == null)
			{
				_logger.LogWarning("未提供字体样式");
				return;
			}

			_logger.LogInformation($"格式化 {shapeList.Count} 个文本框的字体");

			foreach (var shape in shapeList)
			{
				if (shape?.HasTextFrame != true) continue;

				try
				{
					FormatTextBoxFont(shape, fontStyle);
					_logger.LogInformation($"格式化文本框字体: {shape.Name}");
				}
				catch (System.Exception ex)
				{
					_logger.LogError($"格式化文本框字体失败: {ex.Message}", ex);
				}
			}

			_logger.LogInformation("文本框字体格式化完成");
		}

		private void FormatShapeText(IShapeContext shape, TextFormatOptions options)
		{
			if (shape?.NativeShape == null || !shape.HasTextFrame) return;

			_textOps.ApplyTextFormatOptions(shape.NativeShape, options);
		}

		private void ReplaceInShapeRecursive(IShapeContext shape, string find, string replace)
		{
			if (shape?.NativeShape == null || _shapeOps == null) return;

			try
			{
				if (_shapeOps.IsGroup(shape.NativeShape))
				{
					foreach (var child in _shapeOps.GetGroupChildShapes(shape.NativeShape))
					{
						var wrapped = _shapeOps.WrapShape(child);
						if (wrapped != null)
							ReplaceInShapeRecursive(wrapped, find, replace);
					}
					return;
				}

				if (shape.IsTable && shape.Table != null)
				{
					ReplaceInTable(shape.Table, find, replace);
					return;
				}

				if (shape.HasTextFrame)
					_textOps.TryReplaceTextInTextFrame(shape.NativeShape, find, replace);
			}
			catch (System.Exception ex)
			{
				_logger.LogError($"替换形状文本失败 ({shape.Name}): {ex.Message}", ex);
			}
		}

		private void ReplaceInTable(ITableContext table, string find, string replace)
		{
			for (int r = 1; r <= table.RowCount; r++)
			{
				for (int c = 1; c <= table.ColumnCount; c++)
				{
					try
					{
						var cell = table.GetCell(r, c);
						var t = cell?.Text;
						if (string.IsNullOrEmpty(t) || !t.Contains(find))
							continue;
						cell.Text = t.Replace(find, replace);
					}
					catch (System.Exception ex)
					{
						_logger.LogError($"替换表格单元格 ({r},{c}) 失败: {ex.Message}", ex);
					}
				}
			}
		}

		private void FormatTextBoxFont(IShapeContext shape, FontStyle fontStyle)
		{
			if (shape?.NativeShape == null || !shape.HasTextFrame) return;
			_textOps.ApplyTextBoxFont(shape.NativeShape, fontStyle);
		}
	}
}
