using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using PPA.Logging;
using PPA.Universal.Integration;
using stdole;

namespace PPA.Universal.ComAddIn
{
	/// <summary>
	/// Ribbon 回调处理类
	/// </summary>
	[ComVisible(true)]
	public class RibbonCallbacks
	{
		private object _ribbon;
		private AlignmentReference _currentReference = AlignmentReference.SelectedObjects;

		/// <summary>与 <c>PPARibbon.xml</c> 中 <c>ddAlignRef</c> 的 item 顺序一致。</summary>
		private static readonly AlignmentReference[] AlignRefDropdownOrder =
		{
			AlignmentReference.SelectedObjects,
			AlignmentReference.Slide,
			AlignmentReference.FirstObject,
			AlignmentReference.LastObject,
		};

		private static AlignmentReference AlignRefFromDropdownIndex(int index) =>
			(uint)index < (uint)AlignRefDropdownOrder.Length
				? AlignRefDropdownOrder[index]
				: AlignmentReference.SelectedObjects;

		private static int DropdownIndexOfAlignRef(AlignmentReference reference)
		{
			for (var i = 0; i < AlignRefDropdownOrder.Length; i++)
			{
				if (AlignRefDropdownOrder[i] == reference)
					return i;
			}
			return 0;
		}

		/// <summary>
		/// Ribbon 加载时调用
		/// </summary>
		public void Ribbon_OnLoad(object ribbon)
		{
			_ribbon = ribbon;
			UniversalIntegration.Logger?.LogInformation("Ribbon loaded successfully");
		}

		#region Ribbon 资源（图标）

		/// <summary>
		/// RibbonX loadImage 回调：按 image="xxx.png" 加载自定义图标
		/// </summary>
		public IPictureDisp LoadImage(string imageId)
		{
			try
			{
				if (string.IsNullOrWhiteSpace(imageId))
				{
					return null;
				}

				var assembly = Assembly.GetExecutingAssembly();
				var resourceName = $"PPA.Universal.ComAddIn.Resources.{imageId}";

				using (var stream = assembly.GetManifestResourceStream(resourceName))
				{
					if (stream == null)
					{
						UniversalIntegration.Logger?.LogDebug($"图标资源未找到: {resourceName}");
						return null;
					}

					using (var bitmap = new Bitmap(stream))
					{
						return PictureDispConverter.ImageToPictureDisp(bitmap);
					}
				}
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"加载图标失败: {ex.Message}", ex);
				return null;
			}
		}

		private class PictureDispConverter : AxHost
		{
			private PictureDispConverter() : base("") { }

			public static IPictureDisp ImageToPictureDisp(Image image)
			{
				if (image == null) return null;
				return (IPictureDisp)GetIPictureDispFromPicture(image);
			}
		}

		#endregion

		#region 对齐参考（下拉）

		public void OnAlignRefChanged(object control, string selectedId, int selectedIndex)
		{
			_currentReference = AlignRefFromDropdownIndex(selectedIndex);
			UniversalIntegration.Logger?.LogInformation($"Alignment reference changed to: {_currentReference}");
		}

		/// <summary>Ribbon <c>getSelectedItemIndex</c>，必须与下拉项顺序一致。</summary>
		public int GetAlignRefIndex(object control) => DropdownIndexOfAlignRef(_currentReference);

		#endregion

		#region 对齐操作

		public void OnAlignLeft(object control)
		{
			ExecuteAlignment(AlignmentType.Left);
		}

		public void OnAlignRight(object control)
		{
			ExecuteAlignment(AlignmentType.Right);
		}

		public void OnAlignTop(object control)
		{
			ExecuteAlignment(AlignmentType.Top);
		}

		public void OnAlignBottom(object control)
		{
			ExecuteAlignment(AlignmentType.Bottom);
		}

		public void OnAlignCenterH(object control)
		{
			ExecuteAlignment(AlignmentType.CenterHorizontal);
		}

		public void OnAlignCenterV(object control)
		{
			ExecuteAlignment(AlignmentType.CenterVertical);
		}

		#endregion

		#region 分布操作

		public void OnDistributeH(object control)
		{
			ExecuteDistribution(DistributionType.Horizontal);
		}

		public void OnDistributeV(object control)
		{
			ExecuteDistribution(DistributionType.Vertical);
		}

		#endregion

		#region 尺寸操作

		public void OnEqualWidth(object control)
		{
			ExecuteSizeOperation(s => s.SetEqualWidth(GetSelectedShapes()), "等宽");
		}

		public void OnEqualHeight(object control)
		{
			ExecuteSizeOperation(s => s.SetEqualHeight(GetSelectedShapes()), "等高");
		}

		public void OnEqualSize(object control)
		{
			ExecuteSizeOperation(s => s.SetEqualSize(GetSelectedShapes()), "等大小");
		}

		#endregion

		#region 形状吸附

		public void OnSnapLeft(object control)
		{
			ExecuteSnap(SnapDirection.Left);
		}

		public void OnSnapRight(object control)
		{
			ExecuteSnap(SnapDirection.Right);
		}

		public void OnSnapTop(object control)
		{
			ExecuteSnap(SnapDirection.Top);
		}

		public void OnSnapBottom(object control)
		{
			ExecuteSnap(SnapDirection.Bottom);
		}

		#endregion

		#region 延伸对齐

		public void OnExtendLeft(object control)
		{
			ExecuteExtend(ExtendDirection.Left);
		}

		public void OnExtendRight(object control)
		{
			ExecuteExtend(ExtendDirection.Right);
		}

		public void OnExtendTop(object control)
		{
			ExecuteExtend(ExtendDirection.Top);
		}

		public void OnExtendBottom(object control)
		{
			ExecuteExtend(ExtendDirection.Bottom);
		}

		#endregion

		#region 交换位置和大小

		public void OnSwapPositionsAndSize(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || shapes.Count != 2)
				{
					ShowWarning("交换大小和位置需要选择恰好 2 个形状");
					return;
				}

				var service = UniversalIntegration.GetService<IAlignmentService>();
				if (service == null)
				{
					ShowWarning("对齐服务不可用");
					return;
				}

				using (CreateUndoScope("交换大小和位置"))
				{
					service.SwapPositionsAndSize(shapes[0], shapes[1]);
				}

				UniversalIntegration.Logger?.LogInformation("交换大小和位置完成");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"交换大小和位置失败: {ex.Message}", ex);
				ShowError($"交换大小和位置失败: {ex.Message}");
			}
		}

		#endregion

		#region 形状复制

		public void OnMatrixCopy(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || !shapes.Any())
				{
					ShowWarning("请先选择要复制的形状");
					return;
				}

				var cfgRoot = UniversalIntegration.GetService<PPAConfig>();
				var dupCfg = cfgRoot?.Duplicate;
				if (dupCfg == null)
				{
					dupCfg = new DuplicateConfig();
					if (cfgRoot != null)
						cfgRoot.Duplicate = dupCfg;
				}

				if (!DuplicateCopyDialogs.TryShowMatrixDialog(dupCfg, out int rows, out int columns, out float rowSpacing, out float columnSpacing))
					return;

				var duplicateService = UniversalIntegration.GetService<IShapeDuplicateService>();
				if (duplicateService == null)
				{
					ShowWarning("复制服务不可用");
					return;
				}

				using (CreateUndoScope("矩阵复制"))
				{
					duplicateService.MatrixCopy(shapes, rows, columns, rowSpacing, columnSpacing);
				}

				UniversalIntegration.Logger?.LogInformation($"矩阵复制完成: {rows}行 x {columns}列");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"矩阵复制失败: {ex.Message}", ex);
				ShowError($"矩阵复制失败: {ex.Message}");
			}
		}

		public void OnLinearCopy(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || !shapes.Any())
				{
					ShowWarning("请先选择要复制的形状");
					return;
				}

				var cfgRoot = UniversalIntegration.GetService<PPAConfig>();
				var dupCfg = cfgRoot?.Duplicate;
				if (dupCfg == null)
				{
					dupCfg = new DuplicateConfig();
					if (cfgRoot != null)
						cfgRoot.Duplicate = dupCfg;
				}

				if (!DuplicateCopyDialogs.TryShowLinearDialog(dupCfg, out int count, out float spacing, out LinearCopyDirection direction))
					return;

				var duplicateService = UniversalIntegration.GetService<IShapeDuplicateService>();
				if (duplicateService == null)
				{
					ShowWarning("复制服务不可用");
					return;
				}

				using (CreateUndoScope("线性复制"))
				{
					duplicateService.LinearCopy(shapes, count, spacing, direction);
				}

				UniversalIntegration.Logger?.LogInformation($"线性复制完成: {count}个, 方向: {direction}");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"线性复制失败: {ex.Message}", ex);
				ShowError($"线性复制失败: {ex.Message}");
			}
		}

		#endregion

		#region 对象隐藏与显示

		public void OnHideOrShowShapes(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();

				var visibilityService = UniversalIntegration.GetService<IShapeVisibilityService>();
				if (visibilityService == null)
				{
					ShowWarning("可见性服务不可用");
					return;
				}

				if (shapes != null && shapes.Any())
				{
					// 有选中形状，隐藏它们
					using (CreateUndoScope("隐藏对象"))
					{
						visibilityService.HideShapes(shapes);
					}
					UniversalIntegration.Logger?.LogInformation($"已隐藏 {shapes.Count} 个对象");
				}
				else
				{
					// 没有选中形状，显示所有隐藏的对象
					var context = UniversalIntegration.Context;
					if (context?.ActiveWindow?.ActiveSlide == null)
					{
						ShowWarning("无法获取当前幻灯片");
						return;
					}

					using (CreateUndoScope("显示所有对象"))
					{
						visibilityService.ShowAllHiddenShapes(context.ActiveWindow.ActiveSlide);
					}
					UniversalIntegration.Logger?.LogInformation("已显示所有隐藏的对象");
				}
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"隐藏/显示操作失败: {ex.Message}", ex);
				ShowError($"隐藏/显示操作失败: {ex.Message}");
			}
		}

		#endregion

		#region 创建无边框矩形

		public void OnCreateRectangle(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();
				var creationService = UniversalIntegration.GetService<IShapeCreationService>();
				if (creationService == null)
				{
					ShowWarning("形状创建服务不可用");
					return;
				}

				if (shapes != null && shapes.Any())
				{
					// 在选中形状位置创建矩形
					using (CreateUndoScope("创建无边框矩形"))
					{
						creationService.CreateRectanglesAtShapes(shapes);
					}
					UniversalIntegration.Logger?.LogInformation($"已在 {shapes.Count} 个形状位置创建矩形");
				}
				else
				{
					// 在幻灯片上创建矩形
					var context = UniversalIntegration.Context;
					if (context?.ActiveWindow?.ActiveSlide == null)
					{
						ShowWarning("无法获取当前幻灯片");
						return;
					}

					var slides = new[] { context.ActiveWindow.ActiveSlide };
					using (CreateUndoScope("创建无边框矩形"))
					{
						creationService.CreateRectanglesOnSlides(slides);
					}
					UniversalIntegration.Logger?.LogInformation("已在当前幻灯片创建矩形");
				}
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"创建矩形失败: {ex.Message}", ex);
				ShowError($"创建矩形失败: {ex.Message}");
			}
		}

		#endregion

		#region 边角料裁除

		public void OnCropEdges(object control)
		{
			try
			{
				var cropService = UniversalIntegration.GetService<ICropService>();
				if (cropService == null)
				{
					ShowWarning("裁除服务不可用");
					return;
				}

				var context = UniversalIntegration.Context;
				if (context?.ActiveWindow?.ActiveSlide == null)
				{
					ShowWarning("无法获取当前幻灯片");
					return;
				}

				var shapes = GetSelectedShapes();
				if (shapes != null && shapes.Any())
				{
					// 裁除选中对象超出页面的部分
					using (CreateUndoScope("裁除边角料"))
					{
						cropService.CropShapesToSlide(shapes, context.ActiveWindow.ActiveSlide);
					}
					UniversalIntegration.Logger?.LogInformation($"已裁除 {shapes.Count} 个对象的边角料");
				}
				else
				{
					// 裁除当前幻灯片所有对象超出页面的部分
					using (CreateUndoScope("裁除边角料"))
					{
						cropService.CropAllShapesToSlide(context.ActiveWindow.ActiveSlide);
					}
					UniversalIntegration.Logger?.LogInformation("已裁除当前幻灯片所有对象的边角料");
				}
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"裁除边角料失败: {ex.Message}", ex);
				ShowError($"裁除边角料失败: {ex.Message}");
			}
		}

		#endregion

		#region 格式化功能

		public void OnFormatTableFont(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || !shapes.Any())
				{
					ShowWarning("请先选择包含表格的形状");
					return;
				}

				var tableShapes = shapes.Where(s => s?.IsTable == true && s.Table != null).ToList();
				if (tableShapes.Count == 0)
				{
					ShowWarning("选中形状中没有表格");
					return;
				}

				var formatService = UniversalIntegration.GetService<ITableFormatService>();
				if (formatService == null)
				{
					ShowWarning("表格格式化服务不可用");
					return;
				}

				using (CreateUndoScope("格式化表格字体"))
				{
					foreach (var shape in tableShapes)
					{
						formatService.FormatTableFont(shape.Table);
					}
				}

				UniversalIntegration.Logger?.LogInformation($"已格式化 {tableShapes.Count} 个表格的字体");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"格式化表格字体失败: {ex.Message}", ex);
				ShowError($"格式化表格字体失败: {ex.Message}");
			}
		}

		public void OnFormatTextBoxFont(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || !shapes.Any())
				{
					ShowWarning("请先选择文本框");
					return;
				}

				var textBoxShapes = shapes.Where(s => s?.HasTextFrame == true).ToList();
				if (textBoxShapes.Count == 0)
				{
					ShowWarning("选中形状中没有文本框");
					return;
				}

				var textService = UniversalIntegration.GetService<ITextBatchService>();
				if (textService == null)
				{
					ShowWarning("文本服务不可用");
					return;
				}

				var cfgRoot = UniversalIntegration.GetService<PPAConfig>();
				var fontStyle = cfgRoot?.Text?.Font?.ToFontStyle() ?? PpaConfigTemplateFallbacks.TextBoxRibbonFontStyle();

				using (CreateUndoScope("格式化文本框字体"))
				{
					textService.FormatTextBoxFont(textBoxShapes, fontStyle);
				}

				UniversalIntegration.Logger?.LogInformation($"已格式化 {textBoxShapes.Count} 个文本框的字体");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"格式化文本框字体失败: {ex.Message}", ex);
				ShowError($"格式化文本框字体失败: {ex.Message}");
			}
		}

		public void OnFormatChartFont(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || !shapes.Any())
				{
					ShowWarning("请先选择图表");
					return;
				}

				var chartShapes = shapes.Where(s => s?.IsChart == true).ToList();
				if (chartShapes.Count == 0)
				{
					ShowWarning("选中形状中没有图表");
					return;
				}

				var chartService = UniversalIntegration.GetService<IChartBatchService>();
				if (chartService == null)
				{
					ShowWarning("图表服务不可用");
					return;
				}

				using (CreateUndoScope("格式化图表字体"))
				{
					// 由业务层从 PPAConfig 读取标题/图例字体；缺失时使用内置固定默认值
					chartService.FormatChartFont(chartShapes, null);
				}

				UniversalIntegration.Logger?.LogInformation($"已格式化 {chartShapes.Count} 个图表的字体");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"格式化图表字体失败: {ex.Message}", ex);
				ShowError($"格式化图表字体失败: {ex.Message}");
			}
		}

		/// <summary>
		/// 全文查找替换（两栏对话框 + <see cref="ITextBatchService.ReplaceText"/>）。
		/// </summary>
		public void OnFindReplaceText(object control)
		{
			try
			{
				if (!FindReplaceDialog.TryShow(out var find, out var replace))
					return;

				if (string.IsNullOrEmpty(find))
				{
					ShowWarning("查找内容不能为空");
					return;
				}

				var context = UniversalIntegration.Context;
				if (context?.ActivePresentation == null)
				{
					ShowWarning("无法获取当前演示文稿");
					return;
				}

				var textService = UniversalIntegration.GetService<ITextBatchService>();
				if (textService == null)
				{
					ShowWarning("文本服务不可用");
					return;
				}

				using (CreateUndoScope("查找替换"))
				{
					textService.ReplaceText(context, find, replace);
				}

				UniversalIntegration.Logger?.LogInformation("查找替换已完成");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"查找替换失败: {ex.Message}", ex);
				ShowError($"查找替换失败: {ex.Message}");
			}
		}

		#endregion

		#region 表格格式化

		public void OnFormatThreeLineTable(object control)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || !shapes.Any())
				{
					ShowWarning("请先选择包含表格的形状");
					return;
				}

				var tableShapes = shapes.Where(s => s?.IsTable == true && s.Table != null).ToList();
				if (tableShapes.Count == 0)
				{
					ShowWarning("选中形状中没有表格");
					return;
				}

				var formatService = UniversalIntegration.GetService<ITableFormatService>();
				if (formatService == null)
				{
					ShowWarning("表格格式化服务不可用");
					return;
				}

				// 使用撤销作用域包裹所有操作
				// PowerPoint: 创建撤销边界
				// WPS: 使用 BeginUndoGroup/EndUndoGroup 配对
				var context = UniversalIntegration.Context;
				using (CreateUndoScope("三线表格式化"))
				{
					TryClearTableFormatMenu(context);

					foreach (var shape in tableShapes)
					{
						formatService.FormatTableAsThreeLine(shape.Table);
					}
				}

				UniversalIntegration.Logger?.LogInformation($"已对 {tableShapes.Count} 个表格应用三线表格式");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"三线表格式化失败: {ex.Message}", ex);
				ShowError($"三线表格式化失败: {ex.Message}");
			}
		}

		/// <summary>全稿全部表格三线表（<see cref="ITableBatchService.FormatAllTables"/>），确认后执行。</summary>
		public void OnBatchThreeLineAll(object control)
		{
			try
			{
				var confirm = ComDialogChrome.ConfirmWarning(
					"将把当前演示文稿中全部表格设置为三线表样式（含各页组合内表格）。是否继续？");
				if (confirm != DialogResult.Yes)
					return;

				var context = UniversalIntegration.Context;
				var batch = UniversalIntegration.GetService<ITableBatchService>();
				if (batch == null || context?.ActivePresentation == null)
				{
					ShowWarning("批量表格服务不可用或无法获取演示文稿");
					return;
				}

				using (CreateUndoScope("全稿表格三线表"))
				{
					TryClearTableFormatMenu(context);
					batch.FormatAllTables(context);
				}

				UniversalIntegration.Logger?.LogInformation("全稿表格三线表批量处理完成");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"全稿表格三线表失败: {ex.Message}", ex);
				ShowError($"全稿表格三线表失败: {ex.Message}");
			}
		}

		private void TryClearTableFormatMenu(IApplicationContext context)
		{
			try
			{
				var idMsoExecutor = UniversalIntegration.GetService<IIdMsoCommandExecutor>();
				idMsoExecutor?.TryExecute(context, "ClearMenu");
			}
			catch
			{
				// 与单表三线表一致：清除失败时忽略
			}
		}

		#endregion

		#region 毛玻璃卡片

		public void OnCreateGlassCard(object control)
		{
			try
			{
				var context = UniversalIntegration.Context;
				if (context == null)
				{
					ShowWarning("应用上下文不可用，请重启加载项后重试");
					return;
				}

				var service = UniversalIntegration.GetService<IGlassCardService>();
				if (service == null)
				{
					ShowWarning("毛玻璃卡片服务不可用");
					return;
				}

				// 使用撤销作用域包裹所有操作
				using (CreateUndoScope("创建毛玻璃卡片"))
				{
					// 业务层会自动从全局 PPAConfig 中读取 GlassCard 配置，这里传 null 即可
					service.CreateGlassCard(context, null);
				}

				UniversalIntegration.Logger?.LogInformation("GlassCard creation requested from Ribbon");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"CreateGlassCard failed: {ex.Message}", ex);
				ShowError($"创建毛玻璃卡片失败: {ex.Message}");
			}
		}

		#endregion

		#region 调试

		public void OnDebug(object control)
		{
			try
			{
				var platform = UniversalIntegration.Platform;
				var platformName = platform switch
				{
					PPA.Core.Abstraction.PlatformType.PowerPoint => "PowerPoint",
					PPA.Core.Abstraction.PlatformType.WPS => "WPS",
					_ => "Unknown"
				};

				var message = $"调试按钮已触发\n平台: {platformName}\n时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

				UniversalIntegration.Logger?.LogInformation($"[Debug] Platform={platformName}, triggered at {DateTime.Now:O}");

				// 平台特定逻辑占位
				if (platform == PPA.Core.Abstraction.PlatformType.WPS)
				{
					message += "\n\n[WPS 特定调试逻辑占位]";
					UniversalIntegration.Logger?.LogInformation("[Debug] WPS-specific logic placeholder");
				}
				else if (platform == PPA.Core.Abstraction.PlatformType.PowerPoint)
				{
					message += "\n\n[PowerPoint 特定调试逻辑占位]";
					UniversalIntegration.Logger?.LogInformation("[Debug] PowerPoint-specific logic placeholder");
				}

				ShowMessage(message);
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"[Debug] Error: {ex.Message}", ex);
				ShowError($"调试失败: {ex.Message}");
			}
		}

		#endregion

		#region 辅助方法

		private List<IShapeContext> GetSelectedShapes()
		{
			try
			{
				var context = UniversalIntegration.Context;
				var selection = context?.Selection;

				if (selection == null)
				{
					return null;
				}

				if (selection.Type != SelectionType.Shapes && selection.ShapeCount == 0)
				{
					return null;
				}

				return selection.SelectedShapes?.ToList();
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"GetSelectedShapes failed: {ex.Message}", ex);
				return null;
			}
		}

		private void ShowMessage(string message) => ComDialogChrome.NotifyInfo(message);

		private void ShowWarning(string message) => ComDialogChrome.NotifyWarning(message);

		private void ShowError(string message) => ComDialogChrome.NotifyError(message);

		private IDisposable CreateUndoScope(string operationName)
		{
			try
			{
				var undoService = UniversalIntegration.GetService<IUndoService>();
				var context = UniversalIntegration.Context;
				if (undoService != null && context != null)
				{
					return undoService.CreateUndoScope(context, operationName);
				}
			}
			catch
			{
				// 撤销作用域创建失败不应影响主操作
			}
			return new NullDisposable();
		}

		private class NullDisposable : IDisposable
		{
			public void Dispose() { }
		}

		private void ExecuteAlignment(AlignmentType alignmentType)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || !shapes.Any())
				{
					ShowWarning("请先选择要对齐的形状");
					return;
				}

				var service = UniversalIntegration.GetService<IAlignmentService>();
				if (service == null)
				{
					ShowWarning("对齐服务不可用");
					return;
				}

				// 使用撤销作用域包裹操作
				using (CreateUndoScope($"对齐: {alignmentType}"))
				{
					service.Align(shapes, alignmentType, _currentReference);
				}

				UniversalIntegration.Logger?.LogInformation($"Alignment executed: {alignmentType}, Reference: {_currentReference}");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"Alignment failed: {ex.Message}", ex);
				ShowError($"对齐操作失败: {ex.Message}");
			}
		}

		private void ExecuteDistribution(DistributionType distributionType)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || shapes.Count < 3)
				{
					ShowWarning("分布操作需要选择至少 3 个形状");
					return;
				}

				var service = UniversalIntegration.GetService<IAlignmentService>();
				if (service == null)
				{
					ShowWarning("对齐服务不可用");
					return;
				}

				// 使用撤销作用域包裹操作
				using (CreateUndoScope($"分布: {distributionType}"))
				{
					service.Distribute(shapes, distributionType);
				}

				UniversalIntegration.Logger?.LogInformation($"Distribution executed: {distributionType}");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"Distribution failed: {ex.Message}", ex);
				ShowError($"分布操作失败: {ex.Message}");
			}
		}

		private void ExecuteSizeOperation(Action<IAlignmentService> operation, string operationName = "尺寸操作")
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || shapes.Count < 2)
				{
					ShowWarning("尺寸操作需要选择至少 2 个形状");
					return;
				}

				var service = UniversalIntegration.GetService<IAlignmentService>();
				if (service == null)
				{
					ShowWarning("对齐服务不可用");
					return;
				}

				// 使用撤销作用域包裹操作
				using (CreateUndoScope(operationName))
				{
					operation(service);
				}

				UniversalIntegration.Logger?.LogInformation("Size operation executed");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"Size operation failed: {ex.Message}", ex);
				ShowError($"尺寸操作失败: {ex.Message}");
			}
		}

		private void ExecuteSnap(SnapDirection direction)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || shapes.Count < 2)
				{
					ShowWarning("吸附操作需要选择至少 2 个形状");
					return;
				}

				var service = UniversalIntegration.GetService<IAlignmentService>();
				if (service == null)
				{
					ShowWarning("对齐服务不可用");
					return;
				}

				using (CreateUndoScope($"吸附: {direction}"))
				{
					service.SnapToShape(shapes, direction);
				}

				UniversalIntegration.Logger?.LogInformation($"吸附操作完成: {direction}");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"吸附操作失败: {ex.Message}", ex);
				ShowError($"吸附操作失败: {ex.Message}");
			}
		}

		private void ExecuteExtend(ExtendDirection direction)
		{
			try
			{
				var shapes = GetSelectedShapes();
				if (shapes == null || shapes.Count < 2)
				{
					ShowWarning("延伸对齐操作需要选择至少 2 个形状");
					return;
				}

				var service = UniversalIntegration.GetService<IAlignmentService>();
				if (service == null)
				{
					ShowWarning("对齐服务不可用");
					return;
				}

				using (CreateUndoScope($"延伸对齐: {direction}"))
				{
					service.ExtendAlignment(shapes, direction);
				}

				UniversalIntegration.Logger?.LogInformation($"延伸对齐操作完成: {direction}");
			}
			catch (Exception ex)
			{
				UniversalIntegration.Logger?.LogError($"延伸对齐操作失败: {ex.Message}", ex);
				ShowError($"延伸对齐操作失败: {ex.Message}");
			}
		}

		#endregion
	}
}
