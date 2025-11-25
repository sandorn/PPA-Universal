using NetOffice.OfficeApi.Enums;
using PPA.Adapter.PowerPoint;
using PPA.Business.Abstractions;
using PPA.Core;
using PPA.Core.Abstraction;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Logging;
using PPA.Legacy.Bridge;
using PPA.Logging;
using PPA.Manipulation.Config;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Manipulation
{
	/// <summary>
	/// 表格批量操作辅助类 提供表格批量格式化功能，支持同步和异步操作
	/// </summary>
	internal class TableBatchHelper(ITableFormatHelper tableFormatHelper,IShapeHelper shapeHelper,ILogger logger = null):ITableBatchHelper
	{
		private readonly ITableFormatHelper _tableFormatHelper = tableFormatHelper??throw new ArgumentNullException(nameof(tableFormatHelper));
		private readonly IShapeHelper _shapeHelper = shapeHelper??throw new ArgumentNullException(nameof(shapeHelper));
		private readonly ILogger _logger = logger??LoggerProvider.GetLogger();

		#region ITableBatchHelper 实现

		/// <summary>
		/// 格式化表格（同步方法）
		/// </summary>
		/// <param name="netApp"> PowerPoint 应用程序对象 </param>
		public void FormatTables(NETOP.Application netApp)
		{
			if(netApp==null) throw new ArgumentNullException(nameof(netApp));
			FormatTablesInternal(netApp,_tableFormatHelper);
		}

		#endregion ITableBatchHelper 实现

		#region 内部实现

		/// <summary>
		/// 格式化表格的内部实现（同步）
		/// </summary>
		private void FormatTablesInternal(NETOP.Application netApp,ITableFormatHelper tableFormatHelper)
		{
			_logger.LogInformation($"启动，netApp类型={netApp?.GetType().Name??"null"}");
			if(tableFormatHelper==null)
				throw new InvalidOperationException("无法获取 ITableFormatHelper 服务");

			ExHandler.Run(() =>
			{
				var currentApp = netApp;
				if(!TryRefreshContext(ref currentApp))
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var selection = GetSelectionWithRetry(ref currentApp);

				// 调试：记录选中对象信息
				if(selection==null)
				{
					_logger.LogWarning("ValidateSelection 返回 null，没有选中对象");
					Toast.Show(ResourceManager.GetString("Toast_FormatTables_NoSelection"),Toast.ToastType.Warning);
					return;
				}

				UndoHelper.BeginUndoEntry(currentApp,UndoHelper.UndoNames.FormatTables);

				// 调试：记录选中对象的数量和类型
				try
				{
					if(selection is NETOP.ShapeRange shapeRange)
					{
						int count = ExHandler.SafeGet(() => shapeRange.Count, defaultValue: 0);
						_logger.LogInformation($"选中对象类型=ShapeRange, 数量={count}");
					} else if(selection is NETOP.Shape shape)
					{
						_logger.LogInformation("选中对象类型=Shape, 数量=1");
					} else
					{
						_logger.LogInformation($"选中对象类型={selection?.GetType().Name??"null"}");
					}
				} catch(Exception ex)
				{
					_logger.LogWarning($"获取选中对象信息失败: {ex.Message}");
				}

				var tableShapes = new List<(NETOP.Shape shape, NETOP.Table table)>();

				// 收集表格形状（只处理选中的对象）
				CollectTablesFromSelection(selection,currentApp,tableShapes);

				// 处理收集到的表格
				ProcessTables(tableShapes,currentApp,tableFormatHelper,selection);
			},enableTiming: true);
		}

		/// <summary>
		/// 检查形状是否是表格
		/// </summary>
		/// <param name="shape"> 要检查的形状 </param>
		/// <returns> 如果是表格返回 true，否则返回 false </returns>
		private bool IsTableShape(NETOP.Shape shape)
		{
			if(shape==null) return false;

			bool hasTable = ExHandler.SafeGet(() => shape.HasTable == MsoTriState.msoTrue, defaultValue: false);
			if(hasTable) return true;

			var table = ExHandler.SafeGet(() => shape.Table, defaultValue: (NETOP.Table)null);
			return table!=null;
		}

		/// <summary>
		/// 从形状获取表格对象
		/// </summary>
		/// <param name="shape"> 形状对象 </param>
		/// <returns> 表格对象，如果不是表格则返回 null </returns>
		private NETOP.Table GetTableFromShape(NETOP.Shape shape)
		{
			if(shape==null) return null;

			bool hasTable = ExHandler.SafeGet(() => shape.HasTable == MsoTriState.msoTrue, defaultValue: false);
			if(hasTable)
			{
				return ExHandler.SafeGet(() => shape.Table,defaultValue: (NETOP.Table) null);
			}

			// HasTable 不可用，尝试直接检查 Table 属性
			return ExHandler.SafeGet(() => shape.Table,defaultValue: (NETOP.Table) null);
		}

		private void CollectTablesFromSelection(dynamic selection,NETOP.Application netApp,List<(NETOP.Shape shape, NETOP.Table table)> tableShapes)
		{
			var processedKeys = new HashSet<object>();

			if(selection is NETOP.ShapeRange shapeRange)
			{
				foreach(NETOP.Shape shape in shapeRange)
				{
					AddTableShapeIfValid(shape,tableShapes,processedKeys);
				}
			} else if(selection is NETOP.Shape shape)
			{
				AddTableShapeIfValid(shape,tableShapes,processedKeys);
			}
		}

		/// <summary>
		/// 如果形状是表格，则添加到列表
		/// </summary>
		/// <param name="shape"> 形状对象 </param>
		/// <param name="tableShapes"> 表格形状列表 </param>
		/// <param name="processedKeys"> 已处理的对象键列表（用于去重） </param>
		/// <returns> 如果成功添加返回 true，否则返回 false </returns>
		private bool AddTableShapeIfValid(NETOP.Shape shape,List<(NETOP.Shape shape, NETOP.Table table)> tableShapes,HashSet<object> processedKeys)
		{
			if(shape==null) return false;
			// 检查是否已处理
			if(processedKeys.Contains(shape)) return false;

			var table = GetTableFromShape(shape);
			if(table!=null)
			{
				processedKeys.Add(shape);
				tableShapes.Add((shape, table));
				return true;
			}
			return false;
		}

		/// <summary>
		/// 处理收集到的表格
		/// </summary>
		/// <param name="tableShapes"> 表格形状列表 </param>
		/// <param name="netApp"> PowerPoint 应用程序对象 </param>
		/// <param name="tableFormatHelper"> 表格格式化辅助类 </param>
		/// <param name="selection"> 选区对象 </param>
		private void ProcessTables(List<(NETOP.Shape shape, NETOP.Table table)> tableShapes,NETOP.Application netApp,ITableFormatHelper tableFormatHelper,dynamic selection)
		{
			int totalTables = tableShapes.Count;
			_logger.LogInformation($"找到 {totalTables} 个表格形状");

			if(totalTables==0)
			{
				Toast.Show(ResourceManager.GetString("Toast_FormatTables_NoSelection"),Toast.ToastType.Warning);
				return;
			}

			try
			{
				// 尝试使用新架构服务
				var newService = LegacyServiceBridge.GetService<ITableFormatService>();
				bool useNewService = newService != null;
				
				if (useNewService)
				{
					_logger.LogInformation("使用新架构 TableFormatService");
					// 从旧配置构建格式化选项
					var options = BuildFormatOptionsFromConfig();
					
					foreach(var (shape, table) in tableShapes)
					{
						if(table!=null)
						{
							var tableContext = new PowerPointTableContext(table);
							newService.FormatTable(tableContext, options);
						}
					}
				}
				else
				{
					_logger.LogInformation("使用旧实现 TableFormatHelper");
					foreach(var (shape, table) in tableShapes)
					{
						if(table!=null)
						{
							tableFormatHelper.FormatTables(table);
						}
					}
				}

				Toast.Show(ResourceManager.GetString("Toast_FormatTables_Success",totalTables),Toast.ToastType.Success);
			} finally
			{
				// 释放所有收集的 Shape 和 Table 对象
				foreach(var (shape, table) in tableShapes)
				{
					shape?.Dispose();
					table?.Dispose();
				}
			}
		}

		private bool TryRefreshContext(ref NETOP.Application netApp)
		{
			netApp=ApplicationHelper.EnsureValidNetApplication(netApp);
			return netApp!=null;
		}

		private dynamic GetSelectionWithRetry(ref NETOP.Application netApp)
		{
			var selection = _shapeHelper.ValidateSelection(netApp, showWarningWhenInvalid: false);
			if(selection!=null)
			{
				return selection;
			}

			_logger.LogWarning("返回 null，尝试刷新 Application 后重试");
			if(!TryRefreshContext(ref netApp))
			{
				return null;
			}

			return _shapeHelper.ValidateSelection(netApp,showWarningWhenInvalid: false);
		}

		/// <summary>
		/// 从旧配置构建新架构的格式化选项
		/// </summary>
		private static TableFormatOptions BuildFormatOptionsFromConfig()
		{
			var config = FormattingConfig.Instance?.Table;
			if (config == null)
			{
				return new TableFormatOptions();
			}

			return new TableFormatOptions
			{
				FormatHeader = true,
				FormatDataRows = true,
				ApplyBorders = true,
				ApplyFont = true,
				ApplyTableStyle = true,
				TableStyleId = config.StyleId,
				AutoNumberFormat = config.AutoNumberFormat,
				DecimalPlaces = config.DecimalPlaces,
				NegativeTextColor = config.NegativeTextColor,
				Settings = new TableSettings
				{
					FirstRow = config.TableSettings?.FirstRow ?? true,
					FirstCol = config.TableSettings?.FirstCol ?? false,
					LastRow = config.TableSettings?.LastRow ?? false,
					LastCol = config.TableSettings?.LastCol ?? false,
					HorizBanding = config.TableSettings?.HorizBanding ?? false,
					VertBanding = config.TableSettings?.VertBanding ?? false
				},
				HeaderStyle = new RowStyle
				{
					HideBackground = true,
					FontName = config.HeaderRowFont?.Name ?? "+mn-lt",
					FontNameFarEast = config.HeaderRowFont?.NameFarEast ?? "+mn-ea",
					FontSize = config.HeaderRowFont?.Size ?? 10f,
					Bold = config.HeaderRowFont?.Bold ?? true,
					ThemeColorIndex = ConfigHelper.GetThemeColorIndexValue(config.HeaderRowFont?.ThemeColor),
					Alignment = TextAlignment.Center,
					TopBorder = BorderStyle.SolidTheme(
						ConfigHelper.GetThemeColorIndexValue(config.HeaderRowBorderColor) ?? 5,
						config.HeaderRowBorderWidth
					),
					BottomBorder = BorderStyle.SolidTheme(
						ConfigHelper.GetThemeColorIndexValue(config.HeaderRowBorderColor) ?? 5,
						config.HeaderRowBorderWidth
					)
				},
				DataRowStyle = new RowStyle
				{
					HideBackground = true,
					FontName = config.DataRowFont?.Name ?? "+mn-lt",
					FontNameFarEast = config.DataRowFont?.NameFarEast ?? "+mn-ea",
					FontSize = config.DataRowFont?.Size ?? 9f,
					Bold = config.DataRowFont?.Bold ?? false,
					ThemeColorIndex = ConfigHelper.GetThemeColorIndexValue(config.DataRowFont?.ThemeColor),
					TopBorder = BorderStyle.SolidTheme(
						ConfigHelper.GetThemeColorIndexValue(config.DataRowBorderColor) ?? 6,
						config.DataRowBorderWidth
					)
				}
			};
		}

		#endregion 内部实现
	}
}
