using NetOffice.OfficeApi.Enums;
using PPA.Business.Abstractions;
using PPA.Core;
using PPA.Core.Abstraction;
using PPA.Core.Abstraction.Business;
using PPA.Core.Logging;
using PPA.Legacy.Bridge;
using PPA.Logging;
using PPA.Manipulation.Config;
using PPA.Universal.Platform;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Manipulation
{
	/// <summary>
	/// 表格批量操作辅助类 提供表格批量格式化功能，支持同步和异步操作
	/// </summary>
	internal class TableBatchHelper(IShapeHelper shapeHelper,ILogger logger = null):ITableBatchHelper
	{
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
			FormatTablesInternal(netApp);
		}

		#endregion ITableBatchHelper 实现

		#region 内部实现

		/// <summary>
		/// 格式化表格的内部实现（同步）
		/// </summary>
		private void FormatTablesInternal(NETOP.Application netApp)
		{
			_logger.LogInformation($"启动，netApp类型={netApp?.GetType().Name??"null"}");
			
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
				ProcessTables(tableShapes,currentApp,selection);
			},enableTiming: true);
		}

        // ... (中间代码保持不变) ...

		/// <summary>
		/// 处理收集到的表格
		/// </summary>
		/// <param name="tableShapes"> 表格形状列表 </param>
		/// <param name="netApp"> PowerPoint 应用程序对象 </param>
		/// <param name="selection"> 选区对象 </param>
		private void ProcessTables(List<(NETOP.Shape shape, NETOP.Table table)> tableShapes,NETOP.Application netApp,dynamic selection)
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
				// 使用新架构服务
				var newService = LegacyServiceBridge.GetService<ITableFormatService>();
				if (newService == null)
                {
                    _logger.LogError("无法获取 ITableFormatService 服务，格式化中止");
                    Toast.Show("内部错误：服务未注册", Toast.ToastType.Error);
                    return;
                }
				
                // 检测平台并创建适配器工厂
                var platform = PlatformDetector.Detect().ActivePlatform;
                var adapterFactory = new AdapterFactory(_logger);
                _logger.LogInformation($"使用新架构 TableFormatService，平台: {platform}");
                
                // 从旧配置构建格式化选项
                var options = BuildFormatOptionsFromConfig();
                
                foreach(var (shape, table) in tableShapes)
                {
                    if(table!=null)
                    {
                        // --- 即使是新架构，也先尝试使用 ExecuteMso 清除样式（针对 WPS 优化） ---
                        try
                        {
                            shape.Select();
                            _logger.LogInformation("执行 ExecuteMso('TableStyleClearTable')");
                            netApp.CommandBars.ExecuteMso("TableStyleClearTable");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning($"执行 TableStyleClearTable 失败: {ex.Message}");
                        }

                        var tableContext = adapterFactory.CreateTableContext(table, platform);
                        newService.FormatTable(tableContext, options);
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
		/// 从选区中收集表格形状
		/// </summary>
		private void CollectTablesFromSelection(dynamic selection, NETOP.Application netApp, List<(NETOP.Shape shape, NETOP.Table table)> tableShapes)
		{
			if (selection == null) return;

			try
			{
				// 如果是 ShapeRange
				if (selection is NETOP.ShapeRange shapeRange)
				{
					foreach (NETOP.Shape shape in shapeRange)
					{
						if (shape.HasTable == MsoTriState.msoTrue)
						{
							tableShapes.Add((shape, shape.Table));
						}
					}
				}
				// 如果是单个 Shape
				else if (selection is NETOP.Shape shape)
				{
					if (shape.HasTable == MsoTriState.msoTrue)
					{
						tableShapes.Add((shape, shape.Table));
					}
				}
				// 如果是 Selection 对象（通常不会直接传 selection 对象进来，而是传 ValidateSelection 返回的 ShapeRange 或 Shape，但为了保险起见）
				else if (selection is NETOP.Selection sel)
				{
					if (sel.Type == NETOP.Enums.PpSelectionType.ppSelectionShapes && sel.ShapeRange != null)
					{
						foreach (NETOP.Shape subShape in sel.ShapeRange)
						{
							if (subShape.HasTable == MsoTriState.msoTrue)
							{
								tableShapes.Add((subShape, subShape.Table));
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				_logger.LogWarning($"收集表格形状时发生错误: {ex.Message}");
			}
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
