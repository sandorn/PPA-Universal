using System;
using System.Linq;
using System.Reflection;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
	/// <summary>
	/// 撤销/重做服务实现
	/// 
	/// 【工作原理】
	/// PowerPoint: 连续的 COM 操作会被自动合并为一个撤销条目。
	/// 调用 StartNewUndoEntry() 创建撤销边界，使每个操作成为独立的撤销条目。
	/// 
	/// WPS: 在 COM 对象上反射尝试 BeginUndoGroup/StartUndoGroup 与对应 End*；若自动化未暴露则无法合并为单条撤销。
	/// 使用 CreateUndoScope() 的 Dispose 结束撤销组（若成功开始）。
	/// </summary>
	public class UndoService : IUndoService
	{
		private readonly ILogger _logger;

		public UndoService(ILogger logger)
		{
			_logger = logger ?? NullLogger.Instance;
		}

		public bool StartNewUndoEntry(IApplicationContext context, string undoEntryName)
		{
			// 此方法仅用于 PowerPoint
			// WPS 应使用 CreateUndoScope() 方法
			if (context == null || context.Platform != PlatformType.PowerPoint)
			{
				return false;
			}

			try
			{
				var nativeApp = context.NativeApplication;
				if (nativeApp == null) return false;

				return TryStartNewUndoEntryPowerPoint(nativeApp, undoEntryName);
			}
			catch (Exception ex)
			{
				_logger.LogDebug($"设置撤销边界失败: {ex.Message}");
				return false;
			}
		}

		public void EndUndoEntry(IApplicationContext context)
		{
			// PowerPoint 不需要显式结束
			// WPS 应使用 CreateUndoScope() 的 Dispose 来结束
		}

		public IDisposable CreateUndoScope(IApplicationContext context, string undoEntryName)
		{
			if (context == null)
			{
				return new NullUndoScope();
			}

			switch (context.Platform)
			{
				case PlatformType.PowerPoint:
					// PowerPoint: 调用 StartNewUndoEntry 创建边界
					TryStartNewUndoEntryPowerPoint(context.NativeApplication, undoEntryName);
					return new NullUndoScope(); // PowerPoint 不需要结束调用

				case PlatformType.WPS:
					{
						var wps = TryCreateWpsUndoScope(context, undoEntryName);
						return wps ?? new NullUndoScope();
					}

				default:
					return new NullUndoScope();
			}
		}

		public bool IsSupported(IApplicationContext context)
		{
			if (context == null) return false;
			return context.Platform == PlatformType.PowerPoint ||
				   context.Platform == PlatformType.WPS;
		}

		private bool TryStartNewUndoEntryPowerPoint(object nativeApp, string undoEntryName)
		{
			if (nativeApp == null) return false;

			// 尝试通过 dynamic 调用
			try
			{
				dynamic dynamicApp = nativeApp;
				dynamicApp.StartNewUndoEntry();
				_logger.LogDebug($"已设置撤销边界 (PowerPoint): {undoEntryName}");
				return true;
			}
			catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
			{
				// dynamic 调用失败，尝试反射
			}

			// 尝试通过反射调用
			var appType = nativeApp.GetType();
			var method = appType.GetMethod("StartNewUndoEntry",
				BindingFlags.Public | BindingFlags.Instance,
				null,
				Type.EmptyTypes,
				null);

			if (method != null)
			{
				method.Invoke(nativeApp, null);
				_logger.LogDebug($"已设置撤销边界 (PowerPoint, 反射): {undoEntryName}");
				return true;
			}

			_logger.LogDebug("PowerPoint 未找到 StartNewUndoEntry 方法");
			return false;
		}

		/// <summary>
		/// WPS：尝试在应用程序或演示文稿对象上通过反射调用 BeginUndoGroup/EndUndoGroup（若自动化暴露）。
		/// 若不存在或未生效，则退回 <see cref="NullUndoScope"/>（撤销仍为多条，与历史版本一致）。
		/// </summary>
		private IDisposable TryCreateWpsUndoScope(IApplicationContext context, string undoEntryName)
		{
			object[] targets =
			{
				context.NativeApplication,
				context.ActivePresentation?.NativePresentation
			};

			foreach (var target in targets.Where(t => t != null))
			{
				var scope = TryInvokeWpsUndoGroupPair(target, undoEntryName);
				if (scope != null)
					return scope;
			}

			_logger.LogDebug("WPS 未检测到可用的撤销合并 API，单次操作仍可能产生多条撤销记录");
			return null;
		}

		private IDisposable TryInvokeWpsUndoGroupPair(object target, string undoEntryName)
		{
			var t = target.GetType();
			foreach (var beginName in new[] { "BeginUndoGroup", "StartUndoGroup" })
			{
				MethodInfo begin = t.GetMethod(beginName, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
				object[] beginArgs = null;
				if (begin == null)
				{
					begin = t.GetMethod(beginName, BindingFlags.Public | BindingFlags.Instance, null, new[] { typeof(string) }, null);
					beginArgs = new object[] { undoEntryName ?? string.Empty };
				}

				if (begin == null)
					continue;

				var endName = beginName.StartsWith("Begin", StringComparison.Ordinal)
					? "End" + beginName.Substring("Begin".Length)
					: "End" + beginName.Substring("Start".Length);

				var end = t.GetMethod(endName, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
				if (end == null)
					continue;

				try
				{
					begin.Invoke(target, beginArgs);
					_logger.LogDebug($"WPS 撤销合并已开始 ({beginName})");
					return new ReflectionUndoScope(target, end, _logger);
				}
				catch (Exception ex)
				{
					_logger.LogDebug($"WPS 调用 {beginName} 失败: {ex.Message}");
				}
			}

			return null;
		}
	}

	/// <summary>
	/// 通过反射调用 EndUndoGroup 的撤销作用域。
	/// </summary>
	internal sealed class ReflectionUndoScope : IDisposable
	{
		private readonly object _target;
		private readonly MethodInfo _end;
		private readonly ILogger _logger;

		public ReflectionUndoScope(object target, MethodInfo end, ILogger logger)
		{
			_target = target;
			_end = end;
			_logger = logger;
		}

		public void Dispose()
		{
			try
			{
				_end?.Invoke(_target, null);
			}
			catch (Exception ex)
			{
				_logger?.LogDebug($"WPS EndUndoGroup 调用失败: {ex.Message}");
			}
		}
	}

	/// <summary>
	/// 空撤销作用域 - 用于不支持撤销合并的平台
	/// </summary>
	internal class NullUndoScope : IDisposable
	{
		public void Dispose() { }
	}
}

