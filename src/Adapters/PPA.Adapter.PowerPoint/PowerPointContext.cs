using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
	/// <summary>
	/// PowerPoint 应用程序上下文实现
	/// </summary>
	public class PowerPointContext : IApplicationContext
	{
		private readonly NETOP.Application _netApp;
		private readonly object _nativeApp;
		private readonly float _slideWidthFallback;
		private readonly float _slideHeightFallback;

		public PowerPointContext(NETOP.Application netApp, object nativeApp = null,
			float slideWidthFallback = PpaConfigTemplateFallbacks.SlideWidthFallback,
			float slideHeightFallback = PpaConfigTemplateFallbacks.SlideHeightFallback)
		{
			_netApp = netApp;
			_nativeApp = nativeApp;
			_slideWidthFallback = slideWidthFallback > 0 ? slideWidthFallback : PpaConfigTemplateFallbacks.SlideWidthFallback;
			_slideHeightFallback = slideHeightFallback > 0 ? slideHeightFallback : PpaConfigTemplateFallbacks.SlideHeightFallback;
		}

		public PlatformType Platform => PlatformType.PowerPoint;

		public IPresentationContext ActivePresentation
		{
			get
			{
				try
				{
					var pres = _netApp?.ActivePresentation;
					return pres != null ? new PowerPointPresentationContext(pres, _slideWidthFallback, _slideHeightFallback) : null;
				}
				catch
				{
					return null;
				}
			}
		}

		public IWindowContext ActiveWindow
		{
			get
			{
				try
				{
					var window = _netApp?.ActiveWindow;
					return window != null ? new PowerPointWindowContext(window) : null;
				}
				catch
				{
					return null;
				}
			}
		}

		public ISelectionContext Selection
		{
			get
			{
				try
				{
					var selection = _netApp?.ActiveWindow?.Selection;
					return selection != null ? new PowerPointSelectionContext(selection) : null;
				}
				catch
				{
					return null;
				}
			}
		}

		public bool IsFeatureSupported(Feature feature)
		{
			// PowerPoint 支持所有功能
			switch (feature)
			{
				case Feature.TableBasic:
				case Feature.TableAdvancedBorder:
				case Feature.Chart:
				case Feature.ChartAdvanced:
				case Feature.ShapeAlignment:
				case Feature.ShapeBatch:
				case Feature.TextAdvanced:
				case Feature.UndoRedo:
					return true;
				default:
					return false;
			}
		}

		public object NativeApplication => _nativeApp ?? (object)_netApp;

		/// <summary>
		/// 获取 NetOffice Application 实例
		/// </summary>
		public NETOP.Application NetApplication => _netApp;

		/// <summary>
		/// 获取原生 COM Application 实例
		/// </summary>
		public object NativeApp => _nativeApp;
	}
}
