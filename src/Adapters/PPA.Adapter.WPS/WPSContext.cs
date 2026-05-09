using PPA.Core.Abstraction;
using PPA.Core.Configuration;

namespace PPA.Adapter.WPS
{
	/// <summary>
	/// WPS 应用程序上下文实现
	/// </summary>
	public class WPSContext : IApplicationContext
	{
		private readonly dynamic _app;
		private readonly float _slideWidthFallback;
		private readonly float _slideHeightFallback;

		public WPSContext(dynamic app,
			float slideWidthFallback = PpaConfigTemplateFallbacks.SlideWidthFallback,
			float slideHeightFallback = PpaConfigTemplateFallbacks.SlideHeightFallback)
		{
			_app = app;
			_slideWidthFallback = slideWidthFallback > 0 ? slideWidthFallback : PpaConfigTemplateFallbacks.SlideWidthFallback;
			_slideHeightFallback = slideHeightFallback > 0 ? slideHeightFallback : PpaConfigTemplateFallbacks.SlideHeightFallback;
		}

		public PlatformType Platform => PlatformType.WPS;

		public IPresentationContext ActivePresentation
		{
			get
			{
				try
				{
					dynamic pres = _app?.ActivePresentation;
					return pres != null ? new WPSPresentationContext(pres, _slideWidthFallback, _slideHeightFallback) : null;
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
					dynamic window = _app?.ActiveWindow;
					return window != null ? new WPSWindowContext(window) : null;
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
					dynamic selection = _app?.ActiveWindow?.Selection;
					return selection != null ? new WPSSelectionContext(selection) : null;
				}
				catch
				{
					return null;
				}
			}
		}

		public bool IsFeatureSupported(Feature feature)
		{
			// WPS 支持大部分功能，但某些高级功能可能受限
			switch (feature)
			{
				case Feature.TableBasic:
				case Feature.ShapeAlignment:
				case Feature.ShapeBatch:
				case Feature.UndoRedo:
					return true;

				case Feature.TableAdvancedBorder:
				case Feature.Chart:
				case Feature.ChartAdvanced:
				case Feature.TextAdvanced:
					return true; // 基本支持，但可能有差异

				default:
					return false;
			}
		}

		public object NativeApplication => _app;

		/// <summary>
		/// 获取 WPS Application 动态对象
		/// </summary>
		public dynamic Application => _app;

		/// <summary>
		/// 获取应用程序名称
		/// </summary>
		public string ApplicationName
		{
			get
			{
				try { return _app?.Name; }
				catch { return "WPS 演示"; }
			}
		}

		/// <summary>
		/// 获取应用程序版本
		/// </summary>
		public string Version
		{
			get
			{
				try { return _app?.Version; }
				catch { return "Unknown"; }
			}
		}
	}
}
