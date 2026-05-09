using PPA.Core.Abstraction;
using PPA.Core.Configuration;

namespace PPA.Adapter.WPS
{
	/// <summary>
	/// WPS 演示文稿上下文实现
	/// </summary>
	public class WPSPresentationContext : IPresentationContext
	{
		private readonly dynamic _presentation;
		private readonly float _slideWidthFallback;
		private readonly float _slideHeightFallback;

		public WPSPresentationContext(dynamic presentation,
			float slideWidthFallback = PpaConfigTemplateFallbacks.SlideWidthFallback,
			float slideHeightFallback = PpaConfigTemplateFallbacks.SlideHeightFallback)
		{
			_presentation = presentation;
			_slideWidthFallback = slideWidthFallback > 0 ? slideWidthFallback : PpaConfigTemplateFallbacks.SlideWidthFallback;
			_slideHeightFallback = slideHeightFallback > 0 ? slideHeightFallback : PpaConfigTemplateFallbacks.SlideHeightFallback;
		}

		public string Name
		{
			get
			{
				try { return _presentation?.Name ?? string.Empty; }
				catch { return string.Empty; }
			}
		}

		public string FullName
		{
			get
			{
				try { return _presentation?.FullName ?? string.Empty; }
				catch { return string.Empty; }
			}
		}

		public int SlideCount
		{
			get
			{
				try { return _presentation?.Slides?.Count ?? 0; }
				catch { return 0; }
			}
		}

		public float SlideWidth
		{
			get
			{
				try { return _presentation?.PageSetup?.SlideWidth ?? _slideWidthFallback; }
				catch { return _slideWidthFallback; }
			}
		}

		public float SlideHeight
		{
			get
			{
				try { return _presentation?.PageSetup?.SlideHeight ?? _slideHeightFallback; }
				catch { return _slideHeightFallback; }
			}
		}

		public ISlideContext GetSlide(int index)
		{
			try
			{
				if (index < 1 || index > SlideCount) return null;
				dynamic slide = _presentation.Slides[index];
				return slide != null ? new WPSSlideContext(slide) : null;
			}
			catch
			{
				return null;
			}
		}

		public object NativePresentation => _presentation;
	}
}
