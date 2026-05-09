using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
	/// <summary>
	/// PowerPoint 演示文稿上下文实现
	/// </summary>
	public class PowerPointPresentationContext : IPresentationContext
	{
		private readonly NETOP.Presentation _presentation;
		private readonly float _slideWidthFallback;
		private readonly float _slideHeightFallback;

		public PowerPointPresentationContext(NETOP.Presentation presentation,
			float slideWidthFallback = PpaConfigTemplateFallbacks.SlideWidthFallback,
			float slideHeightFallback = PpaConfigTemplateFallbacks.SlideHeightFallback)
		{
			_presentation = presentation;
			_slideWidthFallback = slideWidthFallback > 0 ? slideWidthFallback : PpaConfigTemplateFallbacks.SlideWidthFallback;
			_slideHeightFallback = slideHeightFallback > 0 ? slideHeightFallback : PpaConfigTemplateFallbacks.SlideHeightFallback;
		}

		public string Name => _presentation?.Name ?? string.Empty;

		public string FullName => _presentation?.FullName ?? string.Empty;

		public int SlideCount => _presentation?.Slides?.Count ?? 0;

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
				var slide = _presentation.Slides[index];
				return slide != null ? new PowerPointSlideContext(slide) : null;
			}
			catch
			{
				return null;
			}
		}

		public object NativePresentation => _presentation;
	}
}
