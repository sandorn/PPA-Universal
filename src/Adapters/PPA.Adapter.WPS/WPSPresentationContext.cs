using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 演示文稿上下文实现
    /// </summary>
    public class WPSPresentationContext : IPresentationContext
    {
        private readonly dynamic _presentation;

        public WPSPresentationContext(dynamic presentation)
        {
            _presentation = presentation;
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
