using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 演示文稿上下文实现
    /// </summary>
    public class PowerPointPresentationContext : IPresentationContext
    {
        private readonly NETOP.Presentation _presentation;

        public PowerPointPresentationContext(NETOP.Presentation presentation)
        {
            _presentation = presentation;
        }

        public string Name => _presentation?.Name ?? string.Empty;

        public string FullName => _presentation?.FullName ?? string.Empty;

        public int SlideCount => _presentation?.Slides?.Count ?? 0;

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
