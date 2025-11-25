using System.Collections.Generic;
using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 幻灯片上下文实现
    /// </summary>
    public class WPSSlideContext : ISlideContext
    {
        private readonly dynamic _slide;

        public WPSSlideContext(dynamic slide)
        {
            _slide = slide;
        }

        public int SlideIndex
        {
            get
            {
                try { return _slide?.SlideIndex ?? 0; }
                catch { return 0; }
            }
        }

        public int SlideNumber
        {
            get
            {
                try { return _slide?.SlideNumber ?? 0; }
                catch { return 0; }
            }
        }

        public int ShapeCount
        {
            get
            {
                try { return _slide?.Shapes?.Count ?? 0; }
                catch { return 0; }
            }
        }

        public IEnumerable<IShapeContext> Shapes
        {
            get
            {
                if (_slide?.Shapes == null) yield break;

                int count = ShapeCount;
                for (int i = 1; i <= count; i++)
                {
                    dynamic shape = null;
                    try
                    {
                        shape = _slide.Shapes[i];
                    }
                    catch
                    {
                        continue;
                    }

                    if (shape != null)
                    {
                        yield return new WPSShapeContext(shape);
                    }
                }
            }
        }

        public object NativeSlide => _slide;
    }
}
