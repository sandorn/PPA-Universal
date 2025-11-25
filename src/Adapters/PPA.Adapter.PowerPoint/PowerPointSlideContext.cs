using System.Collections.Generic;
using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 幻灯片上下文实现
    /// </summary>
    public class PowerPointSlideContext : ISlideContext
    {
        private readonly NETOP.Slide _slide;

        public PowerPointSlideContext(NETOP.Slide slide)
        {
            _slide = slide;
        }

        public int SlideIndex => _slide?.SlideIndex ?? 0;

        public int SlideNumber => _slide?.SlideNumber ?? 0;

        public int ShapeCount => _slide?.Shapes?.Count ?? 0;

        public IEnumerable<IShapeContext> Shapes
        {
            get
            {
                if (_slide?.Shapes == null) yield break;

                foreach (NETOP.Shape shape in _slide.Shapes)
                {
                    yield return new PowerPointShapeContext(shape);
                }
            }
        }

        public object NativeSlide => _slide;
    }
}
