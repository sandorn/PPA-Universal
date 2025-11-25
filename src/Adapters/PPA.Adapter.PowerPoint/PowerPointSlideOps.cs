using System.Collections.Generic;
using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 幻灯片操作实现
    /// </summary>
    public class PowerPointSlideOps : ISlideOperations
    {
        public IEnumerable<object> GetShapes(object slide)
        {
            var netSlide = slide as NETOP.Slide;
            if (netSlide?.Shapes == null) yield break;

            foreach (NETOP.Shape shape in netSlide.Shapes)
            {
                yield return shape;
            }
        }

        public int GetShapeCount(object slide)
        {
            var netSlide = slide as NETOP.Slide;
            return netSlide?.Shapes?.Count ?? 0;
        }

        public object AddShape(object slide, ShapeType type, ShapeRect bounds)
        {
            var netSlide = slide as NETOP.Slide;
            if (netSlide == null) return null;

            try
            {
                var msoType = ConvertShapeType(type);
                return netSlide.Shapes.AddShape(msoType, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
            }
            catch
            {
                return null;
            }
        }

        public object AddTable(object slide, int rows, int columns, ShapeRect bounds)
        {
            var netSlide = slide as NETOP.Slide;
            if (netSlide == null) return null;

            try
            {
                return netSlide.Shapes.AddTable(rows, columns, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
            }
            catch
            {
                return null;
            }
        }

        public object DuplicateSlide(object slide)
        {
            var netSlide = slide as NETOP.Slide;
            if (netSlide == null) return null;

            try
            {
                return netSlide.Duplicate()[1];
            }
            catch
            {
                return null;
            }
        }

        public void DeleteSlide(object slide)
        {
            var netSlide = slide as NETOP.Slide;
            try
            {
                netSlide?.Delete();
            }
            catch { }
        }

        public int GetSlideIndex(object slide)
        {
            var netSlide = slide as NETOP.Slide;
            return netSlide?.SlideIndex ?? 0;
        }

        public void MoveSlide(object slide, int newIndex)
        {
            var netSlide = slide as NETOP.Slide;
            if (netSlide == null) return;

            try
            {
                netSlide.MoveTo(newIndex);
            }
            catch { }
        }

        private NetOffice.OfficeApi.Enums.MsoAutoShapeType ConvertShapeType(ShapeType type)
        {
            switch (type)
            {
                case ShapeType.AutoShape:
                    return NetOffice.OfficeApi.Enums.MsoAutoShapeType.msoShapeRectangle;
                case ShapeType.TextBox:
                    return NetOffice.OfficeApi.Enums.MsoAutoShapeType.msoShapeRectangle;
                default:
                    return NetOffice.OfficeApi.Enums.MsoAutoShapeType.msoShapeRectangle;
            }
        }
    }
}
