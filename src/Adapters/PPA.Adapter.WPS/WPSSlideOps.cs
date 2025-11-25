using System.Collections.Generic;
using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 幻灯片操作实现
    /// </summary>
    public class WPSSlideOps : ISlideOperations
    {
        public IEnumerable<object> GetShapes(object slide)
        {
            dynamic dynSlide = slide;
            if (dynSlide?.Shapes == null) yield break;

            int count = 0;
            try { count = dynSlide.Shapes.Count; }
            catch { yield break; }

            for (int i = 1; i <= count; i++)
            {
                object shape = null;
                try { shape = dynSlide.Shapes[i]; }
                catch { continue; }

                if (shape != null)
                {
                    yield return shape;
                }
            }
        }

        public int GetShapeCount(object slide)
        {
            dynamic dynSlide = slide;
            try { return dynSlide?.Shapes?.Count ?? 0; }
            catch { return 0; }
        }

        public object AddShape(object slide, ShapeType type, ShapeRect bounds)
        {
            dynamic dynSlide = slide;
            if (dynSlide == null) return null;

            try
            {
                int msoType = ConvertShapeType(type);
                return dynSlide.Shapes.AddShape(msoType, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
            }
            catch
            {
                return null;
            }
        }

        public object AddTable(object slide, int rows, int columns, ShapeRect bounds)
        {
            dynamic dynSlide = slide;
            if (dynSlide == null) return null;

            try
            {
                return dynSlide.Shapes.AddTable(rows, columns, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
            }
            catch
            {
                return null;
            }
        }

        public object DuplicateSlide(object slide)
        {
            dynamic dynSlide = slide;
            if (dynSlide == null) return null;

            try
            {
                dynamic duplicated = dynSlide.Duplicate();
                return duplicated?[1];
            }
            catch
            {
                return null;
            }
        }

        public void DeleteSlide(object slide)
        {
            dynamic dynSlide = slide;
            try { dynSlide?.Delete(); }
            catch { }
        }

        public int GetSlideIndex(object slide)
        {
            dynamic dynSlide = slide;
            try { return dynSlide?.SlideIndex ?? 0; }
            catch { return 0; }
        }

        public void MoveSlide(object slide, int newIndex)
        {
            dynamic dynSlide = slide;
            try { dynSlide?.MoveTo(newIndex); }
            catch { }
        }

        private int ConvertShapeType(ShapeType type)
        {
            switch (type)
            {
                case ShapeType.AutoShape:
                case ShapeType.TextBox:
                    return 1; // msoShapeRectangle
                default:
                    return 1;
            }
        }
    }
}
