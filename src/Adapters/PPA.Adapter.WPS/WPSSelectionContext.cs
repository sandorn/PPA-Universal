using System.Collections.Generic;
using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 选择上下文实现
    /// </summary>
    public class WPSSelectionContext : ISelectionContext
    {
        private readonly dynamic _selection;

        public WPSSelectionContext(dynamic selection)
        {
            _selection = selection;
        }

        public SelectionType Type
        {
            get
            {
                try
                {
                    int selType = _selection?.Type ?? 0;
                    switch (selType)
                    {
                        case 0: // ppSelectionNone
                            return SelectionType.None;
                        case 1: // ppSelectionSlides
                            return SelectionType.Slides;
                        case 2: // ppSelectionShapes
                            return SelectionType.Shapes;
                        case 3: // ppSelectionText
                            return SelectionType.Text;
                        default:
                            return SelectionType.None;
                    }
                }
                catch
                {
                    return SelectionType.None;
                }
            }
        }

        public bool HasSelection => Type != SelectionType.None;

        public int ShapeCount
        {
            get
            {
                try
                {
                    return _selection?.ShapeRange?.Count ?? 0;
                }
                catch
                {
                    return 0;
                }
            }
        }

        public IEnumerable<IShapeContext> SelectedShapes
        {
            get
            {
                dynamic shapeRange = null;
                try
                {
                    shapeRange = _selection?.ShapeRange;
                }
                catch
                {
                    yield break;
                }

                if (shapeRange == null) yield break;

                int count = ShapeCount;
                for (int i = 1; i <= count; i++)
                {
                    dynamic shape = null;
                    try
                    {
                        shape = shapeRange[i];
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

        public object NativeSelection => _selection;

        /// <summary>
        /// 获取原生 ShapeRange
        /// </summary>
        public dynamic ShapeRange
        {
            get
            {
                try { return _selection?.ShapeRange; }
                catch { return null; }
            }
        }
    }
}
