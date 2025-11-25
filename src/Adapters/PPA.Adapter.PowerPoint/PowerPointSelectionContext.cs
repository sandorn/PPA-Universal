using System.Collections.Generic;
using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 选择上下文实现
    /// </summary>
    public class PowerPointSelectionContext : ISelectionContext
    {
        private readonly NETOP.Selection _selection;

        public PowerPointSelectionContext(NETOP.Selection selection)
        {
            _selection = selection;
        }

        public SelectionType Type
        {
            get
            {
                try
                {
                    var selType = _selection?.Type;
                    if (selType == null) return SelectionType.None;

                    switch (selType.Value)
                    {
                        case NETOP.Enums.PpSelectionType.ppSelectionNone:
                            return SelectionType.None;
                        case NETOP.Enums.PpSelectionType.ppSelectionSlides:
                            return SelectionType.Slides;
                        case NETOP.Enums.PpSelectionType.ppSelectionShapes:
                            return SelectionType.Shapes;
                        case NETOP.Enums.PpSelectionType.ppSelectionText:
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
                if (_selection?.ShapeRange == null) yield break;

                try
                {
                    foreach (NETOP.Shape shape in _selection.ShapeRange)
                    {
                        yield return new PowerPointShapeContext(shape);
                    }
                }
                finally { }
            }
        }

        public object NativeSelection => _selection;

        /// <summary>
        /// 获取原生 ShapeRange
        /// </summary>
        public NETOP.ShapeRange ShapeRange => _selection?.ShapeRange;
    }
}
