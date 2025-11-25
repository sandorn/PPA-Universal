using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 形状操作实现
    /// </summary>
    public class PowerPointShapeOps : IShapeOperations
    {
        public ShapeRect GetBounds(object shape)
        {
            var netShape = shape as NETOP.Shape;
            if (netShape == null) return new ShapeRect();

            return new ShapeRect(
                netShape.Left,
                netShape.Top,
                netShape.Width,
                netShape.Height
            );
        }

        public void SetBounds(object shape, ShapeRect bounds)
        {
            var netShape = shape as NETOP.Shape;
            if (netShape == null) return;

            netShape.Left = bounds.Left;
            netShape.Top = bounds.Top;
            netShape.Width = bounds.Width;
            netShape.Height = bounds.Height;
        }

        public float GetRotation(object shape)
        {
            var netShape = shape as NETOP.Shape;
            return netShape?.Rotation ?? 0;
        }

        public void SetRotation(object shape, float angle)
        {
            var netShape = shape as NETOP.Shape;
            if (netShape != null)
            {
                netShape.Rotation = angle;
            }
        }

        public bool IsTable(object shape)
        {
            var netShape = shape as NETOP.Shape;
            if (netShape == null) return false;

            try
            {
                return netShape.HasTable == NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
            }
            catch
            {
                return false;
            }
        }

        public bool IsChart(object shape)
        {
            var netShape = shape as NETOP.Shape;
            if (netShape == null) return false;

            try
            {
                return netShape.HasChart == NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
            }
            catch
            {
                return false;
            }
        }

        public bool IsTextBox(object shape)
        {
            var netShape = shape as NETOP.Shape;
            if (netShape == null) return false;

            try
            {
                return netShape.Type == NetOffice.OfficeApi.Enums.MsoShapeType.msoTextBox;
            }
            catch
            {
                return false;
            }
        }

        public bool IsGroup(object shape)
        {
            var netShape = shape as NETOP.Shape;
            if (netShape == null) return false;

            try
            {
                return netShape.Type == NetOffice.OfficeApi.Enums.MsoShapeType.msoGroup;
            }
            catch
            {
                return false;
            }
        }

        public object CopyShape(object shape)
        {
            var netShape = shape as NETOP.Shape;
            if (netShape == null) return null;

            try
            {
                netShape.Copy();
                var slide = netShape.Parent as NETOP.Slide;
                slide?.Shapes.Paste();
                return slide?.Shapes[slide.Shapes.Count];
            }
            catch
            {
                return null;
            }
        }

        public void DeleteShape(object shape)
        {
            var netShape = shape as NETOP.Shape;
            try
            {
                netShape?.Delete();
            }
            catch { }
        }
    }
}
