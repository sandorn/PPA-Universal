using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 形状操作实现
    /// </summary>
    public class WPSShapeOps : IShapeOperations
    {
        public ShapeRect GetBounds(object shape)
        {
            dynamic dynShape = shape;
            if (dynShape == null) return new ShapeRect();

            try
            {
                return new ShapeRect(
                    (float)(dynShape.Left ?? 0),
                    (float)(dynShape.Top ?? 0),
                    (float)(dynShape.Width ?? 0),
                    (float)(dynShape.Height ?? 0)
                );
            }
            catch
            {
                return new ShapeRect();
            }
        }

        public void SetBounds(object shape, ShapeRect bounds)
        {
            dynamic dynShape = shape;
            if (dynShape == null) return;

            try
            {
                dynShape.Left = bounds.Left;
                dynShape.Top = bounds.Top;
                dynShape.Width = bounds.Width;
                dynShape.Height = bounds.Height;
            }
            catch { }
        }

        public float GetRotation(object shape)
        {
            dynamic dynShape = shape;
            try { return (float)(dynShape?.Rotation ?? 0); }
            catch { return 0; }
        }

        public void SetRotation(object shape, float angle)
        {
            dynamic dynShape = shape;
            try { if (dynShape != null) dynShape.Rotation = angle; }
            catch { }
        }

        public bool IsTable(object shape)
        {
            dynamic dynShape = shape;
            if (dynShape == null) return false;

            try
            {
                int hasTable = dynShape.HasTable ?? 0;
                return hasTable == WPSHelper.TriState.True;
            }
            catch
            {
                return false;
            }
        }

        public bool IsChart(object shape)
        {
            dynamic dynShape = shape;
            if (dynShape == null) return false;

            try
            {
                int hasChart = dynShape.HasChart ?? 0;
                return hasChart == WPSHelper.TriState.True;
            }
            catch
            {
                return false;
            }
        }

        public bool IsTextBox(object shape)
        {
            dynamic dynShape = shape;
            if (dynShape == null) return false;

            try
            {
                int type = dynShape.Type ?? 0;
                return type == 17; // msoTextBox
            }
            catch
            {
                return false;
            }
        }

        public bool IsGroup(object shape)
        {
            dynamic dynShape = shape;
            if (dynShape == null) return false;

            try
            {
                int type = dynShape.Type ?? 0;
                return type == 6; // msoGroup
            }
            catch
            {
                return false;
            }
        }

        public object CopyShape(object shape)
        {
            dynamic dynShape = shape;
            if (dynShape == null) return null;

            try
            {
                dynShape.Copy();
                dynamic slide = dynShape.Parent;
                slide?.Shapes.Paste();
                int count = slide?.Shapes.Count ?? 0;
                return count > 0 ? slide.Shapes[count] : null;
            }
            catch
            {
                return null;
            }
        }

        public void DeleteShape(object shape)
        {
            dynamic dynShape = shape;
            try { dynShape?.Delete(); }
            catch { }
        }
    }
}
