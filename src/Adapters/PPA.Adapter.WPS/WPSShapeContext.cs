using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 形状上下文实现
    /// </summary>
    public class WPSShapeContext : IShapeContext
    {
        private readonly dynamic _shape;

        public WPSShapeContext(dynamic shape)
        {
            _shape = shape;
        }

        public string Name
        {
            get
            {
                try { return _shape?.Name ?? string.Empty; }
                catch { return string.Empty; }
            }
            set
            {
                try { if (_shape != null) _shape.Name = value; }
                catch { }
            }
        }

        public ShapeType ShapeType
        {
            get
            {
                try
                {
                    int type = _shape?.Type ?? 0;
                    switch (type)
                    {
                        case 1: return ShapeType.AutoShape;
                        case 17: return ShapeType.TextBox;
                        case 19: return ShapeType.Table;
                        case 3: return ShapeType.Chart;
                        case 13: return ShapeType.Picture;
                        case 6: return ShapeType.Group;
                        case 14: return ShapeType.Placeholder;
                        case 16: return ShapeType.Media;
                        default: return ShapeType.Other;
                    }
                }
                catch
                {
                    return ShapeType.Other;
                }
            }
        }

        public ShapeRect Bounds
        {
            get
            {
                try
                {
                    return new ShapeRect(
                        (float)(_shape?.Left ?? 0),
                        (float)(_shape?.Top ?? 0),
                        (float)(_shape?.Width ?? 0),
                        (float)(_shape?.Height ?? 0)
                    );
                }
                catch
                {
                    return new ShapeRect();
                }
            }
            set
            {
                try
                {
                    if (_shape == null) return;
                    _shape.Left = value.Left;
                    _shape.Top = value.Top;
                    _shape.Width = value.Width;
                    _shape.Height = value.Height;
                }
                catch { }
            }
        }

        public float Rotation
        {
            get
            {
                try { return (float)(_shape?.Rotation ?? 0); }
                catch { return 0; }
            }
            set
            {
                try { if (_shape != null) _shape.Rotation = value; }
                catch { }
            }
        }

        public bool IsTable
        {
            get
            {
                try
                {
                    // WPS 使用 HasTable 属性，返回 -1 (True) 或 0 (False)
                    int hasTable = _shape?.HasTable ?? 0;
                    return hasTable == WPSHelper.TriState.True;
                }
                catch
                {
                    return false;
                }
            }
        }

        public bool IsChart
        {
            get
            {
                try
                {
                    int hasChart = _shape?.HasChart ?? 0;
                    return hasChart == WPSHelper.TriState.True;
                }
                catch
                {
                    return false;
                }
            }
        }

        public bool HasTextFrame
        {
            get
            {
                try
                {
                    int hasTextFrame = _shape?.HasTextFrame ?? 0;
                    return hasTextFrame == WPSHelper.TriState.True;
                }
                catch
                {
                    return false;
                }
            }
        }

        public ITableContext Table
        {
            get
            {
                try
                {
                    if (!IsTable) return null;
                    dynamic table = _shape?.Table;
                    return table != null ? new WPSTableContext(table) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public object NativeShape => _shape;

        /// <summary>
        /// 获取 WPS Shape 动态对象
        /// </summary>
        public dynamic Shape => _shape;
    }
}
