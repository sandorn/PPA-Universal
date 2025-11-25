using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 形状上下文实现
    /// </summary>
    public class PowerPointShapeContext : IShapeContext
    {
        private readonly NETOP.Shape _shape;

        public PowerPointShapeContext(NETOP.Shape shape)
        {
            _shape = shape;
        }

        public string Name
        {
            get => _shape?.Name ?? string.Empty;
            set { if (_shape != null) _shape.Name = value; }
        }

        public ShapeType ShapeType
        {
            get
            {
                try
                {
                    var type = _shape?.Type;
                    if (type == null) return ShapeType.Other;

                    switch ((int)type.Value)
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
                        _shape?.Left ?? 0,
                        _shape?.Top ?? 0,
                        _shape?.Width ?? 0,
                        _shape?.Height ?? 0
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
                try { return _shape?.Rotation ?? 0; }
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
                try { return _shape?.HasTable == NetOffice.OfficeApi.Enums.MsoTriState.msoTrue; }
                catch { return false; }
            }
        }

        public bool IsChart
        {
            get
            {
                try { return _shape?.HasChart == NetOffice.OfficeApi.Enums.MsoTriState.msoTrue; }
                catch { return false; }
            }
        }

        public bool HasTextFrame
        {
            get
            {
                try { return _shape?.HasTextFrame == NetOffice.OfficeApi.Enums.MsoTriState.msoTrue; }
                catch { return false; }
            }
        }

        public ITableContext Table
        {
            get
            {
                try
                {
                    if (!IsTable) return null;
                    var table = _shape?.Table;
                    return table != null ? new PowerPointTableContext(table) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public object NativeShape => _shape;

        /// <summary>
        /// 获取原生 Shape 对象
        /// </summary>
        public NETOP.Shape NetShape => _shape;
    }
}
