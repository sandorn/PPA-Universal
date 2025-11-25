using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 应用程序上下文实现
    /// </summary>
    public class PowerPointContext : IApplicationContext
    {
        private readonly NETOP.Application _netApp;
        private readonly object _nativeApp;

        public PowerPointContext(NETOP.Application netApp, object nativeApp = null)
        {
            _netApp = netApp;
            _nativeApp = nativeApp;
        }

        public PlatformType Platform => PlatformType.PowerPoint;

        public IPresentationContext ActivePresentation
        {
            get
            {
                try
                {
                    var pres = _netApp?.ActivePresentation;
                    return pres != null ? new PowerPointPresentationContext(pres) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public IWindowContext ActiveWindow
        {
            get
            {
                try
                {
                    var window = _netApp?.ActiveWindow;
                    return window != null ? new PowerPointWindowContext(window) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public ISelectionContext Selection
        {
            get
            {
                try
                {
                    var selection = _netApp?.ActiveWindow?.Selection;
                    return selection != null ? new PowerPointSelectionContext(selection) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public bool IsFeatureSupported(Feature feature)
        {
            // PowerPoint 支持所有功能
            switch (feature)
            {
                case Feature.TableBasic:
                case Feature.TableAdvancedBorder:
                case Feature.Chart:
                case Feature.ChartAdvanced:
                case Feature.ShapeAlignment:
                case Feature.ShapeBatch:
                case Feature.TextAdvanced:
                case Feature.UndoRedo:
                case Feature.Shortcuts:
                    return true;
                default:
                    return false;
            }
        }

        public object NativeApplication => _nativeApp ?? (object)_netApp;

        /// <summary>
        /// 获取 NetOffice Application 实例
        /// </summary>
        public NETOP.Application NetApplication => _netApp;

        /// <summary>
        /// 获取原生 COM Application 实例
        /// </summary>
        public object NativeApp => _nativeApp;
    }
}
