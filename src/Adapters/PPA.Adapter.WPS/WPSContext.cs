using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 应用程序上下文实现
    /// </summary>
    public class WPSContext : IApplicationContext
    {
        private readonly dynamic _app;

        public WPSContext(dynamic app)
        {
            _app = app;
        }

        public PlatformType Platform => PlatformType.WPS;

        public IPresentationContext ActivePresentation
        {
            get
            {
                try
                {
                    dynamic pres = _app?.ActivePresentation;
                    return pres != null ? new WPSPresentationContext(pres) : null;
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
                    dynamic window = _app?.ActiveWindow;
                    return window != null ? new WPSWindowContext(window) : null;
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
                    dynamic selection = _app?.ActiveWindow?.Selection;
                    return selection != null ? new WPSSelectionContext(selection) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public bool IsFeatureSupported(Feature feature)
        {
            // WPS 支持大部分功能，但某些高级功能可能受限
            switch (feature)
            {
                case Feature.TableBasic:
                case Feature.ShapeAlignment:
                case Feature.ShapeBatch:
                case Feature.UndoRedo:
                    return true;

                case Feature.TableAdvancedBorder:
                case Feature.Chart:
                case Feature.ChartAdvanced:
                case Feature.TextAdvanced:
                    return true; // 基本支持，但可能有差异

                case Feature.Shortcuts:
                    return false; // WPS 快捷键系统不同

                default:
                    return false;
            }
        }

        public object NativeApplication => _app;

        /// <summary>
        /// 获取 WPS Application 动态对象
        /// </summary>
        public dynamic Application => _app;

        /// <summary>
        /// 获取应用程序名称
        /// </summary>
        public string ApplicationName
        {
            get
            {
                try { return _app?.Name; }
                catch { return "WPS 演示"; }
            }
        }

        /// <summary>
        /// 获取应用程序版本
        /// </summary>
        public string Version
        {
            get
            {
                try { return _app?.Version; }
                catch { return "Unknown"; }
            }
        }
    }
}
