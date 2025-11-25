namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 平台无关的应用程序上下文接口
    /// </summary>
    public interface IApplicationContext
    {
        /// <summary>当前平台类型</summary>
        PlatformType Platform { get; }

        /// <summary>当前活动的演示文稿</summary>
        IPresentationContext ActivePresentation { get; }

        /// <summary>当前活动的窗口</summary>
        IWindowContext ActiveWindow { get; }

        /// <summary>当前选择</summary>
        ISelectionContext Selection { get; }

        /// <summary>检查特性是否支持</summary>
        /// <param name="feature">要检查的特性</param>
        /// <returns>如果支持返回 true</returns>
        bool IsFeatureSupported(Feature feature);

        /// <summary>获取原生应用程序对象（平台特定）</summary>
        object NativeApplication { get; }
    }
}
