namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 窗口上下文接口
    /// </summary>
    public interface IWindowContext
    {
        /// <summary>当前视图类型</summary>
        ViewType ViewType { get; }

        /// <summary>当前活动的幻灯片</summary>
        ISlideContext ActiveSlide { get; }

        /// <summary>缩放比例</summary>
        int Zoom { get; set; }

        /// <summary>获取原生窗口对象</summary>
        object NativeWindow { get; }
    }

    /// <summary>
    /// 视图类型枚举
    /// </summary>
    public enum ViewType
    {
        /// <summary>普通视图</summary>
        Normal = 1,

        /// <summary>大纲视图</summary>
        Outline = 2,

        /// <summary>幻灯片浏览视图</summary>
        SlideSorter = 3,

        /// <summary>备注页视图</summary>
        NotesPage = 4,

        /// <summary>幻灯片放映视图</summary>
        SlideShow = 5,

        /// <summary>母版视图</summary>
        Master = 6
    }
}
