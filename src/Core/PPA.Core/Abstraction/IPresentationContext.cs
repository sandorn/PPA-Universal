namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 演示文稿上下文接口
    /// </summary>
    public interface IPresentationContext
    {
        /// <summary>演示文稿名称</summary>
        string Name { get; }

        /// <summary>演示文稿完整路径</summary>
        string FullName { get; }

        /// <summary>幻灯片数量</summary>
        int SlideCount { get; }

        /// <summary>幻灯片宽度（磅）</summary>
        float SlideWidth { get; }

        /// <summary>幻灯片高度（磅）</summary>
        float SlideHeight { get; }

        /// <summary>获取指定索引的幻灯片</summary>
        /// <param name="index">1-based 索引</param>
        ISlideContext GetSlide(int index);

        /// <summary>获取原生演示文稿对象</summary>
        object NativePresentation { get; }
    }
}
