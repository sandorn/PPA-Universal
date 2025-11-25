using PPA.Core.Abstraction;

namespace PPA.Business.Abstractions
{
    /// <summary>
    /// 表格批量操作服务接口
    /// </summary>
    public interface ITableBatchService
    {
        /// <summary>
        /// 批量格式化表格
        /// </summary>
        /// <param name="context">应用程序上下文</param>
        void FormatAllTables(IApplicationContext context);

        /// <summary>
        /// 批量格式化选中的表格
        /// </summary>
        void FormatSelectedTables(IApplicationContext context);

        /// <summary>
        /// 批量格式化当前幻灯片的表格
        /// </summary>
        void FormatCurrentSlideTables(IApplicationContext context);
    }

    /// <summary>
    /// 形状批量操作服务接口
    /// </summary>
    public interface IShapeBatchService
    {
        /// <summary>
        /// 批量删除选中的形状
        /// </summary>
        void DeleteSelectedShapes(IApplicationContext context);

        /// <summary>
        /// 批量复制选中的形状
        /// </summary>
        void DuplicateSelectedShapes(IApplicationContext context);

        /// <summary>
        /// 批量设置形状大小
        /// </summary>
        void ResizeSelectedShapes(IApplicationContext context, float width, float height);
    }

    /// <summary>
    /// 文本批量操作服务接口
    /// </summary>
    public interface ITextBatchService
    {
        /// <summary>
        /// 批量格式化选中形状的文本
        /// </summary>
        void FormatSelectedText(IApplicationContext context, TextFormatOptions options);

        /// <summary>
        /// 批量替换文本
        /// </summary>
        void ReplaceText(IApplicationContext context, string find, string replace);
    }

    /// <summary>
    /// 图表批量操作服务接口
    /// </summary>
    public interface IChartBatchService
    {
        /// <summary>
        /// 批量格式化选中的图表
        /// </summary>
        void FormatSelectedCharts(IApplicationContext context);

        /// <summary>
        /// 批量格式化当前幻灯片的图表
        /// </summary>
        void FormatCurrentSlideCharts(IApplicationContext context);
    }

    /// <summary>
    /// 文本格式化选项
    /// </summary>
    public class TextFormatOptions
    {
        /// <summary>字体名称</summary>
        public string FontName { get; set; }

        /// <summary>字体大小</summary>
        public float? FontSize { get; set; }

        /// <summary>字体颜色</summary>
        public int? FontColor { get; set; }

        /// <summary>是否加粗</summary>
        public bool? Bold { get; set; }

        /// <summary>是否斜体</summary>
        public bool? Italic { get; set; }
    }
}
