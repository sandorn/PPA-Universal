using PPA.Core.Abstraction;
using PPA.Core.Configuration;

namespace PPA.Business.Abstractions
{
    /// <summary>
    /// 毛玻璃卡片渲染器接口（平台无关）
    /// 具体绘制由各平台适配器实现
    /// </summary>
    public interface IGlassCardRenderer
    {
        /// <summary>
        /// 在当前活动幻灯片上渲染毛玻璃卡片
        /// </summary>
        /// <param name="app">应用程序上下文</param>
        /// <param name="rect">卡片位置和大小</param>
        /// <param name="config">卡片配置</param>
        void RenderGlassCard(IApplicationContext app, ShapeRect rect, GlassCardConfig config);
    }
}
