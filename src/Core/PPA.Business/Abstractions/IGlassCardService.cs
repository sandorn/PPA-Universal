using PPA.Core.Abstraction;
using PPA.Core.Configuration;

namespace PPA.Business.Abstractions
{
    /// <summary>
    /// 毛玻璃卡片（Glass Card）业务服务接口（平台无关）
    /// </summary>
    public interface IGlassCardService
    {
        /// <summary>
        /// 基于当前选中形状（如有）或默认位置/尺寸创建一张毛玻璃卡片。
        /// </summary>
        /// <param name="application">当前应用上下文（用于获取选区、幻灯片大小等）</param>
        /// <param name="config">毛玻璃卡片配置（通常来自 PPAConfig.GlassCard）</param>
        void CreateGlassCard(IApplicationContext application, GlassCardConfig config);
    }
}
