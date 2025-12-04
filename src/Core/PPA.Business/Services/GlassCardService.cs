using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using PPA.Logging;

namespace PPA.Business.Services
{
    /// <summary>
    /// 毛玻璃卡片业务服务实现（后续在此封装平台无关的业务流程，
    /// 具体绘制由各平台适配器完成）
    /// </summary>
    public class GlassCardService : IGlassCardService
    {
        private readonly ILogger _logger;
        private readonly PPAConfig _config;
        private readonly IGlassCardRenderer _renderer;

        public GlassCardService(ILogger logger, PPAConfig config, IGlassCardRenderer renderer)
        {
            _logger = logger ?? NullLogger.Instance;
            _config = config;
            _renderer = renderer;
        }

        /// <inheritdoc />
        public void CreateGlassCard(IApplicationContext application, GlassCardConfig config)
        {
            if (application == null)
            {
                _logger.LogWarning("应用上下文为空，无法创建毛玻璃卡片");
                return;
            }

            // 如果外部没有传配置，则使用全局 PPAConfig 中的 GlassCard 配置
            var effectiveConfig = config ?? _config?.GlassCard;
            if (effectiveConfig == null)
            {
                // 正常情况下不应发生：PPAConfig.LoadOrCreate 已保证解析失败时重写默认配置
                _logger.LogError("GlassCard 配置为空，已跳过毛玻璃卡片创建。请检查 PPAConfig.xml 是否被手动破坏。");
                return;
            }
            
            // 2A 阶段：仅计算目标矩形并记录日志，不实际创建形状
            var rect = GetTargetRect(application, effectiveConfig);

            _logger.LogInformation(
                $"准备创建毛玻璃卡片，平台: {application.Platform}, BorderColorIndex: {effectiveConfig.BorderColorIndex}, BlurRadius: {effectiveConfig.BlurRadius}, " +
                $"Rect: Left={rect.Left}, Top={rect.Top}, Width={rect.Width}, Height={rect.Height}");

            // 调用平台特定渲染器创建卡片形状
            _renderer?.RenderGlassCard(application, rect, effectiveConfig);
        }

        /// <summary>
        /// 计算毛玻璃卡片目标矩形：
        /// - 如果有选中形状，则使用第一个选中形状的 Bounds
        /// - 否则基于幻灯片尺寸和配置中的默认宽高比例，居中放置
        /// </summary>
        private ShapeRect GetTargetRect(IApplicationContext app, GlassCardConfig config)
        {
            try
            {
                var selection = app.Selection;
                if (selection != null && selection.Type == SelectionType.Shapes && selection.ShapeCount > 0)
                {
                    var firstShape = selection.SelectedShapes?.FirstOrDefault();
                    if (firstShape != null)
                    {
                        return firstShape.Bounds;
                    }
                }
            }
            catch (System.Exception ex)
            {
                _logger.LogWarning($"根据选中形状计算卡片矩形时出错，将退回到默认居中矩形: {ex.Message}");
            }

            // 无选中形状或发生异常时：使用幻灯片尺寸和默认比例，生成居中矩形
            try
            {
                var presentation = app.ActivePresentation;
                if (presentation == null)
                {
                    return new ShapeRect(0, 0, 400, 200);
                }

                float slideWidth = presentation.SlideWidth;
                float slideHeight = presentation.SlideHeight;

                // 防御：比例异常时使用安全默认值
                float widthRatio = (config.DefaultWidthRatio > 0 && config.DefaultWidthRatio <= 1) ? config.DefaultWidthRatio : 0.6f;
                float heightRatio = (config.DefaultHeightRatio > 0 && config.DefaultHeightRatio <= 1) ? config.DefaultHeightRatio : 0.25f;

                float cardWidth = slideWidth * widthRatio;
                float cardHeight = slideHeight * heightRatio;

                float left = (slideWidth - cardWidth) / 2f;
                float top = (slideHeight - cardHeight) / 2f;

                return new ShapeRect(left, top, cardWidth, cardHeight);
            }
            catch (System.Exception ex)
            {
                _logger.LogWarning($"根据幻灯片尺寸计算默认卡片矩形时出错，将使用固定大小: {ex.Message}");
                return new ShapeRect(100, 100, 400, 200);
            }
        }
    }
}
