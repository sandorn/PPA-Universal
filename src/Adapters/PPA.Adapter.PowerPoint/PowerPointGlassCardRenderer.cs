using System;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Core.Configuration;
using PPA.Logging;
using NETOP = NetOffice.PowerPointApi;
using NetOffice.OfficeApi.Enums;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 平台的毛玻璃卡片渲染器
    /// </summary>
    public class PowerPointGlassCardRenderer : IGlassCardRenderer
    {
        private readonly ILogger _logger;

        public PowerPointGlassCardRenderer(ILogger logger = null)
        {
            _logger = logger ?? NullLogger.Instance;
        }

        public void RenderGlassCard(IApplicationContext app, ShapeRect rect, GlassCardConfig config)
        {
            try
            {
                // 1. 获取原生幻灯片对象
                // 这里的 app.ActivePresentation 和 ActiveWindow 在 Context 里可能是包装过的
                // 最稳妥的是通过 app.ActiveWindow.ActiveSlide.NativeSlide 获取
                var slideContext = app?.ActiveWindow?.ActiveSlide;
                if (slideContext == null)
                {
                    _logger.LogWarning("无法获取当前幻灯片上下文，跳过渲染");
                    return;
                }

                // 这里的 NativeSlide 应该是 NetOffice.PowerPointApi.Slide
                if (!(slideContext.NativeSlide is NETOP.Slide slide))
                {
                    _logger.LogWarning("NativeSlide 不是有效的 PowerPoint Slide 对象");
                    return;
                }

                _logger.LogInformation("开始在 PowerPoint 上创建形状");

                // 2. 创建圆角矩形 (msoShapeRoundedRectangle = 5)
                var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, rect.Left, rect.Top, rect.Width, rect.Height);
                shape.Name = "PPA_GlassCard_" + DateTime.Now.Ticks;

                // 3. 初步设置样式（渐变毛玻璃填充 + 边框）
                // 边框：优先使用配置中的 BorderColorIndex（主题色），失败时回退到白色
                if (config.BorderWidth > 0)
                {
                    shape.Line.Visible = MsoTriState.msoTrue;
                    shape.Line.Weight = config.BorderWidth;

                    bool borderSet = false;
                    // 特殊处理：默认的 BorderColorIndex=13 在很多主题下是深色背景，
                    // 对毛玻璃卡片而言更期望是白色细边，这里直接将 13 映射为白色边框。
                    if (config.BorderColorIndex == 13)
                    {
                        shape.Line.ForeColor.RGB = 16777215; // White
                        borderSet = true;
                    }
                    else if (config.BorderColorIndex > 0)
                    {
                        try
                        {
                            shape.Line.ForeColor.ObjectThemeColor = (MsoThemeColorIndex)config.BorderColorIndex;
                            borderSet = true;
                        }
                        catch
                        {
                            // 忽略主题色设置失败，后续回退到白色
                        }
                    }

                    if (!borderSet)
                    {
                        // 默认使用白色边框，避免出现意外的深色边框
                        shape.Line.ForeColor.RGB = 16777215; // White
                    }
                }
                else
                {
                    shape.Line.Visible = MsoTriState.msoFalse;
                }

                // 填充：尝试使用配置中的渐变停靠点创建毛玻璃效果
                try
                {
                    var fill = shape.Fill;
                    fill.Visible = MsoTriState.msoTrue;

                    // 使用四个白色渐变停靠点，透明度由配置决定
                    var stopsConfig = config.GradientStops;
                    if (stopsConfig != null && stopsConfig.Length > 0)
                    {
                        foreach (var stop in stopsConfig)
                        {
                            // RGB: 白色，位置 0-1，透明度 0-1
                            var position = stop.Position / 100f;
                            var transparency = stop.Opacity / 100f;
                            try
                            {
                                fill.GradientStops.Insert(16777215, position, transparency);
                            }
                            catch {}
                        }

                        // 设置渐变角度
                        fill.GradientAngle = config.GradientDirection;
                    }
                }
                catch { }

                // 圆角调整（NetOffice 里 Adjustments 是个集合）
                // PowerPoint 圆角矩形通常用 Adjustments[1] 控制圆角大小，值范围 0-1 左右
                // 这里先用默认值，后续可以根据 config.CornerRadius 调整

                // 阴影：右下角柔和阴影（仅通过数值属性控制，避免锁死 UI 参数）
                try
                {
                    var shadow = shape.Shadow;
                    shadow.Visible = MsoTriState.msoTrue;
                    shadow.ForeColor.RGB = 0x000000;
                    shadow.Type = MsoShadowType.msoShadow23;
                    shadow.Transparency = 0.8f; // 透明
                    // 模糊半径
                    var blurRadius = config.BlurRadius > 0 ? config.BlurRadius : 10f;
                    shadow.Blur = blurRadius; // 模糊 10 磅，视觉距离就是 10 磅
                    shadow.OffsetX = 7.1f; // 水平偏移
                    shadow.OffsetY = 7.1f; // 垂直偏移
                    // shadow.Size = 100f;     // 大小 100%
                }
                catch { }

                _logger.LogInformation($"成功创建 PowerPoint 卡片形状: {shape.Name}");
            }
            catch (Exception ex)
            {
                _logger.LogError($"PowerPoint 渲染毛玻璃卡片失败: {ex.Message}", ex);
            }
        }
    }
}
