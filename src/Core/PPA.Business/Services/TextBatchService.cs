using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;
using NetOffice.PowerPointApi;

namespace PPA.Business.Services
{
    /// <summary>
    /// 文本批量操作服务实现
    /// </summary>
    public class TextBatchService : ITextBatchService
    {
        private readonly ILogger _logger;

        public TextBatchService(ILogger logger)
        {
            _logger = logger ?? NullLogger.Instance;
        }

        public void FormatSelectedText(IApplicationContext context, TextFormatOptions options)
        {
            if (context?.Selection == null)
            {
                _logger.LogWarning("无法获取选择对象");
                return;
            }

            var shapes = context.Selection.SelectedShapes?.ToList();
            if (shapes == null || shapes.Count == 0)
            {
                _logger.LogWarning("没有选中任何形状");
                return;
            }

            _logger.LogInformation($"格式化 {shapes.Count} 个形状的文本");

            foreach (var shape in shapes)
            {
                if (shape?.HasTextFrame != true) continue;

                try
                {
                    FormatShapeText(shape, options);
                    _logger.LogInformation($"格式化形状文本: {shape.Name}");
                }
                catch (System.Exception ex)
                {
                    _logger.LogError($"格式化形状文本失败: {ex.Message}", ex);
                }
            }

            _logger.LogInformation("文本格式化完成");
        }

        public void ReplaceText(IApplicationContext context, string find, string replace)
        {
            if (context?.ActivePresentation == null)
            {
                _logger.LogWarning("无法获取演示文稿");
                return;
            }

            if (string.IsNullOrEmpty(find))
            {
                _logger.LogWarning("查找文本不能为空");
                return;
            }

            _logger.LogInformation($"批量替换文本: '{find}' -> '{replace}'");

            // TODO: 实现文本替换逻辑
            // 需要遍历所有幻灯片和形状，查找并替换文本

            _logger.LogInformation("文本替换完成");
        }

        public void FormatTextBoxFont(IEnumerable<IShapeContext> shapes, FontStyle fontStyle = null)
        {
            var shapeList = shapes?.ToList();
            if (shapeList == null || shapeList.Count == 0)
            {
                _logger.LogWarning("没有选中任何形状");
                return;
            }

            if (fontStyle == null)
            {
                _logger.LogWarning("未提供字体样式");
                return;
            }

            _logger.LogInformation($"格式化 {shapeList.Count} 个文本框的字体");

            foreach (var shape in shapeList)
            {
                if (shape?.HasTextFrame != true) continue;

                try
                {
                    FormatTextBoxFont(shape, fontStyle);
                    _logger.LogInformation($"格式化文本框字体: {shape.Name}");
                }
                catch (System.Exception ex)
                {
                    _logger.LogError($"格式化文本框字体失败: {ex.Message}", ex);
                }
            }

            _logger.LogInformation("文本框字体格式化完成");
        }

        private void FormatShapeText(IShapeContext shape, TextFormatOptions options)
        {
            if (shape?.NativeShape == null || !shape.HasTextFrame) return;

            var platform = GetPlatform(shape);
            if (platform == PlatformType.PowerPoint)
            {
                FormatShapeTextPowerPoint(shape, options);
            }
            else if (platform == PlatformType.WPS)
            {
                FormatShapeTextWPS(shape, options);
            }
        }

        private void FormatShapeTextPowerPoint(IShapeContext shape, TextFormatOptions options)
        {
            try
            {
                var netShape = shape.NativeShape as Shape;
                if (netShape?.TextFrame?.TextRange == null) return;

                var textRange = netShape.TextFrame.TextRange;
                var font = textRange.Font;

                if (!string.IsNullOrEmpty(options.FontName))
                {
                    font.Name = options.FontName;
                }
                if (options.FontSize.HasValue)
                {
                    font.Size = options.FontSize.Value;
                }
                if (options.FontColor.HasValue)
                {
                    font.Color.RGB = options.FontColor.Value;
                }
                if (options.Bold.HasValue)
                {
                    font.Bold = options.Bold.Value 
                        ? NetOffice.OfficeApi.Enums.MsoTriState.msoTrue 
                        : NetOffice.OfficeApi.Enums.MsoTriState.msoFalse;
                }
                if (options.Italic.HasValue)
                {
                    font.Italic = options.Italic.Value 
                        ? NetOffice.OfficeApi.Enums.MsoTriState.msoTrue 
                        : NetOffice.OfficeApi.Enums.MsoTriState.msoFalse;
                }
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"PowerPoint 文本格式化失败: {ex.Message}", ex);
                throw;
            }
        }

        private void FormatShapeTextWPS(IShapeContext shape, TextFormatOptions options)
        {
            try
            {
                dynamic nativeShape = shape.NativeShape;
                if (nativeShape?.TextFrame?.TextRange == null) return;

                var textRange = nativeShape.TextFrame.TextRange;
                var font = textRange.Font;

                if (!string.IsNullOrEmpty(options.FontName))
                {
                    font.Name = options.FontName;
                }
                if (options.FontSize.HasValue)
                {
                    font.Size = options.FontSize.Value;
                }
                if (options.FontColor.HasValue)
                {
                    font.Color.RGB = options.FontColor.Value;
                }
                if (options.Bold.HasValue)
                {
                    font.Bold = options.Bold.Value ? -1 : 0; // msoTrue = -1, msoFalse = 0
                }
                if (options.Italic.HasValue)
                {
                    font.Italic = options.Italic.Value ? -1 : 0;
                }
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"WPS 文本格式化失败: {ex.Message}", ex);
                throw;
            }
        }

        private void FormatTextBoxFont(IShapeContext shape, FontStyle fontStyle)
        {
            if (shape?.NativeShape == null || !shape.HasTextFrame) return;

            var platform = GetPlatform(shape);
            if (platform == PlatformType.PowerPoint)
            {
                FormatTextBoxFontPowerPoint(shape, fontStyle);
            }
            else if (platform == PlatformType.WPS)
            {
                FormatTextBoxFontWPS(shape, fontStyle);
            }
        }

        private void FormatTextBoxFontPowerPoint(IShapeContext shape, FontStyle fontStyle)
        {
            try
            {
                var netShape = shape.NativeShape as Shape;
                if (netShape?.TextFrame?.TextRange == null) return;

                var textRange = netShape.TextFrame.TextRange;
                var font = textRange.Font;

                if (!string.IsNullOrEmpty(fontStyle.Name))
                {
                    font.Name = fontStyle.Name;
                }
                if (!string.IsNullOrEmpty(fontStyle.NameFarEast))
                {
                    font.NameFarEast = fontStyle.NameFarEast;
                }
                if (fontStyle.Size > 0)
                {
                    font.Size = fontStyle.Size;
                }
                font.Bold = fontStyle.Bold 
                    ? NetOffice.OfficeApi.Enums.MsoTriState.msoTrue 
                    : NetOffice.OfficeApi.Enums.MsoTriState.msoFalse;
                if (fontStyle.ColorRgb.HasValue)
                {
                    font.Color.RGB = fontStyle.ColorRgb.Value;
                }
                if (fontStyle.ThemeColorIndex.HasValue)
                {
                    font.Color.ObjectThemeColor = (NetOffice.OfficeApi.Enums.MsoThemeColorIndex)fontStyle.ThemeColorIndex.Value;
                }
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"PowerPoint 文本框字体格式化失败: {ex.Message}", ex);
                throw;
            }
        }

        private void FormatTextBoxFontWPS(IShapeContext shape, FontStyle fontStyle)
        {
            try
            {
                dynamic nativeShape = shape.NativeShape;
                if (nativeShape?.TextFrame?.TextRange == null) return;

                var textRange = nativeShape.TextFrame.TextRange;
                var font = textRange.Font;

                if (!string.IsNullOrEmpty(fontStyle.Name))
                {
                    font.Name = fontStyle.Name;
                }
                if (!string.IsNullOrEmpty(fontStyle.NameFarEast))
                {
                    font.NameFarEast = fontStyle.NameFarEast;
                }
                if (fontStyle.Size > 0)
                {
                    font.Size = fontStyle.Size;
                }
                font.Bold = fontStyle.Bold ? -1 : 0;
                if (fontStyle.ColorRgb.HasValue)
                {
                    font.Color.RGB = fontStyle.ColorRgb.Value;
                }
                if (fontStyle.ThemeColorIndex.HasValue)
                {
                    font.Color.ObjectThemeColor = fontStyle.ThemeColorIndex.Value;
                }
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"WPS 文本框字体格式化失败: {ex.Message}", ex);
                throw;
            }
        }

        private PlatformType GetPlatform(IShapeContext shape)
        {
            if (shape?.NativeShape == null) return PlatformType.Unknown;

            var typeName = shape.NativeShape.GetType().FullName;
            if (typeName?.Contains("NetOffice.PowerPointApi") == true)
            {
                return PlatformType.PowerPoint;
            }
            else if (typeName?.Contains("WPS") == true || typeName?.Contains("Kingsoft") == true)
            {
                return PlatformType.WPS;
            }

            return PlatformType.Unknown;
        }
    }
}

