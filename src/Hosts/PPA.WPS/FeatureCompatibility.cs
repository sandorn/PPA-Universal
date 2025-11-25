using System;
using System.Collections.Generic;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.WPS
{
    /// <summary>
    /// WPS 功能兼容性检查器
    /// 用于检测和处理 WPS 与 PowerPoint 的功能差异
    /// </summary>
    public class FeatureCompatibility
    {
        private readonly IApplicationContext _context;
        private readonly ILogger _logger;
        private readonly Dictionary<Feature, FeatureStatus> _featureCache;

        public FeatureCompatibility(IApplicationContext context, ILogger logger = null)
        {
            _context = context;
            _logger = logger ?? NullLogger.Instance;
            _featureCache = new Dictionary<Feature, FeatureStatus>();
        }

        /// <summary>
        /// 检查功能是否可用
        /// </summary>
        public bool IsFeatureAvailable(Feature feature)
        {
            var status = GetFeatureStatus(feature);
            return status == FeatureStatus.FullSupport || status == FeatureStatus.PartialSupport;
        }

        /// <summary>
        /// 获取功能状态
        /// </summary>
        public FeatureStatus GetFeatureStatus(Feature feature)
        {
            if (_featureCache.TryGetValue(feature, out var cached))
            {
                return cached;
            }

            var status = CheckFeatureSupport(feature);
            _featureCache[feature] = status;
            return status;
        }

        private FeatureStatus CheckFeatureSupport(Feature feature)
        {
            // WPS 功能支持状态
            switch (feature)
            {
                case Feature.TableBasic:
                case Feature.ShapeAlignment:
                case Feature.ShapeBatch:
                case Feature.UndoRedo:
                    return FeatureStatus.FullSupport;

                case Feature.TableAdvancedBorder:
                case Feature.Chart:
                case Feature.TextAdvanced:
                    return FeatureStatus.PartialSupport;

                case Feature.ChartAdvanced:
                    return FeatureStatus.LimitedSupport;

                case Feature.Shortcuts:
                    return FeatureStatus.NotSupported;

                default:
                    return FeatureStatus.Unknown;
            }
        }

        /// <summary>
        /// 获取功能降级建议
        /// </summary>
        public string GetFallbackSuggestion(Feature feature)
        {
            var status = GetFeatureStatus(feature);

            switch (status)
            {
                case FeatureStatus.PartialSupport:
                    return $"功能 {feature} 在 WPS 中部分支持，可能有细微差异";

                case FeatureStatus.LimitedSupport:
                    return $"功能 {feature} 在 WPS 中支持有限，建议使用基础功能";

                case FeatureStatus.NotSupported:
                    return $"功能 {feature} 在 WPS 中不支持";

                default:
                    return null;
            }
        }

        /// <summary>
        /// 执行带有兼容性检查的操作
        /// </summary>
        public void ExecuteWithCompatibilityCheck(Feature feature, Action action, Action fallbackAction = null)
        {
            var status = GetFeatureStatus(feature);

            if (status == FeatureStatus.NotSupported)
            {
                _logger.LogWarning($"功能 {feature} 在 WPS 中不支持，跳过操作");

                if (fallbackAction != null)
                {
                    _logger.LogInformation($"执行降级操作");
                    fallbackAction();
                }
                return;
            }

            if (status == FeatureStatus.LimitedSupport || status == FeatureStatus.PartialSupport)
            {
                _logger.LogInformation($"功能 {feature} 在 WPS 中支持受限，尝试执行");
            }

            try
            {
                action();
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"执行功能 {feature} 时出错: {ex.Message}");

                if (fallbackAction != null)
                {
                    _logger.LogInformation($"执行降级操作");
                    try
                    {
                        fallbackAction();
                    }
                    catch (Exception fallbackEx)
                    {
                        _logger.LogError($"降级操作也失败: {fallbackEx.Message}", fallbackEx);
                    }
                }
            }
        }
    }

    /// <summary>
    /// 功能支持状态
    /// </summary>
    public enum FeatureStatus
    {
        /// <summary>未知</summary>
        Unknown = 0,

        /// <summary>完全支持</summary>
        FullSupport = 1,

        /// <summary>部分支持（可能有细微差异）</summary>
        PartialSupport = 2,

        /// <summary>有限支持（功能受限）</summary>
        LimitedSupport = 3,

        /// <summary>不支持</summary>
        NotSupported = 4
    }
}
