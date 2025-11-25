using System;
using Microsoft.Extensions.DependencyInjection;
using PPA.Adapter.PowerPoint;
using PPA.Adapter.PowerPoint.DI;
using PPA.Adapter.WPS;
using PPA.Adapter.WPS.DI;
using PPA.Core.Abstraction;
using PPA.Logging;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Universal.Platform
{
    /// <summary>
    /// 适配器工厂
    /// 根据检测到的平台动态创建适配器
    /// </summary>
    public class AdapterFactory
    {
        private readonly ILogger _logger;

        public AdapterFactory(ILogger logger = null)
        {
            _logger = logger ?? NullLogger.Instance;
        }

        /// <summary>
        /// 根据平台类型注册适配器服务
        /// </summary>
        public void RegisterAdapter(IServiceCollection services, PlatformType platform)
        {
            switch (platform)
            {
                case PlatformType.PowerPoint:
                    _logger.LogInformation("注册 PowerPoint 适配器");
                    services.AddPowerPointAdapter();
                    break;

                case PlatformType.WPS:
                    _logger.LogInformation("注册 WPS 适配器");
                    services.AddWPSAdapter();
                    break;

                default:
                    throw new InvalidOperationException($"不支持的平台类型: {platform}");
            }
        }

        /// <summary>
        /// 根据应用程序对象创建应用程序上下文
        /// </summary>
        public IApplicationContext CreateContext(object app, PlatformType platform)
        {
            switch (platform)
            {
                case PlatformType.PowerPoint:
                    return CreatePowerPointContext(app);

                case PlatformType.WPS:
                    return CreateWPSContext(app);

                default:
                    throw new InvalidOperationException($"不支持的平台类型: {platform}");
            }
        }

        /// <summary>
        /// 创建 PowerPoint 上下文
        /// </summary>
        private IApplicationContext CreatePowerPointContext(object app)
        {
            _logger.LogInformation("创建 PowerPoint 上下文");

            // 如果是 NetOffice Application
            if (app is NETOP.Application netApp)
            {
                return new PowerPointContext(netApp);
            }

            // 如果是原生 COM 对象，创建 NetOffice 包装器
            var wrappedApp = new NETOP.Application(null, app);
            return new PowerPointContext(wrappedApp, app);
        }

        /// <summary>
        /// 创建 WPS 上下文
        /// </summary>
        private IApplicationContext CreateWPSContext(object app)
        {
            _logger.LogInformation("创建 WPS 上下文");
            return new WPSContext(app);
        }

        /// <summary>
        /// 自动检测并获取运行中的应用程序
        /// </summary>
        public (object app, PlatformType platform) GetRunningApplication()
        {
            // 优先检测 PowerPoint
            var pptApp = PlatformDetector.GetRunningPowerPoint();
            if (pptApp != null)
            {
                _logger.LogInformation("检测到运行中的 PowerPoint");
                return (pptApp, PlatformType.PowerPoint);
            }

            // 检测 WPS
            var wpsApp = PlatformDetector.GetRunningWPS();
            if (wpsApp != null)
            {
                _logger.LogInformation("检测到运行中的 WPS");
                return (wpsApp, PlatformType.WPS);
            }

            _logger.LogWarning("未检测到运行中的演示应用程序");
            return (null, PlatformType.Unknown);
        }
    }
}
