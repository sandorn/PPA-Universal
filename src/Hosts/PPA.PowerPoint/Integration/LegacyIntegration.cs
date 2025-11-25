using System;
using Microsoft.Extensions.DependencyInjection;
using PPA.Adapter.PowerPoint;
using PPA.Adapter.PowerPoint.DI;
using PPA.Business.DI;
using PPA.Core.Abstraction;
using PPA.Core.DI;
using PPA.Logging;
using MSOP = Microsoft.Office.Interop.PowerPoint;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.PowerPoint.Integration
{
    /// <summary>
    /// 用于集成到现有 PPA 项目的帮助类
    /// 提供与原有代码兼容的接口
    /// </summary>
    public static class LegacyIntegration
    {
        private static IServiceProvider _serviceProvider;
        private static PowerPointContext _context;

        /// <summary>
        /// 初始化新架构服务（在现有 ThisAddIn.Startup 中调用）
        /// </summary>
        public static void Initialize(NETOP.Application netApp, MSOP.Application nativeApp)
        {
            var services = new ServiceCollection();

            // 注册核心服务
            services.AddPPACore();

            // 注册业务服务
            services.AddPPABusiness();

            // 注册 PowerPoint 适配器
            services.AddPowerPointAdapter();
            services.AddPowerPointContext(netApp, nativeApp);

            _serviceProvider = services.BuildServiceProvider();
            _context = new PowerPointContext(netApp, nativeApp);

            var logger = _serviceProvider.GetService<ILogger>();
            logger?.LogInformation("新架构服务初始化完成");
        }

        /// <summary>
        /// 获取服务提供者
        /// </summary>
        public static IServiceProvider ServiceProvider => _serviceProvider;

        /// <summary>
        /// 获取应用程序上下文
        /// </summary>
        public static IApplicationContext Context => _context;

        /// <summary>
        /// 获取服务
        /// </summary>
        public static T GetService<T>() where T : class
        {
            return _serviceProvider?.GetService<T>();
        }

        /// <summary>
        /// 释放资源（在现有 ThisAddIn.Shutdown 中调用）
        /// </summary>
        public static void Cleanup()
        {
            if (_serviceProvider is IDisposable disposable)
            {
                disposable.Dispose();
            }
            _serviceProvider = null;
            _context = null;
        }
    }
}
