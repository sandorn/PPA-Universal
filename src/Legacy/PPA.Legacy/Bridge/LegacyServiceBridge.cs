using System;
using Microsoft.Extensions.DependencyInjection;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Legacy.Bridge
{
    /// <summary>
    /// 服务桥接器 - 提供旧代码访问新架构服务的方式
    /// </summary>
    public static class LegacyServiceBridge
    {
        private static IServiceProvider _serviceProvider;
        private static IApplicationContext _applicationContext;
        private static ILogger _logger;

        /// <summary>
        /// 初始化桥接器
        /// </summary>
        public static void Initialize(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
            _logger = _serviceProvider.GetService<ILogger>() ?? NullLogger.Instance;
            _applicationContext = _serviceProvider.GetService<IApplicationContext>();
        }

        /// <summary>
        /// 获取服务提供者
        /// </summary>
        public static IServiceProvider ServiceProvider => _serviceProvider;

        /// <summary>
        /// 获取应用程序上下文
        /// </summary>
        public static IApplicationContext ApplicationContext => _applicationContext;

        /// <summary>
        /// 获取日志实例
        /// </summary>
        public static ILogger Logger => _logger ?? NullLogger.Instance;

        /// <summary>
        /// 获取服务
        /// </summary>
        public static T GetService<T>() where T : class
        {
            return _serviceProvider?.GetService<T>();
        }

        /// <summary>
        /// 获取必需服务
        /// </summary>
        public static T GetRequiredService<T>() where T : class
        {
            return _serviceProvider?.GetRequiredService<T>();
        }

        /// <summary>
        /// 清理资源
        /// </summary>
        public static void Cleanup()
        {
            _applicationContext = null;
            _logger = null;
            if (_serviceProvider is IDisposable disposable)
            {
                disposable.Dispose();
            }
            _serviceProvider = null;
        }
    }
}
