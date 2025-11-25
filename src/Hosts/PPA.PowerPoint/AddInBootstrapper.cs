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

namespace PPA.PowerPoint
{
    /// <summary>
    /// PPA PowerPoint 插件引导程序
    /// 负责初始化 DI 容器和应用程序上下文
    /// </summary>
    public class AddInBootstrapper : IDisposable
    {
        private IServiceProvider _serviceProvider;
        private NETOP.Application _netApp;
        private MSOP.Application _nativeApp;
        private ILogger _logger;
        private bool _disposed;

        /// <summary>
        /// 获取服务提供者
        /// </summary>
        public IServiceProvider ServiceProvider => _serviceProvider;

        /// <summary>
        /// 获取日志实例
        /// </summary>
        public ILogger Logger => _logger ?? NullLogger.Instance;

        /// <summary>
        /// 获取应用程序上下文
        /// </summary>
        public IApplicationContext ApplicationContext => _serviceProvider?.GetService<IApplicationContext>();

        /// <summary>
        /// 获取 NetOffice Application 实例
        /// </summary>
        public NETOP.Application NetApplication => _netApp;

        /// <summary>
        /// 初始化插件
        /// </summary>
        /// <param name="nativeApp">原生 PowerPoint Application</param>
        public void Initialize(MSOP.Application nativeApp)
        {
            _nativeApp = nativeApp ?? throw new ArgumentNullException(nameof(nativeApp));

            // 创建 NetOffice 包装器
            _netApp = new NETOP.Application(null, _nativeApp);

            // 初始化 DI 容器
            InitializeDIContainer();

            _logger?.LogInformation("PPA PowerPoint 插件初始化完成");
        }

        private void InitializeDIContainer()
        {
            var services = new ServiceCollection();

            // 注册核心服务
            services.AddPPACore();

            // 注册业务服务
            services.AddPPABusiness();

            // 注册 PowerPoint 适配器
            services.AddPowerPointAdapter();
            services.AddPowerPointContext(_netApp, _nativeApp);

            // 构建服务提供者
            _serviceProvider = services.BuildServiceProvider();

            // 获取日志服务
            _logger = _serviceProvider.GetService<ILogger>() ?? NullLogger.Instance;
            _logger.LogInformation("DI 容器初始化成功");
        }

        /// <summary>
        /// 获取服务
        /// </summary>
        public T GetService<T>() where T : class
        {
            return _serviceProvider?.GetService<T>();
        }

        /// <summary>
        /// 获取必需服务
        /// </summary>
        public T GetRequiredService<T>() where T : class
        {
            return _serviceProvider?.GetRequiredService<T>();
        }

        public void Dispose()
        {
            if (_disposed) return;

            try
            {
                _logger?.LogInformation("正在释放 PPA PowerPoint 插件资源");

                // 释放 DI 容器
                if (_serviceProvider is IDisposable disposable)
                {
                    disposable.Dispose();
                }
                _serviceProvider = null;

                // 释放 NetOffice Application
                _netApp?.Dispose();
                _netApp = null;

                _nativeApp = null;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"释放资源时出错: {ex.Message}", ex);
            }
            finally
            {
                _disposed = true;
            }
        }
    }
}
