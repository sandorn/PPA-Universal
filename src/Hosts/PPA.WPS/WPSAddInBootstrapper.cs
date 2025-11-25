using System;
using Microsoft.Extensions.DependencyInjection;
using PPA.Adapter.WPS;
using PPA.Adapter.WPS.DI;
using PPA.Business.DI;
using PPA.Core.Abstraction;
using PPA.Core.DI;
using PPA.Logging;

namespace PPA.WPS
{
    /// <summary>
    /// PPA WPS 插件引导程序
    /// 负责初始化 DI 容器和应用程序上下文
    /// </summary>
    public class WPSAddInBootstrapper : IDisposable
    {
        private IServiceProvider _serviceProvider;
        private dynamic _wpsApp;
        private ILogger _logger;
        private bool _disposed;
        private bool _ownsApp;

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
        /// 获取 WPS Application 实例
        /// </summary>
        public dynamic WPSApplication => _wpsApp;

        /// <summary>
        /// 初始化插件（使用现有的 WPS 应用程序实例）
        /// </summary>
        /// <param name="wpsApp">WPS Application COM 对象</param>
        public void Initialize(dynamic wpsApp)
        {
            _wpsApp = wpsApp ?? throw new ArgumentNullException(nameof(wpsApp));
            _ownsApp = false;

            // 初始化 DI 容器
            InitializeDIContainer();

            _logger?.LogInformation("PPA WPS 插件初始化完成");
        }

        /// <summary>
        /// 初始化插件（自动获取或创建 WPS 应用程序实例）
        /// </summary>
        public void InitializeAuto()
        {
            // 尝试获取正在运行的 WPS 实例
            _wpsApp = WPSHelper.GetRunningWPSApplication();

            if (_wpsApp == null)
            {
                // 创建新的 WPS 实例
                _wpsApp = WPSHelper.CreateWPSApplication();
                _ownsApp = true;
            }
            else
            {
                _ownsApp = false;
            }

            // 初始化 DI 容器
            InitializeDIContainer();

            _logger?.LogInformation($"PPA WPS 插件初始化完成（自动模式，owns={_ownsApp}）");
        }

        private void InitializeDIContainer()
        {
            var services = new ServiceCollection();

            // 注册核心服务
            services.AddPPACore();

            // 注册业务服务
            services.AddPPABusiness();

            // 注册 WPS 适配器
            services.AddWPSAdapter();

            // 注册 WPS 应用程序上下文（直接注册，避免 dynamic 扩展方法问题）
            services.AddSingleton<IApplicationContext>(sp => new WPSContext(_wpsApp));

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
                _logger?.LogInformation("正在释放 PPA WPS 插件资源");

                // 释放 DI 容器
                if (_serviceProvider is IDisposable disposable)
                {
                    disposable.Dispose();
                }
                _serviceProvider = null;

                // 如果是我们创建的 WPS 实例，需要释放
                if (_ownsApp && _wpsApp != null)
                {
                    try
                    {
                        _wpsApp.Quit();
                    }
                    catch { }

                    WPSHelper.SafeRelease(_wpsApp);
                }
                _wpsApp = null;
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
