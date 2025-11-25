using System;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.WPS
{
    /// <summary>
    /// PPA WPS 插件主类
    /// </summary>
    public class WPSAddIn : IDisposable
    {
        private readonly WPSAddInBootstrapper _bootstrapper;
        private bool _disposed;

        public WPSAddIn()
        {
            _bootstrapper = new WPSAddInBootstrapper();
        }

        /// <summary>
        /// 获取引导程序
        /// </summary>
        public WPSAddInBootstrapper Bootstrapper => _bootstrapper;

        /// <summary>
        /// 获取服务提供者
        /// </summary>
        public IServiceProvider ServiceProvider => _bootstrapper.ServiceProvider;

        /// <summary>
        /// 获取应用程序上下文
        /// </summary>
        public IApplicationContext Context => _bootstrapper.ApplicationContext;

        /// <summary>
        /// 获取日志实例
        /// </summary>
        public ILogger Logger => _bootstrapper.Logger;

        /// <summary>
        /// 启动插件（使用现有的 WPS 应用程序实例）
        /// </summary>
        public void Startup(dynamic wpsApplication)
        {
            try
            {
                _bootstrapper.Initialize(wpsApplication);
                Logger.LogInformation("PPA WPS 插件已启动");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启动插件失败: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// 启动插件（自动获取或创建 WPS 应用程序实例）
        /// </summary>
        public void StartupAuto()
        {
            try
            {
                _bootstrapper.InitializeAuto();
                Logger.LogInformation("PPA WPS 插件已启动（自动模式）");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启动插件失败: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// 关闭插件
        /// </summary>
        public void Shutdown()
        {
            Logger.LogInformation("PPA WPS 插件正在关闭");
            Dispose();
        }

        public void Dispose()
        {
            if (_disposed) return;

            _bootstrapper?.Dispose();
            _disposed = true;
        }
    }
}
