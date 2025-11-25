using System;
using PPA.Core.Abstraction;
using PPA.Logging;
using MSOP = Microsoft.Office.Interop.PowerPoint;

namespace PPA.PowerPoint
{
    /// <summary>
    /// PPA PowerPoint 插件主类
    /// 这是一个示例类，展示如何在 VSTO Add-in 中使用新架构
    /// 实际使用时需要集成到现有的 ThisAddIn 类中
    /// </summary>
    public class PPAAddIn : IDisposable
    {
        private readonly AddInBootstrapper _bootstrapper;
        private bool _disposed;

        public PPAAddIn()
        {
            _bootstrapper = new AddInBootstrapper();
        }

        /// <summary>
        /// 获取引导程序
        /// </summary>
        public AddInBootstrapper Bootstrapper => _bootstrapper;

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
        /// 启动插件
        /// </summary>
        public void Startup(MSOP.Application application)
        {
            try
            {
                _bootstrapper.Initialize(application);
                Logger.LogInformation("PPA PowerPoint 插件已启动");
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
            Logger.LogInformation("PPA PowerPoint 插件正在关闭");
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
