using System;
using PPA.Core.Abstraction;
using PPA.Logging;
using PPA.Universal.Platform;

namespace PPA.Universal
{
    /// <summary>
    /// PPA 通用版主入口类
    /// 提供简化的 API，自动适配 PowerPoint 和 WPS
    /// </summary>
    public class PPAUniversal : IDisposable
    {
        private readonly UniversalBootstrapper _bootstrapper;
        private bool _disposed;

        public PPAUniversal()
        {
            _bootstrapper = new UniversalBootstrapper();
        }

        /// <summary>
        /// 获取引导程序
        /// </summary>
        public UniversalBootstrapper Bootstrapper => _bootstrapper;

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
        /// 获取当前平台类型
        /// </summary>
        public PlatformType Platform => _bootstrapper.Platform;

        /// <summary>
        /// 检测当前平台信息
        /// </summary>
        public static PlatformInfo DetectPlatform()
        {
            return PlatformDetector.Detect();
        }

        /// <summary>
        /// 启动插件（使用指定的应用程序对象）
        /// </summary>
        public void Startup(object application)
        {
            try
            {
                _bootstrapper.Initialize(application);
                Logger.LogInformation($"PPA 通用版已启动，平台: {Platform}");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启动插件失败: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// 启动插件（自动检测平台）
        /// </summary>
        public void StartupAuto()
        {
            try
            {
                _bootstrapper.InitializeAuto();
                Logger.LogInformation($"PPA 通用版已启动（自动模式），平台: {Platform}");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启动插件失败: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// 启动插件（指定平台类型）
        /// </summary>
        public void Startup(object application, PlatformType platform)
        {
            try
            {
                _bootstrapper.Initialize(application, platform);
                Logger.LogInformation($"PPA 通用版已启动，平台: {Platform}");
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
            Logger.LogInformation("PPA 通用版正在关闭");
            Dispose();
        }

        /// <summary>
        /// 获取服务
        /// </summary>
        public T GetService<T>() where T : class
        {
            return _bootstrapper.GetService<T>();
        }

        public void Dispose()
        {
            if (_disposed) return;

            _bootstrapper?.Dispose();
            _disposed = true;
        }
    }
}
