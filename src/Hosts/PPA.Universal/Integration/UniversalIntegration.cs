using System;
using PPA.Core.Abstraction;
using PPA.Logging;
using PPA.Universal.Platform;

namespace PPA.Universal.Integration
{
    /// <summary>
    /// 通用版集成帮助类
    /// 用于集成到现有 PPA 项目
    /// </summary>
    public static class UniversalIntegration
    {
        private static PPAUniversal _instance;
        private static readonly object _lock = new object();

        /// <summary>
        /// 获取单例实例
        /// </summary>
        public static PPAUniversal Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new PPAUniversal();
                        }
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// 初始化（使用指定应用程序对象）
        /// </summary>
        public static void Initialize(object application)
        {
            Instance.Startup(application);
        }

        /// <summary>
        /// 初始化（自动检测平台）
        /// </summary>
        public static void InitializeAuto()
        {
            Instance.StartupAuto();
        }

        /// <summary>
        /// 初始化（指定平台类型）
        /// </summary>
        public static void Initialize(object application, PlatformType platform)
        {
            Instance.Startup(application, platform);
        }

        /// <summary>
        /// 获取服务提供者
        /// </summary>
        public static IServiceProvider ServiceProvider => Instance.ServiceProvider;

        /// <summary>
        /// 获取应用程序上下文
        /// </summary>
        public static IApplicationContext Context => Instance.Context;

        /// <summary>
        /// 获取当前平台
        /// </summary>
        public static PlatformType Platform => Instance.Platform;

        /// <summary>
        /// 获取日志实例
        /// </summary>
        public static ILogger Logger => Instance.Logger;

        /// <summary>
        /// 获取服务
        /// </summary>
        public static T GetService<T>() where T : class
        {
            return Instance.GetService<T>();
        }

        /// <summary>
        /// 检测平台信息
        /// </summary>
        public static PlatformInfo DetectPlatform()
        {
            return PlatformDetector.Detect();
        }

        /// <summary>
        /// 清理资源
        /// </summary>
        public static void Cleanup()
        {
            lock (_lock)
            {
                _instance?.Dispose();
                _instance = null;
            }
        }
    }
}
