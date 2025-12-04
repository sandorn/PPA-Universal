using System;
using System.IO;
using Microsoft.Extensions.DependencyInjection;
using PPA.Adapter.PowerPoint;
using PPA.Adapter.WPS;
using PPA.Business.DI;
using PPA.Core.Abstraction;
using PPA.Core.DI;
using PPA.Core.Configuration;
using PPA.Logging;
using PPA.Universal.Platform;
using System.Xml.Linq;

namespace PPA.Universal
{
    /// <summary>
    /// PPA 通用版引导程序
    /// 自动检测平台并初始化相应的适配器
    /// </summary>
    public class UniversalBootstrapper : IDisposable
    {
        private IServiceProvider _serviceProvider;
        private object _app;
        private PlatformType _platform;
        private ILogger _logger;
        private bool _disposed;
        private readonly AdapterFactory _adapterFactory;

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
        /// 获取当前平台类型
        /// </summary>
        public PlatformType Platform => _platform;

        /// <summary>
        /// 获取原生应用程序对象
        /// </summary>
        public object NativeApplication => _app;

        public UniversalBootstrapper()
        {
            _adapterFactory = new AdapterFactory();
        }

        /// <summary>
        /// 使用指定的应用程序对象初始化
        /// </summary>
        public void Initialize(object app)
        {
            if (app == null)
                throw new ArgumentNullException(nameof(app));

            _app = app;
            _platform = PlatformDetector.DetectFromApplication(app);

            if (_platform == PlatformType.Unknown)
            {
                throw new InvalidOperationException("无法识别应用程序类型");
            }

            InitializeDIContainer();
            _logger?.LogInformation($"PPA 通用版初始化完成，平台: {_platform}");
        }

        /// <summary>
        /// 自动检测并初始化
        /// </summary>
        public void InitializeAuto()
        {
            var (app, platform) = _adapterFactory.GetRunningApplication();

            if (app == null || platform == PlatformType.Unknown)
            {
                throw new InvalidOperationException("未找到运行中的 PowerPoint 或 WPS 实例");
            }

            _app = app;
            _platform = platform;

            InitializeDIContainer();
            _logger?.LogInformation($"PPA 通用版自动初始化完成，平台: {_platform}");
        }

        /// <summary>
        /// 指定平台类型初始化
        /// </summary>
        public void Initialize(object app, PlatformType platform)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _platform = platform;

            if (_platform == PlatformType.Unknown)
            {
                throw new ArgumentException("必须指定有效的平台类型", nameof(platform));
            }

            InitializeDIContainer();
            _logger?.LogInformation($"PPA 通用版初始化完成，平台: {_platform}");
        }

        private void InitializeDIContainer()
        {
            var services = new ServiceCollection();

            // 注册核心服务（会注册默认的 ConsoleLogger）
            services.AddPPACore();

            // 计算日志和配置路径
            var logFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PPA.Universal",
                "PPA.Universal.log");

            var logDirectory = Path.GetDirectoryName(logFilePath) ?? AppDomain.CurrentDomain.BaseDirectory;
            var configPath = Path.Combine(logDirectory, "PPAConfig.xml");

            // 加载或创建配置文件（所有 XML 解析逻辑集中在 PPAConfig 内部）
            var ppaConfig = PPAConfig.LoadOrCreate(configPath);

            // 注册配置对象
            services.AddSingleton(ppaConfig);

            // 根据配置决定是否启用文件日志及其参数
            var logging = ppaConfig.Logging;
            var enableFileLogging = logging == null || logging.EnableFileLogging;

            if (enableFileLogging)
            {
                var minLevel = ParseLogLevel(logging?.MinimumLogLevel) ?? LogLevel.Information;
                var maxLogFiles = (logging?.MaxLogFiles ?? 0) > 0 ? logging.MaxLogFiles : 14;
                int? maxLogAgeDays = (logging?.MaxLogAgeDays ?? 0) > 0 ? logging.MaxLogAgeDays : (int?)null;
                var rollingSizeMb = (logging?.RollingFileSizeMB ?? 0) > 0 ? logging.RollingFileSizeMB : 50;
                var maxFileSizeBytes = (long)rollingSizeMb * 1024 * 1024;

                // 注册文件日志服务（覆盖默认的 ConsoleLogger）
                services.AddPPALogger(new FileLogger(logFilePath, minLevel, maxLogFiles, maxLogAgeDays, maxFileSizeBytes));
            }

            // 注册业务服务
            services.AddPPABusiness();

            // 根据平台注册适配器
            _adapterFactory.RegisterAdapter(services, _platform);

            // 注册应用程序上下文
            var context = _adapterFactory.CreateContext(_app, _platform);
            services.AddSingleton<IApplicationContext>(context);

            // 构建服务提供者
            _serviceProvider = services.BuildServiceProvider();

            // 获取日志服务
            _logger = _serviceProvider.GetService<ILogger>() ?? NullLogger.Instance;

            // 诊断配置路径，避免过多的详细调试日志
            try
            {
                _logger.LogInformation($"配置文件路径: {configPath}");
            }
            catch
            {
                // 诊断日志失败时忽略，避免影响主流程
            }

            _logger.LogInformation($"DI 容器初始化成功，平台: {_platform}");
        }

        private static LogLevel? ParseLogLevel(string level)
        {
            if (string.IsNullOrWhiteSpace(level)) return null;

            try
            {
                if (Enum.TryParse<LogLevel>(level, ignoreCase: true, out var result))
                {
                    return result;
                }
            }
            catch
            {
                // 忽略解析失败，回退到默认级别
            }

            return null;
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
                _logger?.LogInformation("正在释放 PPA 通用版资源");

                // 释放 DI 容器
                if (_serviceProvider is IDisposable disposable)
                {
                    disposable.Dispose();
                }
                _serviceProvider = null;
                _app = null;
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
