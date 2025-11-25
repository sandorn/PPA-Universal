using Microsoft.Extensions.DependencyInjection;
using PPA.Logging;

namespace PPA.Core.DI
{
    /// <summary>
    /// PPA.Core DI 服务注册扩展
    /// </summary>
    public static class CoreServiceExtensions
    {
        /// <summary>
        /// 注册核心服务（所有版本共享）
        /// </summary>
        public static IServiceCollection AddPPACore(this IServiceCollection services)
        {
            // 注册默认日志服务（可被平台特定实现覆盖）
            services.AddSingleton<ILogger, ConsoleLogger>();

            return services;
        }

        /// <summary>
        /// 添加自定义日志实现
        /// </summary>
        public static IServiceCollection AddPPALogger<TLogger>(this IServiceCollection services)
            where TLogger : class, ILogger
        {
            services.AddSingleton<ILogger, TLogger>();
            return services;
        }

        /// <summary>
        /// 添加指定的日志实例
        /// </summary>
        public static IServiceCollection AddPPALogger(this IServiceCollection services, ILogger logger)
        {
            services.AddSingleton(logger);
            return services;
        }
    }
}
