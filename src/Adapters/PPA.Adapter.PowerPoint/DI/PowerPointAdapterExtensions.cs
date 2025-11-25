using Microsoft.Extensions.DependencyInjection;
using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint.DI
{
    /// <summary>
    /// PowerPoint 适配器 DI 服务注册扩展
    /// </summary>
    public static class PowerPointAdapterExtensions
    {
        /// <summary>
        /// 注册 PowerPoint 适配器服务
        /// </summary>
        public static IServiceCollection AddPowerPointAdapter(this IServiceCollection services)
        {
            // 注册 PowerPoint 特定实现
            services.AddSingleton<IShapeOperations, PowerPointShapeOps>();
            services.AddSingleton<ITableOperations, PowerPointTableOps>();
            services.AddSingleton<ISlideOperations, PowerPointSlideOps>();

            // 注意：IApplicationContext 需要在运行时由 Host 项目注册
            // 因为它依赖于 NETOP.Application 实例

            return services;
        }

        /// <summary>
        /// 注册 PowerPoint 应用程序上下文
        /// </summary>
        public static IServiceCollection AddPowerPointContext(
            this IServiceCollection services,
            NETOP.Application netApp,
            object nativeApp = null)
        {
            services.AddSingleton<IApplicationContext>(sp =>
                new PowerPointContext(netApp, nativeApp));

            return services;
        }
    }
}
