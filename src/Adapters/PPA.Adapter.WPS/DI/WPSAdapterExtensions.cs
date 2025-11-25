using Microsoft.Extensions.DependencyInjection;
using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS.DI
{
    /// <summary>
    /// WPS 适配器 DI 服务注册扩展
    /// </summary>
    public static class WPSAdapterExtensions
    {
        /// <summary>
        /// 注册 WPS 适配器服务
        /// </summary>
        public static IServiceCollection AddWPSAdapter(this IServiceCollection services)
        {
            // 注册 WPS 特定实现
            services.AddSingleton<IShapeOperations, WPSShapeOps>();
            services.AddSingleton<ITableOperations, WPSTableOps>();
            services.AddSingleton<ISlideOperations, WPSSlideOps>();

            // 注意：IApplicationContext 需要在运行时由 Host 项目注册
            // 因为它依赖于 WPS Application 实例

            return services;
        }

        /// <summary>
        /// 注册 WPS 应用程序上下文
        /// </summary>
        public static IServiceCollection AddWPSContext(
            this IServiceCollection services,
            dynamic wpsApp)
        {
            services.AddSingleton<IApplicationContext>(sp =>
                new WPSContext(wpsApp));

            return services;
        }

        /// <summary>
        /// 注册 WPS 应用程序上下文（延迟创建）
        /// </summary>
        public static IServiceCollection AddWPSContextFactory(
            this IServiceCollection services)
        {
            services.AddSingleton<IApplicationContext>(sp =>
            {
                var app = WPSHelper.GetRunningWPSApplication()
                    ?? WPSHelper.CreateWPSApplication();
                return new WPSContext(app);
            });

            return services;
        }
    }
}
