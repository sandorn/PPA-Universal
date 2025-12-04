using Microsoft.Extensions.DependencyInjection;
using PPA.Business.Abstractions;
using PPA.Business.Services;

namespace PPA.Business.DI
{
    /// <summary>
    /// PPA.Business DI 服务注册扩展
    /// </summary>
    public static class BusinessServiceExtensions
    {
        /// <summary>
        /// 注册业务服务（所有版本共享）
        /// </summary>
        public static IServiceCollection AddPPABusiness(this IServiceCollection services)
        {
            // 格式化服务（依赖 ILogger、ITableOperations、PPAConfig，由容器注入）
            services.AddTransient<ITableFormatService, TableFormatService>();

            // 对齐服务
            services.AddTransient<IAlignmentService, AlignmentService>();

            // 毛玻璃卡片服务
            services.AddTransient<IGlassCardService, GlassCardService>();

            // 批量操作服务
            // 注：这些服务的具体实现需要依赖平台适配器
            // services.AddTransient<ITableBatchService, TableBatchService>();
            // services.AddTransient<IShapeBatchService, ShapeBatchService>();
            // services.AddTransient<ITextBatchService, TextBatchService>();
            // services.AddTransient<IChartBatchService, ChartBatchService>();

            return services;
        }
    }
}
