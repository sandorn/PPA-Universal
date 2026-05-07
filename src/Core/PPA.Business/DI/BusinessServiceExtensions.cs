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

			// 撤销/重做服务
			services.AddTransient<IUndoService, UndoService>();

			// 形状可见性服务
			services.AddTransient<IShapeVisibilityService, ShapeVisibilityService>();

			// 形状创建服务
			services.AddTransient<IShapeCreationService, ShapeCreationService>();

			// 裁除服务
			services.AddTransient<ICropService, CropService>();

			// 形状复制服务
			services.AddTransient<IShapeDuplicateService, ShapeDuplicateService>();

			// 批量操作服务（依赖 ITableFormatService / IShapeOperations 等，由适配器注入）
			services.AddTransient<ITableBatchService, TableBatchService>();
			services.AddTransient<IShapeBatchService, ShapeBatchService>();
			services.AddTransient<ITextBatchService, TextBatchService>();
			services.AddTransient<IChartBatchService, ChartBatchService>();

			return services;
		}
	}
}
