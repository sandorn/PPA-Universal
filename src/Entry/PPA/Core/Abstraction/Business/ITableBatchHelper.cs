using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 表格批量操作辅助接口 提供表格的批量格式化功能
	/// </summary>
	/// <remarks>
	/// 此接口定义了表格批量操作的接口，通过依赖注入使用，便于测试和扩展。
	/// <para>
	/// <b>迁移说明：</b>此接口将逐步迁移到新架构 <c>PPA.Business.Abstractions.ITableBatchService</c>。
	/// 新架构使用平台无关的 <c>IApplicationContext</c> 替代 NetOffice 类型。
	/// </para>
	/// </remarks>
	[Obsolete("建议迁移到 PPA.Business.Abstractions.ITableBatchService")]
	public interface ITableBatchHelper
	{
		/// <summary>
		/// 同步美化表格 对选中的表格形状或当前幻灯片上的所有表格形状进行格式化
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例，不能为 null </param>
		/// <remarks>
		/// 格式化行为：
		/// <list type="bullet">
		/// <item>
		/// <description> 如果选中了表格形状，则格式化选中的表格 </description>
		/// </item>
		/// <item>
		/// <description> 如果没有选中表格，则格式化当前幻灯片上的所有表格形状 </description>
		/// </item>
		/// <item>
		/// <description> 如果光标位于表格内，则格式化当前表格 </description>
		/// </item>
		/// </list>
		/// 格式化参数从配置文件（ <see cref="IFormattingConfig" />）中读取。
		/// </remarks>
		void FormatTables(NETOP.Application app);
	}
}
