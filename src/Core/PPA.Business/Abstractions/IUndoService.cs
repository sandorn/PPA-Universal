using System;
using PPA.Core.Abstraction;

namespace PPA.Business.Abstractions
{
	/// <summary>
	/// 撤销/重做服务接口（平台无关）
	/// </summary>
	public interface IUndoService
	{
		/// <summary>
		/// 开始新的撤销条目（仅 PowerPoint 有效）
		/// PowerPoint: 创建撤销边界，使后续操作成为独立的撤销条目
		/// WPS: 应使用 CreateUndoScope() 方法
		/// </summary>
		/// <param name="context">应用程序上下文</param>
		/// <param name="undoEntryName">撤销条目名称</param>
		/// <returns>是否成功</returns>
		bool StartNewUndoEntry(IApplicationContext context, string undoEntryName);

		/// <summary>
		/// 结束撤销条目（已废弃，请使用 CreateUndoScope）
		/// </summary>
		/// <param name="context">应用程序上下文</param>
		void EndUndoEntry(IApplicationContext context);

		/// <summary>
		/// 创建撤销作用域（推荐方式）
		/// 使用 using 语句包裹操作代码，自动处理撤销边界
		/// PowerPoint: 调用 StartNewUndoEntry 创建边界
		/// WPS: 使用 BeginUndoGroup/EndUndoGroup 配对
		/// </summary>
		/// <param name="context">应用程序上下文</param>
		/// <param name="undoEntryName">撤销条目名称</param>
		/// <returns>IDisposable 作用域，使用 using 语句自动释放</returns>
		IDisposable CreateUndoScope(IApplicationContext context, string undoEntryName);

		/// <summary>
		/// 检查平台是否支持撤销/重做功能
		/// </summary>
		/// <param name="context">应用程序上下文</param>
		/// <returns>如果支持返回 true</returns>
		bool IsSupported(IApplicationContext context);
	}
}

