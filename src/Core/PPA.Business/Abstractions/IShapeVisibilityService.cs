using System.Collections.Generic;
using PPA.Core.Abstraction;

namespace PPA.Business.Abstractions
{
	/// <summary>
	/// 形状可见性服务接口
	/// </summary>
	public interface IShapeVisibilityService
	{
		/// <summary>
		/// 隐藏选中的形状
		/// </summary>
		/// <param name="shapes">要隐藏的形状集合</param>
		void HideShapes(IEnumerable<IShapeContext> shapes);

		/// <summary>
		/// 显示当前幻灯片中所有隐藏的形状
		/// </summary>
		/// <param name="slide">幻灯片上下文</param>
		void ShowAllHiddenShapes(ISlideContext slide);
	}
}

