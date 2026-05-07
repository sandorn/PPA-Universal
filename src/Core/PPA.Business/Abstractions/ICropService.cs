using System.Collections.Generic;
using PPA.Core.Abstraction;

namespace PPA.Business.Abstractions
{
	/// <summary>
	/// 裁除服务接口
	/// </summary>
	public interface ICropService
	{
		/// <summary>
		/// 将选中形状裁除到幻灯片边界内
		/// </summary>
		/// <param name="shapes">要裁除的形状集合</param>
		/// <param name="slide">幻灯片上下文</param>
		void CropShapesToSlide(IEnumerable<IShapeContext> shapes, ISlideContext slide);

		/// <summary>
		/// 将当前幻灯片所有形状裁除到边界内
		/// </summary>
		/// <param name="slide">幻灯片上下文</param>
		void CropAllShapesToSlide(ISlideContext slide);
	}
}

