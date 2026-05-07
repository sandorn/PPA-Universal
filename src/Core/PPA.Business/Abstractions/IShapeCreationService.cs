using System.Collections.Generic;
using PPA.Core.Abstraction;

namespace PPA.Business.Abstractions
{
	/// <summary>
	/// 形状创建服务接口
	/// </summary>
	public interface IShapeCreationService
	{
		/// <summary>
		/// 在选中形状的相同位置创建无边框矩形
		/// </summary>
		/// <param name="shapes">选中的形状集合</param>
		/// <returns>创建的矩形形状集合</returns>
		IEnumerable<IShapeContext> CreateRectanglesAtShapes(IEnumerable<IShapeContext> shapes);

		/// <summary>
		/// 在选中的幻灯片上添加无边框矩形
		/// </summary>
		/// <param name="slides">选中的幻灯片集合</param>
		/// <param name="bounds">矩形位置和大小（如果为 null，则使用默认位置）</param>
		/// <returns>创建的矩形形状集合</returns>
		IEnumerable<IShapeContext> CreateRectanglesOnSlides(IEnumerable<ISlideContext> slides, ShapeRect? bounds = null);
	}
}

