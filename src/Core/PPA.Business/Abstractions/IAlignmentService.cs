using System.Collections.Generic;
using PPA.Core.Abstraction;

namespace PPA.Business.Abstractions
{
	/// <summary>
	/// 对齐服务接口（平台无关）
	/// </summary>
	public interface IAlignmentService
	{
		/// <summary>
		/// 执行对齐操作
		/// </summary>
		/// <param name="shapes">要对齐的形状集合</param>
		/// <param name="alignment">对齐类型</param>
		/// <param name="reference">对齐参考</param>
		void Align(IEnumerable<IShapeContext> shapes, AlignmentType alignment, AlignmentReference reference);

		/// <summary>
		/// 执行分布操作
		/// </summary>
		/// <param name="shapes">要分布的形状集合</param>
		/// <param name="distribution">分布类型</param>
		void Distribute(IEnumerable<IShapeContext> shapes, DistributionType distribution);

		/// <summary>
		/// 设置形状等宽
		/// </summary>
		void SetEqualWidth(IEnumerable<IShapeContext> shapes);

		/// <summary>
		/// 设置形状等高
		/// </summary>
		void SetEqualHeight(IEnumerable<IShapeContext> shapes);

		/// <summary>
		/// 设置形状等大小
		/// </summary>
		void SetEqualSize(IEnumerable<IShapeContext> shapes);

		/// <summary>
		/// 交换两个形状的位置
		/// </summary>
		void SwapPositions(IShapeContext shape1, IShapeContext shape2);

		/// <summary>
		/// 以第一个形状为基准，将其余形状整体贴靠到基准一侧：
		/// 左：他形左 = 基准左 − (他形右 − 他形左)；右：他形左 = 基准右；
		/// 上：他形上 = 基准上 − (他形下 − 他形上)；下：他形上 = 基准下。
		/// </summary>
		/// <param name="shapes">要吸附的形状集合（第一个为基准）</param>
		/// <param name="snapDirection">吸附方向</param>
		void SnapToShape(IEnumerable<IShapeContext> shapes, SnapDirection snapDirection);

		/// <summary>
		/// 将选中对象延伸到指定方向
		/// </summary>
		/// <param name="shapes">要延伸的形状集合</param>
		/// <param name="extendDirection">延伸方向</param>
		void ExtendAlignment(IEnumerable<IShapeContext> shapes, ExtendDirection extendDirection);

		/// <summary>
		/// 交换两个形状的位置和大小
		/// </summary>
		void SwapPositionsAndSize(IShapeContext shape1, IShapeContext shape2);
	}

	/// <summary>
	/// 吸附方向
	/// </summary>
	public enum SnapDirection
	{
		/// <summary>向左：他形左 = 基准左 − (他形右 − 他形左)，整体在基准左侧</summary>
		Left,
		/// <summary>向右：他形左 = 基准右，整体在基准右侧</summary>
		Right,
		/// <summary>向上：他形上 = 基准上 − (他形下 − 他形上)，整体在基准上方</summary>
		Top,
		/// <summary>向下：他形上 = 基准下，整体在基准下方</summary>
		Bottom
	}

	/// <summary>
	/// 延伸方向
	/// </summary>
	public enum ExtendDirection
	{
		/// <summary>向左延伸</summary>
		Left,
		/// <summary>向右延伸</summary>
		Right,
		/// <summary>向上延伸</summary>
		Top,
		/// <summary>向下延伸</summary>
		Bottom
	}

	/// <summary>
	/// 对齐类型
	/// </summary>
	public enum AlignmentType
	{
		/// <summary>左对齐</summary>
		Left,

		/// <summary>右对齐</summary>
		Right,

		/// <summary>顶部对齐</summary>
		Top,

		/// <summary>底部对齐</summary>
		Bottom,

		/// <summary>水平居中对齐</summary>
		CenterHorizontal,

		/// <summary>垂直居中对齐</summary>
		CenterVertical
	}

	/// <summary>
	/// 对齐参考
	/// </summary>
	public enum AlignmentReference
	{
		/// <summary>相对于选中对象</summary>
		SelectedObjects,

		/// <summary>相对于幻灯片</summary>
		Slide,

		/// <summary>相对于第一个选中对象</summary>
		FirstObject,

		/// <summary>相对于最后一个选中对象</summary>
		LastObject
	}

	/// <summary>
	/// 分布类型
	/// </summary>
	public enum DistributionType
	{
		/// <summary>水平分布</summary>
		Horizontal,

		/// <summary>垂直分布</summary>
		Vertical
	}
}
