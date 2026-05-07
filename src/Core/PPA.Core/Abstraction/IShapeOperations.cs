using System.Collections.Generic;

namespace PPA.Core.Abstraction
{
	/// <summary>
	/// 形状操作接口（平台无关）
	/// </summary>
	public interface IShapeOperations
	{
		/// <summary>获取形状位置</summary>
		ShapeRect GetBounds(object shape);

		/// <summary>设置形状位置</summary>
		void SetBounds(object shape, ShapeRect bounds);

		/// <summary>获取形状旋转角度</summary>
		float GetRotation(object shape);

		/// <summary>设置形状旋转角度</summary>
		void SetRotation(object shape, float angle);

		/// <summary>判断是否为表格</summary>
		bool IsTable(object shape);

		/// <summary>判断是否为图表</summary>
		bool IsChart(object shape);

		/// <summary>判断是否为文本框</summary>
		bool IsTextBox(object shape);

		/// <summary>判断是否为组合形状</summary>
		bool IsGroup(object shape);

		/// <summary>复制形状</summary>
		object CopyShape(object shape);

		/// <summary>删除形状</summary>
		void DeleteShape(object shape);

		/// <summary>获取形状可见性</summary>
		bool GetVisible(object shape);

		/// <summary>设置形状可见性</summary>
		void SetVisible(object shape, bool visible);

		/// <summary>创建矩形形状</summary>
		object CreateRectangle(ISlideContext slide, ShapeRect bounds);

		/// <summary>将原生形状包装为 <see cref="IShapeContext"/>（用于遍历组合子项等）。</summary>
		IShapeContext WrapShape(object nativeShape);

		/// <summary>若为组合形状，返回子项原生对象序列；否则返回空序列。</summary>
		IEnumerable<object> GetGroupChildShapes(object groupShape);
	}
}
