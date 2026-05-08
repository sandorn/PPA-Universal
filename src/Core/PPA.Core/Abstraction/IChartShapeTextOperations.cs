namespace PPA.Core.Abstraction
{
	/// <summary>
	/// 图表形状上标题/图例/刻度与数据标签等字体设置（由各平台 Adapter 实现）。
	/// </summary>
	public interface IChartShapeTextOperations
	{
		/// <summary>
		/// <paramref name="primaryFont"/> 用于标题、坐标轴刻度、数据标签等；<paramref name="legendFont"/> 用于图例。
		/// </summary>
		void ApplyChartFonts(object nativeShape, FontStyle primaryFont, FontStyle legendFont);
	}
}
