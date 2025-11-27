using System.Xml.Serialization;

namespace PPA.Manipulation.Config
{
	/// <summary>
	/// 图表格式化配置
	/// </summary>
	public class ChartFormattingConfig
	{
		/// <summary>
		/// 常规字体配置
		/// </summary>
		/// <remarks>
		/// 注意：图表字体不支持主题占位符(+mn-lt)，必须使用实际字体名
		/// </remarks>
		[XmlElement("RegularFont")]
		public FontConfig RegularFont { get; set; } = new FontConfig
		{
			Name="微软雅黑",
			NameFarEast="微软雅黑",
			Size=8.0f,
			Bold=false,
			ThemeColor="Dark1"
		};

		/// <summary>
		/// 标题字体配置
		/// </summary>
		/// <remarks>
		/// 注意：图表字体不支持主题占位符(+mn-lt)，必须使用实际字体名
		/// </remarks>
		[XmlElement("TitleFont")]
		public FontConfig TitleFont { get; set; } = new FontConfig
		{
			Name="微软雅黑",
			NameFarEast="微软雅黑",
			Size=11.0f,
			Bold=true,
			ThemeColor="Dark1"
		};
	}
}
