namespace PPA.Core.Abstraction
{
	/// <summary>
	/// 形状内文本格式化选项（与平台无关的 DTO）。
	/// </summary>
	public class TextFormatOptions
	{
		/// <summary>字体名称</summary>
		public string FontName { get; set; }

		/// <summary>字体大小</summary>
		public float? FontSize { get; set; }

		/// <summary>字体颜色（RGB）</summary>
		public int? FontColor { get; set; }

		/// <summary>是否加粗</summary>
		public bool? Bold { get; set; }

		/// <summary>是否斜体</summary>
		public bool? Italic { get; set; }
	}
}
