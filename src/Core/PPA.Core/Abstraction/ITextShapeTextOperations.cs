namespace PPA.Core.Abstraction
{
	/// <summary>
	/// 原生形状上的文本框/文本范围操作（由各平台 Adapter 实现）。
	/// </summary>
	public interface ITextShapeTextOperations
	{
		/// <summary>对带文本框的形状应用 <see cref="TextFormatOptions"/>。</summary>
		void ApplyTextFormatOptions(object nativeShape, TextFormatOptions options);

		/// <summary>对文本框整段应用 <see cref="FontStyle"/>。</summary>
		void ApplyTextBoxFont(object nativeShape, FontStyle fontStyle);

		/// <summary>在文本框内执行查找替换；无文本框或不含 find 时返回 false。</summary>
		bool TryReplaceTextInTextFrame(object nativeShape, string find, string replace);
	}
}
