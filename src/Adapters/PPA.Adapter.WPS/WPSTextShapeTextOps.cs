using PPA.Core.Abstraction;

namespace PPA.Adapter.WPS
{
	/// <summary>
	/// WPS 形状文本操作（dynamic COM）。
	/// </summary>
	public class WPSTextShapeTextOps : ITextShapeTextOperations
	{
		public void ApplyTextFormatOptions(object nativeShape, TextFormatOptions options)
		{
			if (options == null) return;
			dynamic shape = nativeShape;
			if (shape?.TextFrame?.TextRange == null) return;

			var textRange = shape.TextFrame.TextRange;
			var font = textRange.Font;

			if (!string.IsNullOrEmpty(options.FontName))
				font.Name = options.FontName;
			if (options.FontSize.HasValue)
				font.Size = options.FontSize.Value;
			if (options.FontColor.HasValue)
				font.Color.RGB = options.FontColor.Value;
			if (options.Bold.HasValue)
				font.Bold = options.Bold.Value ? -1 : 0;
			if (options.Italic.HasValue)
				font.Italic = options.Italic.Value ? -1 : 0;
		}

		public void ApplyTextBoxFont(object nativeShape, FontStyle fontStyle)
		{
			if (fontStyle == null) return;
			dynamic shape = nativeShape;
			if (shape?.TextFrame?.TextRange == null) return;

			var textRange = shape.TextFrame.TextRange;
			var font = textRange.Font;

			if (!string.IsNullOrEmpty(fontStyle.Name))
				font.Name = fontStyle.Name;
			if (!string.IsNullOrEmpty(fontStyle.NameFarEast))
				font.NameFarEast = fontStyle.NameFarEast;
			if (fontStyle.Size > 0)
				font.Size = fontStyle.Size;
			font.Bold = fontStyle.Bold ? -1 : 0;
			if (fontStyle.ColorRgb.HasValue)
				font.Color.RGB = fontStyle.ColorRgb.Value;
			if (fontStyle.ThemeColorIndex.HasValue)
				font.Color.ObjectThemeColor = fontStyle.ThemeColorIndex.Value;
		}

		public bool TryReplaceTextInTextFrame(object nativeShape, string find, string replace)
		{
			dynamic shape = nativeShape;
			if (shape?.TextFrame?.TextRange == null) return false;
			var tr = shape.TextFrame.TextRange;
			string t = tr.Text ?? string.Empty;
			if (!t.Contains(find)) return false;
			tr.Text = t.Replace(find, replace ?? string.Empty);
			return true;
		}
	}
}
