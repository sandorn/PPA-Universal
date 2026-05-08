using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;
using OFFICE = NetOffice.OfficeApi.Enums;

namespace PPA.Adapter.PowerPoint
{
	/// <summary>
	/// PowerPoint 形状文本操作（NetOffice）。
	/// </summary>
	public class PowerPointTextShapeTextOps : ITextShapeTextOperations
	{
		public void ApplyTextFormatOptions(object nativeShape, TextFormatOptions options)
		{
			if (options == null) return;
			var netShape = nativeShape as NETOP.Shape;
			if (netShape?.TextFrame?.TextRange == null) return;

			var textRange = netShape.TextFrame.TextRange;
			var font = textRange.Font;

			if (!string.IsNullOrEmpty(options.FontName))
				font.Name = options.FontName;
			if (options.FontSize.HasValue)
				font.Size = options.FontSize.Value;
			if (options.FontColor.HasValue)
				font.Color.RGB = options.FontColor.Value;
			if (options.Bold.HasValue)
				font.Bold = options.Bold.Value ? OFFICE.MsoTriState.msoTrue : OFFICE.MsoTriState.msoFalse;
			if (options.Italic.HasValue)
				font.Italic = options.Italic.Value ? OFFICE.MsoTriState.msoTrue : OFFICE.MsoTriState.msoFalse;
		}

		public void ApplyTextBoxFont(object nativeShape, FontStyle fontStyle)
		{
			if (fontStyle == null) return;
			var netShape = nativeShape as NETOP.Shape;
			if (netShape?.TextFrame?.TextRange == null) return;

			var textRange = netShape.TextFrame.TextRange;
			var font = textRange.Font;

			if (!string.IsNullOrEmpty(fontStyle.Name))
				font.Name = fontStyle.Name;
			if (!string.IsNullOrEmpty(fontStyle.NameFarEast))
				font.NameFarEast = fontStyle.NameFarEast;
			if (fontStyle.Size > 0)
				font.Size = fontStyle.Size;
			font.Bold = fontStyle.Bold ? OFFICE.MsoTriState.msoTrue : OFFICE.MsoTriState.msoFalse;
			if (fontStyle.ColorRgb.HasValue)
				font.Color.RGB = fontStyle.ColorRgb.Value;
			if (fontStyle.ThemeColorIndex.HasValue)
				font.Color.ObjectThemeColor = (OFFICE.MsoThemeColorIndex)fontStyle.ThemeColorIndex.Value;
		}

		public bool TryReplaceTextInTextFrame(object nativeShape, string find, string replace)
		{
			var netShape = nativeShape as NETOP.Shape;
			var tr = netShape?.TextFrame?.TextRange;
			if (tr == null) return false;
			var t = tr.Text ?? string.Empty;
			if (!t.Contains(find)) return false;
			tr.Text = t.Replace(find, replace ?? string.Empty);
			return true;
		}
	}
}
