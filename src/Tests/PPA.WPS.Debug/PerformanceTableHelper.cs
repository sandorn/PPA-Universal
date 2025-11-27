using NetOffice.OfficeApi.Enums;
using System.Collections.Generic;
using System.Globalization;
using NETOP = NetOffice.PowerPointApi;
using System;

namespace PPA.WPS.Debug
{
    // --- 本地定义的配置接口，解耦 PPA.Core ---
    public interface IDebugFormattingConfig
    {
        IDebugTableConfig Table { get; }
    }

    public interface IDebugTableConfig
    {
        string StyleId { get; }
        IDebugFontConfig HeaderRowFont { get; }
        IDebugFontConfig DataRowFont { get; }
        int HeaderRowBorderColor { get; }
        int DataRowBorderColor { get; }
        float HeaderRowBorderWidth { get; }
        float DataRowBorderWidth { get; }
        bool AutoNumberFormat { get; }
        int DecimalPlaces { get; }
        int NegativeTextColor { get; }
        IDebugTableSettings TableSettings { get; }
    }

    public interface IDebugFontConfig
    {
        string Name { get; }
        string NameFarEast { get; }
        float Size { get; }
        int ThemeColor { get; }
    }

    public interface IDebugTableSettings
    {
        bool FirstRow { get; }
        bool FirstCol { get; }
        bool LastRow { get; }
        bool LastCol { get; }
        bool HorizBanding { get; }
        bool VertBanding { get; }
    }

	/// <summary>
	/// 表格格式化辅助类 (Debug 版)
	/// </summary>
	public class PerformanceTableHelper
	{
		private readonly IDebugFormattingConfig _config;

		// WPS 兼容性探测标志 - 首次失败后跳过该类型的所有操作
		// 这样只需支付一次异常开销，而不是每个单元格都支付
		private bool _canSetFontName = true;
		private bool _canSetFontNameFarEast = true;
		private bool _canSetFontSize = true;
		private bool _canSetFontBold = true;
		private bool _canSetFontColor = true;
		private bool _canSetAlignment = true;
		private bool _canSetBorderVisible = true;
		private bool _canSetBorderWeight = true;
		private bool _canSetBorderTransparency = true;
		private bool _canSetBorderColor = true;

		public PerformanceTableHelper(IDebugFormattingConfig config)
		{
			_config = config ?? throw new ArgumentNullException(nameof(config));
		}

		/// <summary>
		/// 格式化表格
		/// </summary>
		/// <remarks>
		/// WPS 性能限制说明：
		/// 经过大量测试，WPS 的 COM API 存在以下限制：
		/// 1. 不支持 WM_SETREDRAW 消息暂停重绘
		/// 2. 不支持 LockWindowUpdate API
		/// 3. 不支持 Application.ScreenUpdating 属性
		/// 4. 不支持修改表格单元格 Fill 属性（背景色）
		/// 5. 不支持 ExecuteMso/ApplyStyle 清除表格样式
		/// 
		/// 因此，WPS 表格格式化的性能约为 PowerPoint 的 1/10。
		/// 这是 WPS COM 实现的固有限制，无法通过编程方式优化。
		/// 
		/// 对于需要高性能的场景，建议使用 Microsoft PowerPoint。
		/// </remarks>
		public void FormatTables(NETOP.Table tbl)
		{
			FormatTablesInternal(tbl);
		}

		private void FormatTablesInternal(NETOP.Table tbl)
		{
			var tableConfig = _config.Table;

			string styleId = tableConfig.StyleId;
			// 简化：直接强转颜色
			MsoThemeColorIndex dk1 = (MsoThemeColorIndex)tableConfig.DataRowFont.ThemeColor;
			MsoThemeColorIndex a1 = (MsoThemeColorIndex)tableConfig.HeaderRowBorderColor;
			MsoThemeColorIndex a2 = (MsoThemeColorIndex)tableConfig.DataRowBorderColor;

			float thin = tableConfig.DataRowBorderWidth;
			float thick = tableConfig.HeaderRowBorderWidth;
			float fontSize = tableConfig.DataRowFont.Size;
			float bigFontSize = tableConfig.HeaderRowFont.Size;
			string fontName = tableConfig.DataRowFont.Name;
			string fontNameFarEast = tableConfig.DataRowFont.NameFarEast;

			bool useAutoNum = tableConfig.AutoNumberFormat;
			int decimalPlacesValue = tableConfig.DecimalPlaces;

			int rows = tbl.Rows.Count;
			int cols = tbl.Columns.Count;

			// WPS 兼容性限制：
			// ExecuteMso("TableStyleClearTable") 和 ApplyStyle() 在 WPS 中均不可用。
			// 跳过这些操作以保持性能。背景清除需要用户手动完成。
			// （保留代码结构以便将来 WPS 更新后可能重新启用）
			/*
			try 
			{
				try {
					dynamic parentShape = tbl.Parent;
					parentShape.Select();
					System.Threading.Thread.Sleep(100);
				} catch { }
                
				try {
					tbl.Application.CommandBars.ExecuteMso("TableStyleClearTable");
				} catch { }
                
				try {
					tbl.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}", true);
				} catch {
					try {
						tbl.ApplyStyle("{2D5ABB26-0587-4C30-8999-92F81FD0307C}", true);
					} catch { }
				}
			}
			catch { }
			*/

			var tableSettings = tableConfig.TableSettings;
			tbl.FirstRow = tableSettings.FirstRow;
			tbl.FirstCol = tableSettings.FirstCol;
			tbl.LastRow = tableSettings.LastRow;
			tbl.LastCol = tableSettings.LastCol;
			tbl.HorizBanding = tableSettings.HorizBanding;
			tbl.VertBanding = tableSettings.VertBanding;

			var firstRowCells = new List<NETOP.Cell>();
			var lastRowCells = new List<NETOP.Cell>();
			var dataRowCells = new List<NETOP.Cell>();
            var rowsToDispose = new List<NETOP.Row>();

			try
			{
				for(int r = 1; r <= rows; r++)
				{
					var row = tbl.Rows[r];
                    rowsToDispose.Add(row);

					for(int c = 1; c <= cols; c++)
					{
						var cell = row.Cells[c];
						dataRowCells.Add(cell);
						if(r == 1) firstRowCells.Add(cell);
						else if(r == rows) lastRowCells.Add(cell);
					}
				}

				FormatDataRowCells(dataRowCells, fontName, fontNameFarEast, fontSize, dk1, thin, a2, useAutoNum, decimalPlacesValue, tableConfig.NegativeTextColor);
				FormatOutsideRowCells(firstRowCells, lastRowCells, tableConfig.HeaderRowFont.Name, tableConfig.HeaderRowFont.NameFarEast, bigFontSize, dk1, thick, a1);
			}
            finally
            {
                // 统一释放资源
                foreach(var cell in dataRowCells) cell.Dispose();
                foreach(var row in rowsToDispose) row.Dispose(); 
            }
		}

		private void FormatOutsideRowCells(List<NETOP.Cell> firstRowCells, List<NETOP.Cell> lastRowCells, string fontName, string fontNameFarEast, float fontSize, MsoThemeColorIndex txtColor, float borderWidth, MsoThemeColorIndex borderColor)
		{
			for(int i = 0; i < firstRowCells.Count; i++)
			{
                try
                {
				    var cell = firstRowCells[i];
				    ClearCellBackground(cell);

				    var textRange = cell.Shape.TextFrame.TextRange;
				    SetFontProperties(textRange, fontName, fontNameFarEast, fontSize, MsoTriState.msoTrue, txtColor);

				    if (_canSetAlignment) { try { textRange.ParagraphFormat.Alignment = NETOP.Enums.PpParagraphAlignment.ppAlignCenter; } catch { _canSetAlignment = false; } }

				    SetBorder(cell, NETOP.Enums.PpBorderType.ppBorderTop, borderWidth, borderColor);
				    SetBorder(cell, NETOP.Enums.PpBorderType.ppBorderBottom, borderWidth, borderColor);
                }
                catch { /* 忽略单个单元格错误 */ }
			}

			for(int i = 0; i < lastRowCells.Count; i++)
			{
                try
                {
				    var cell = lastRowCells[i];
				    SetBorder(cell, NETOP.Enums.PpBorderType.ppBorderBottom, borderWidth, borderColor);
                }
                catch { }
			}
		}

		private void FormatDataRowCells(List<NETOP.Cell> cells, string fontName, string fontNameFarEast, float fontSize, MsoThemeColorIndex txtColor, float thinBorderWidth, MsoThemeColorIndex thinBorderColor, bool autonum, int decimalPlaces, int negativeTextColor)
		{
			int cellCount = cells.Count;

			for(int i = 0; i < cellCount; i++)
			{
                try
                {
				    var cell = cells[i];
				    ClearCellBackground(cell);

				    var textRange = cell.Shape.TextFrame.TextRange;
				    SetFontProperties(textRange, fontName, fontNameFarEast, fontSize, MsoTriState.msoFalse, txtColor);

				    if(autonum && !string.IsNullOrEmpty(textRange.Text.Trim()))
				    {
					    SmartNumberFormat(textRange, decimalPlaces, negativeTextColor);
				    }

				    SetBorder(cell, NETOP.Enums.PpBorderType.ppBorderTop, thinBorderWidth, thinBorderColor, 0.5f);
                }
                catch { /* 忽略单个单元格错误 */ }
			}
		}

		private void ClearCellBackground(NETOP.Cell cell)
		{
			// WPS 兼容性限制：
			// 经过大量测试，WPS 的 COM API 不支持修改表格单元格的 Fill 属性。
			// 所有尝试（Fill.Visible, Fill.Solid, Fill.ForeColor 等）都会抛出 PropertySetCOMException。
			// 这是 WPS 的 API 限制，无法通过编程方式解决。
			// 
			// 解决方案：
			// 1. 在 WPS 中，用户需要手动清除表格背景（选中表格 -> 表格设计 -> 清除表格样式）
			// 2. 或者在 PowerPoint 中使用此功能（PowerPoint 完全支持）
			// 
			// 跳过此操作以保持性能（避免大量异常开销）
			return;
		}

		private void SetFontProperties(NETOP.TextRange textRange, string name, string nameFarEast, float size, MsoTriState bold, MsoThemeColorIndex color)
		{
			// WPS 兼容：使用探测标志，首次失败后跳过该操作类型
			if (_canSetFontName) { try { textRange.Font.Name = name; } catch { _canSetFontName = false; } }
			if (_canSetFontNameFarEast) { try { textRange.Font.NameFarEast = nameFarEast; } catch { _canSetFontNameFarEast = false; } }
			if (_canSetFontSize) { try { textRange.Font.Size = size; } catch { _canSetFontSize = false; } }
			if (_canSetFontBold) { try { textRange.Font.Bold = bold; } catch { _canSetFontBold = false; } }
			if (_canSetFontColor) { try { textRange.Font.Color.ObjectThemeColor = color; } catch { _canSetFontColor = false; } }
		}

		private void SmartNumberFormat(NETOP.TextRange textRange, int decimalPlaces, int negativeTextColor)
		{
			string text = textRange.Text;
			if(string.IsNullOrEmpty(text)) return;

			int length = text.Length;
			bool isPercentage = length > 0 && text[length - 1] == '%';
			string numStr = isPercentage ? text.Substring(0, length - 1).Trim() : text.Trim();

			if(string.IsNullOrEmpty(numStr) || (!char.IsDigit(numStr[0]) && numStr[0] != '-' && numStr[0] != '.' && numStr[0] != '+'))
				return;

			if(!double.TryParse(numStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
				return;

			string format = decimalPlaces switch
			{
				0 => "N0",
				1 => "N1",
				2 => "N2",
				3 => "N3",
				_ => "N" + decimalPlaces,
			};
			string formatted = num.ToString(format);
			if(isPercentage) formatted += "%";

			try { if(text != formatted) textRange.Text = formatted; } catch { }
			try { if(num < 0) textRange.Font.Color.RGB = negativeTextColor; } catch { }
		}

		private void SetBorder(NETOP.Cell cell, NETOP.Enums.PpBorderType borderType, float setWeight, object tcolor, float transparency = 0)
		{
			// WPS 兼容：使用探测标志，首次失败后跳过该操作类型
			try
			{
				var border = cell.Borders[borderType];

				if(setWeight <= 0f)
				{
					if (_canSetBorderWeight) { try { border.Weight = setWeight; } catch { _canSetBorderWeight = false; } }
					if (_canSetBorderVisible) { try { border.Visible = MsoTriState.msoFalse; } catch { _canSetBorderVisible = false; } }
				} 
				else
				{
					if (_canSetBorderVisible) { try { border.Visible = MsoTriState.msoTrue; } catch { _canSetBorderVisible = false; } }
					if (_canSetBorderWeight) { try { border.Weight = setWeight; } catch { _canSetBorderWeight = false; } }
					if (_canSetBorderTransparency) { try { border.Transparency = transparency; } catch { _canSetBorderTransparency = false; } }

					if (_canSetBorderColor) 
					{
						try 
						{
							if(tcolor is MsoThemeColorIndex themeColor) border.ForeColor.ObjectThemeColor = themeColor;
							else if(tcolor is int rgbColor) border.ForeColor.RGB = rgbColor;
							else border.ForeColor.RGB = 0;
						} catch { _canSetBorderColor = false; }
					}
				}
			}
			catch { /* 获取 Borders 集合失败 */ }
		}
	}
}
