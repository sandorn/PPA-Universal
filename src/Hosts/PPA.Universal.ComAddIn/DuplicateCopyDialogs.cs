using System;
using System.Drawing;
using System.Windows.Forms;
using PPA.Business.Abstractions;
using PPA.Core.Configuration;

namespace PPA.Universal.ComAddIn
{
	/// <summary>
	/// 矩阵复制 / 线性复制参数对话框；默认值来自 <see cref="DuplicateConfig"/>（PPAConfig.xml）。
	/// </summary>
	internal static class DuplicateCopyDialogs
	{
		public static bool TryShowMatrixDialog(DuplicateConfig cfg, out int rows, out int columns, out float rowSpacing, out float columnSpacing)
		{
			rows = columns = 0;
			rowSpacing = columnSpacing = 0;
			cfg ??= new DuplicateConfig();

			using var f = NewDialogShell("矩阵复制", 400, 268, 340, 240);
			var layout = NewGrid();

			var numRows = NewNumeric(cfg.MatrixRows, 1, 50);
			var numCols = NewNumeric(cfg.MatrixColumns, 1, 50);
			var numRowGap = NewNumericDecimal((decimal)cfg.MatrixRowSpacing, 0, 500);
			var numColGap = NewNumericDecimal((decimal)cfg.MatrixColumnSpacing, 0, 500);

			AddRow(layout, "行数", numRows);
			AddRow(layout, "列数", numCols);
			AddRow(layout, "行间距", numRowGap);
			AddRow(layout, "列间距", numColGap);
			AppendStretchRow(layout);

			var buttons = NewOkCancel(out var ok, out var cancel);

			f.Controls.Add(buttons);
			f.Controls.Add(layout);
			f.AcceptButton = ok;
			f.CancelButton = cancel;

			if (f.ShowDialog() != DialogResult.OK)
				return false;

			rows = (int)numRows.Value;
			columns = (int)numCols.Value;
			rowSpacing = (float)numRowGap.Value;
			columnSpacing = (float)numColGap.Value;

			cfg.MatrixRows = rows;
			cfg.MatrixColumns = columns;
			cfg.MatrixRowSpacing = rowSpacing;
			cfg.MatrixColumnSpacing = columnSpacing;
			return true;
		}

		public static bool TryShowLinearDialog(DuplicateConfig cfg, out int count, out float spacing, out LinearCopyDirection direction)
		{
			count = 0;
			spacing = 0;
			direction = LinearCopyDirection.Horizontal;
			cfg ??= new DuplicateConfig();

			using var f = NewDialogShell("线性复制", 400, 240, 340, 220);
			var layout = NewGrid();

			var numCount = NewNumeric(cfg.LinearCopyCount, 1, 200);
			var numSpacing = NewNumericDecimal((decimal)cfg.LinearSpacing, 0, 500);

			var dirPanel = new FlowLayoutPanel
			{
				FlowDirection = FlowDirection.LeftToRight,
				AutoSize = true,
				WrapContents = false,
				Dock = DockStyle.Fill,
				Margin = new Padding(0, 2, 0, 4),
				Padding = new Padding(0)
			};
			var rbH = new RadioButton
			{
				Text = "水平（向右）",
				AutoSize = true,
				Checked = !IsVertical(cfg.LinearDirection),
				Margin = new Padding(0, 6, 20, 0)
			};
			var rbV = new RadioButton
			{
				Text = "垂直（向下）",
				AutoSize = true,
				Checked = IsVertical(cfg.LinearDirection),
				Margin = new Padding(0, 6, 0, 0)
			};
			dirPanel.Controls.Add(rbH);
			dirPanel.Controls.Add(rbV);

			AddRow(layout, "复制份数", numCount);
			AddRow(layout, "间距", numSpacing);
			AddRow(layout, "方向", dirPanel);
			AppendStretchRow(layout);

			var buttons = NewOkCancel(out var ok, out var cancel);

			f.Controls.Add(buttons);
			f.Controls.Add(layout);
			f.AcceptButton = ok;
			f.CancelButton = cancel;

			if (f.ShowDialog() != DialogResult.OK)
				return false;

			count = (int)numCount.Value;
			spacing = (float)numSpacing.Value;
			direction = rbV.Checked ? LinearCopyDirection.Vertical : LinearCopyDirection.Horizontal;

			cfg.LinearCopyCount = count;
			cfg.LinearSpacing = spacing;
			cfg.LinearDirection = direction == LinearCopyDirection.Vertical ? "Vertical" : "Horizontal";
			return true;
		}

		private static Form NewDialogShell(string title, int w, int h, int minW, int minH)
		{
			var f = new Form { ClientSize = new Size(w, h) };
			ComDialogChrome.ApplyModalForm(f, true, new Size(minW, minH));
			f.Text = ComDialogChrome.SubstantiveTitle(title);
			return f;
		}

		private static TableLayoutPanel NewGrid()
		{
			var layout = new TableLayoutPanel
			{
				ColumnCount = 2,
				AutoSize = false,
				Dock = DockStyle.Fill,
				Padding = ComDialogChrome.ContentPadding
			};
			layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, ComDialogChrome.LabelColumnWidthPx));
			layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
			return layout;
		}

		private static FlowLayoutPanel NewOkCancel(out Button ok, out Button cancel) =>
			ComDialogChrome.CreateOkCancelFooter(out ok, out cancel);

		private static bool IsVertical(string linearDirection)
		{
			return string.Equals(linearDirection?.Trim(), "Vertical", StringComparison.OrdinalIgnoreCase);
		}

		private static void AddRow(TableLayoutPanel layout, string labelText, Control editor)
		{
			var row = layout.RowCount++;
			layout.RowStyles.Add(new System.Windows.Forms.RowStyle(SizeType.AutoSize));

			var lbl = new Label
			{
				Text = labelText,
				AutoSize = true,
				Anchor = AnchorStyles.Left,
				TextAlign = ContentAlignment.MiddleLeft,
				Margin = new Padding(0, 10, 12, 4)
			};

			editor.Margin = new Padding(0, 6, 0, 4);
			if (editor is FlowLayoutPanel)
				editor.Dock = DockStyle.Fill;
			else
			{
				editor.Dock = DockStyle.None;
				editor.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
			}

			layout.Controls.Add(lbl, 0, row);
			layout.Controls.Add(editor, 1, row);
		}

		/// <summary>底部弹性行：窗口拉高时留白在下方，避免控件被垂直拉长变形。</summary>
		private static void AppendStretchRow(TableLayoutPanel layout)
		{
			var row = layout.RowCount++;
			layout.RowStyles.Add(new System.Windows.Forms.RowStyle(SizeType.Percent, 100f));
			var filler = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
			layout.Controls.Add(filler, 0, row);
			layout.SetColumnSpan(filler, 2);
		}

		private static NumericUpDown NewNumeric(int value, int min, int max)
		{
			return new NumericUpDown
			{
				Minimum = min,
				Maximum = max,
				Value = Clamp(value, min, max),
				DecimalPlaces = 0,
				ThousandsSeparator = false,
				MinimumSize = new Size(120, 0)
			};
		}

		private static NumericUpDown NewNumericDecimal(decimal value, decimal min, decimal max)
		{
			return new NumericUpDown
			{
				Minimum = min,
				Maximum = max,
				DecimalPlaces = 2,
				Increment = 1,
				Value = ClampDecimal(value, min, max),
				ThousandsSeparator = false,
				MinimumSize = new Size(120, 0)
			};
		}

		private static int Clamp(int v, int min, int max)
		{
			if (v < min) return min;
			if (v > max) return max;
			return v;
		}

		private static decimal ClampDecimal(decimal v, decimal min, decimal max)
		{
			if (v < min) return min;
			if (v > max) return max;
			return v;
		}
	}
}
