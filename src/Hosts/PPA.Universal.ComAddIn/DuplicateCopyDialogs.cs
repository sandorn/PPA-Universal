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

			using var f = NewDialogShell("矩阵复制", 360, 240);
			var layout = NewGrid();

			var numRows = NewNumeric(cfg.MatrixRows, 1, 50);
			var numCols = NewNumeric(cfg.MatrixColumns, 1, 50);
			var numRowGap = NewNumericDecimal((decimal)cfg.MatrixRowSpacing, 0, 500);
			var numColGap = NewNumericDecimal((decimal)cfg.MatrixColumnSpacing, 0, 500);

			AddRow(layout, "行数", numRows);
			AddRow(layout, "列数", numCols);
			AddRow(layout, "行间距", numRowGap);
			AddRow(layout, "列间距", numColGap);

			var buttons = NewOkCancel(out var ok, out var cancel);
			buttons.Dock = DockStyle.Bottom;

			layout.Dock = DockStyle.Fill;
			layout.Padding = new Padding(12, 12, 12, 4);

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

			using var f = NewDialogShell("线性复制", 360, 240);
			var layout = NewGrid();

			var numCount = NewNumeric(cfg.LinearCopyCount, 1, 200);
			var numSpacing = NewNumericDecimal((decimal)cfg.LinearSpacing, 0, 500);

			var dirPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false };
			var rbH = new RadioButton { Text = "水平（向右）", AutoSize = true, Checked = !IsVertical(cfg.LinearDirection) };
			var rbV = new RadioButton { Text = "垂直（向下）", AutoSize = true, Checked = IsVertical(cfg.LinearDirection) };
			dirPanel.Controls.Add(rbH);
			dirPanel.Controls.Add(rbV);

			AddRow(layout, "复制份数", numCount);
			AddRow(layout, "间距", numSpacing);
			AddRow(layout, "方向", dirPanel);

			var buttons = NewOkCancel(out var ok, out var cancel);
			buttons.Dock = DockStyle.Bottom;

			layout.Dock = DockStyle.Fill;
			layout.Padding = new Padding(12, 12, 12, 4);

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

		private static Form NewDialogShell(string title, int w, int h)
		{
			return new Form
			{
				Text = title,
				FormBorderStyle = FormBorderStyle.FixedDialog,
				StartPosition = FormStartPosition.CenterScreen,
				MinimizeBox = false,
				MaximizeBox = false,
				ShowInTaskbar = false,
				ClientSize = new Size(w, h)
			};
		}

		private static TableLayoutPanel NewGrid()
		{
			var layout = new TableLayoutPanel
			{
				ColumnCount = 2,
				AutoSize = true
			};
			layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(SizeType.Percent, 42f));
			layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(SizeType.Percent, 58f));
			return layout;
		}

		private static FlowLayoutPanel NewOkCancel(out Button ok, out Button cancel)
		{
			var buttons = new FlowLayoutPanel
			{
				FlowDirection = FlowDirection.RightToLeft,
				Padding = new Padding(12, 4, 12, 8),
				AutoSize = true,
				WrapContents = false
			};
			ok = new Button { Text = "确定", DialogResult = DialogResult.OK, AutoSize = true };
			cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
			buttons.Controls.Add(ok);
			buttons.Controls.Add(cancel);
			return buttons;
		}

		private static bool IsVertical(string linearDirection)
		{
			return string.Equals(linearDirection?.Trim(), "Vertical", StringComparison.OrdinalIgnoreCase);
		}

		private static void AddRow(TableLayoutPanel layout, string labelText, Control editor)
		{
			int row = layout.RowCount++;
			layout.RowStyles.Add(new System.Windows.Forms.RowStyle(SizeType.AutoSize));
			layout.Controls.Add(new Label { Text = labelText, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);
			layout.Controls.Add(editor, 1, row);
		}

		private static NumericUpDown NewNumeric(int value, int min, int max)
		{
			return new NumericUpDown
			{
				Minimum = min,
				Maximum = max,
				Value = Clamp(value, min, max),
				DecimalPlaces = 0,
				Width = 120
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
				Width = 120
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
