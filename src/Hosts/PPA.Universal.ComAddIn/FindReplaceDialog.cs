using System.Drawing;
using System.Windows.Forms;

namespace PPA.Universal.ComAddIn
{
	/// <summary>
	/// 全文查找替换：查找内容（单行）、替换为（多行，支持换行与自适应高度）。
	/// </summary>
	internal static class FindReplaceDialog
	{
		/// <returns>用户点击确定时为 true，并输出非 null 字符串（可为空串）。</returns>
		public static bool TryShow(out string find, out string replace)
		{
			find = replace = null;

			using var f = new Form();
			ComDialogChrome.ApplyModalForm(f, true, new Size(400, 220));
			f.Text = ComDialogChrome.SubstantiveTitle("查找替换");
			f.ClientSize = new Size(520, 300);

			var layout = new TableLayoutPanel
			{
				Dock = DockStyle.Fill,
				ColumnCount = 2,
				RowCount = 2,
				Padding = ComDialogChrome.ContentPadding,
				AutoSize = false
			};
			layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, ComDialogChrome.LabelColumnWidthPx));
			layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
			layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36f));
			layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

			var txtFind = new TextBox
			{
				Dock = DockStyle.Fill,
				Margin = new Padding(0, 4, 0, 8),
				Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
			};

			var txtReplace = new TextBox
			{
				Dock = DockStyle.Fill,
				Multiline = true,
				AcceptsReturn = true,
				WordWrap = true,
				ScrollBars = ScrollBars.Vertical,
				Margin = new Padding(0, 4, 0, 4),
				MinimumSize = new Size(0, 96)
			};

			using var tip = new ToolTip { ShowAlways = true };
			tip.SetToolTip(txtFind, "要查找的文本。");
			tip.SetToolTip(txtReplace, "替换后的内容；支持多行与换行（Enter 插入换行，不会直接关闭窗口）。");

			var lblFind = new Label
			{
				Text = "查找",
				AutoSize = false,
				Dock = DockStyle.Fill,
				TextAlign = ContentAlignment.MiddleLeft,
				Margin = new Padding(0, 4, 10, 0)
			};
			var lblReplace = new Label
			{
				Text = "替换为",
				AutoSize = false,
				Dock = DockStyle.Fill,
				TextAlign = ContentAlignment.TopLeft,
				Margin = new Padding(0, 8, 10, 0)
			};

			layout.Controls.Add(lblFind, 0, 0);
			layout.Controls.Add(txtFind, 1, 0);
			layout.Controls.Add(lblReplace, 0, 1);
			layout.Controls.Add(txtReplace, 1, 1);

			var buttons = ComDialogChrome.CreateOkCancelFooter(out var ok, out var cancel);

			f.Controls.Add(buttons);
			f.Controls.Add(layout);
			f.AcceptButton = ok;
			f.CancelButton = cancel;

			if (f.ShowDialog() != DialogResult.OK)
				return false;

			find = txtFind.Text ?? string.Empty;
			replace = txtReplace.Text ?? string.Empty;
			return true;
		}
	}
}
