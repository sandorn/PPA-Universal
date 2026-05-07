using System.Drawing;
using System.Windows.Forms;

namespace PPA.Universal.ComAddIn
{
	/// <summary>
	/// 全文查找替换：查找内容、替换为（两栏文本框）。
	/// </summary>
	internal static class FindReplaceDialog
	{
		/// <returns>用户点击确定时为 true，并输出非 null 字符串（可为空串）。</returns>
		public static bool TryShow(out string find, out string replace)
		{
			find = replace = null;

			using var f = new Form
			{
				Text = "查找替换",
				FormBorderStyle = FormBorderStyle.FixedDialog,
				StartPosition = FormStartPosition.CenterScreen,
				MinimizeBox = false,
				MaximizeBox = false,
				ShowInTaskbar = false,
				ClientSize = new Size(420, 140)
			};

			var layout = new TableLayoutPanel
			{
				Dock = DockStyle.Fill,
				ColumnCount = 2,
				RowCount = 2,
				Padding = new Padding(12, 12, 12, 4)
			};
			layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72f));
			layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

			var txtFind = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 2, 0, 4) };
			var txtReplace = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 2, 0, 4) };

			layout.Controls.Add(new Label { Text = "查找", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
			layout.Controls.Add(txtFind, 1, 0);
			layout.Controls.Add(new Label { Text = "替换为", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
			layout.Controls.Add(txtReplace, 1, 1);

			var buttons = new FlowLayoutPanel
			{
				Dock = DockStyle.Bottom,
				FlowDirection = FlowDirection.RightToLeft,
				Padding = new Padding(12, 4, 12, 8),
				AutoSize = true
			};
			var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, AutoSize = true };
			var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
			buttons.Controls.Add(ok);
			buttons.Controls.Add(cancel);

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
