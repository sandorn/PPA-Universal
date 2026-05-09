using System.Drawing;
using System.Windows.Forms;

namespace PPA.Universal.ComAddIn
{
	/// <summary>
	/// COM 加载项内 Modal 窗体与 <see cref="MessageBox"/> 的统一标题、字体与常用控件工厂。
	/// </summary>
	internal static class ComDialogChrome
	{
		public const string Caption = "PPA Universal";

		/// <summary>与矩阵/线性复制等表单内容区内边距一致。</summary>
		public static readonly Padding ContentPadding = new Padding(16, 14, 16, 8);

		/// <summary>底部按钮行内边距。</summary>
		public static readonly Padding FooterPadding = new Padding(16, 6, 16, 12);

		/// <summary>标签列默认宽度（与 <see cref="DuplicateCopyDialogs"/> 一致）。</summary>
		public const int LabelColumnWidthPx = 96;

		/// <summary>确定/取消最小宽度（与业务对话框一致）。</summary>
		public static readonly Size DialogButtonMinSize = new Size(88, 28);

		/// <summary>功能主题 + 产品名，用于模式窗体标题栏。</summary>
		public static string SubstantiveTitle(string topic) => string.IsNullOrEmpty(topic) ? Caption : $"{topic} — {Caption}";

		/// <summary>
		/// 可缩放业务对话框的公共属性（字体、缩放、位置、边框）。
		/// </summary>
		public static void ApplyModalForm(Form form, bool sizable, Size minimumClientSize)
		{
			form.Font = SystemFonts.MessageBoxFont;
			form.AutoScaleMode = AutoScaleMode.Font;
			form.StartPosition = FormStartPosition.CenterScreen;
			form.ShowInTaskbar = false;
			form.MinimizeBox = false;
			form.MaximizeBox = sizable;
			form.FormBorderStyle = sizable ? FormBorderStyle.Sizable : FormBorderStyle.FixedDialog;
			form.MinimumSize = minimumClientSize;
			form.Padding = Padding.Empty;
		}

		/// <summary>标准确定/取消条（靠右、自下而上添加以符合 FlowDirection.RightToLeft）。</summary>
		public static FlowLayoutPanel CreateOkCancelFooter(out Button ok, out Button cancel)
		{
			var buttons = new FlowLayoutPanel
			{
				FlowDirection = FlowDirection.RightToLeft,
				Padding = FooterPadding,
				AutoSize = true,
				WrapContents = false,
				Dock = DockStyle.Bottom
			};
			ok = new Button
			{
				Text = "确定",
				DialogResult = DialogResult.OK,
				AutoSize = true,
				Padding = new Padding(16, 4, 16, 4),
				Margin = new Padding(8, 0, 0, 0),
				MinimumSize = DialogButtonMinSize
			};
			cancel = new Button
			{
				Text = "取消",
				DialogResult = DialogResult.Cancel,
				AutoSize = true,
				Padding = new Padding(16, 4, 16, 4),
				MinimumSize = DialogButtonMinSize
			};
			buttons.Controls.Add(ok);
			buttons.Controls.Add(cancel);
			return buttons;
		}

		public static void NotifyInfo(string text)
		{
			MessageBox.Show(text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		public static void NotifyWarning(string text)
		{
			MessageBox.Show(text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
		}

		public static void NotifyError(string text)
		{
			MessageBox.Show(text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
		}

		/// <summary>是/否确认；默认焦点在「否」以降低误触风险。</summary>
		public static DialogResult ConfirmWarning(string text)
		{
			return MessageBox.Show(
				text,
				Caption,
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2);
		}
	}
}
