using NetOffice.OfficeApi.Enums;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Manipulation
{
	/// <summary>
	/// 配置辅助类 提供配置值的转换和辅助方法
	/// </summary>
	public static class ConfigHelper
	{
		/// <summary>
		/// 将主题颜色名称转换为 MsoThemeColorIndex
		/// </summary>
		public static MsoThemeColorIndex GetThemeColorIndex(string themeColorName)
		{
			if(string.IsNullOrEmpty(themeColorName))
				return MsoThemeColorIndex.msoThemeColorDark1;

			return themeColorName.ToLower() switch
			{
				"dark1" => MsoThemeColorIndex.msoThemeColorDark1,
				"dark2" => MsoThemeColorIndex.msoThemeColorDark2,
				"light1" => MsoThemeColorIndex.msoThemeColorLight1,
				"light2" => MsoThemeColorIndex.msoThemeColorLight2,
				"accent1" => MsoThemeColorIndex.msoThemeColorAccent1,
				"accent2" => MsoThemeColorIndex.msoThemeColorAccent2,
				"accent3" => MsoThemeColorIndex.msoThemeColorAccent3,
				"accent4" => MsoThemeColorIndex.msoThemeColorAccent4,
				"accent5" => MsoThemeColorIndex.msoThemeColorAccent5,
				"accent6" => MsoThemeColorIndex.msoThemeColorAccent6,
				"hyperlink" => MsoThemeColorIndex.msoThemeColorHyperlink,
				"followedhyperlink" => MsoThemeColorIndex.msoThemeColorFollowedHyperlink,
				_ => MsoThemeColorIndex.msoThemeColorDark1
			};
		}

		/// <summary>
		/// 将段落对齐字符串转换为枚举
		/// </summary>
		public static NETOP.Enums.PpParagraphAlignment GetParagraphAlignment(string alignment)
		{
			if(string.IsNullOrEmpty(alignment))
				return NETOP.Enums.PpParagraphAlignment.ppAlignJustify;

			return alignment.ToLower() switch
			{
				"left" => NETOP.Enums.PpParagraphAlignment.ppAlignLeft,
				"center" => NETOP.Enums.PpParagraphAlignment.ppAlignCenter,
				"right" => NETOP.Enums.PpParagraphAlignment.ppAlignRight,
				"justify" => NETOP.Enums.PpParagraphAlignment.ppAlignJustify,
				"distribute" => NETOP.Enums.PpParagraphAlignment.ppAlignDistribute,
				_ => NETOP.Enums.PpParagraphAlignment.ppAlignJustify
			};
		}

		/// <summary>
		/// 将项目符号类型字符串转换为枚举
		/// </summary>
		public static NETOP.Enums.PpBulletType GetBulletType(string bulletType)
		{
			if(string.IsNullOrEmpty(bulletType))
				return NETOP.Enums.PpBulletType.ppBulletUnnumbered;

			return bulletType.ToLower() switch
			{
				"none" => NETOP.Enums.PpBulletType.ppBulletNone,
				"numbered" => NETOP.Enums.PpBulletType.ppBulletNumbered,
				"unnumbered" => NETOP.Enums.PpBulletType.ppBulletUnnumbered,
				"picture" => NETOP.Enums.PpBulletType.ppBulletPicture,
				_ => NETOP.Enums.PpBulletType.ppBulletUnnumbered
			};
		}

		/// <summary>
		/// 将厘米转换为磅（1 厘米 = 28.35 磅）
		/// </summary>
		public static float CmToPoints(float cm)
		{
			return cm*28.35f;
		}

		/// <summary>
		/// 将主题颜色名称转换为整数索引值（用于新架构）
		/// </summary>
		public static int? GetThemeColorIndexValue(string themeColorName)
		{
			if(string.IsNullOrEmpty(themeColorName))
				return null;

			return themeColorName.ToLower() switch
			{
				"dark1" => 13,   // msoThemeColorDark1
				"dark2" => 14,   // msoThemeColorDark2
				"light1" => 15,  // msoThemeColorLight1
				"light2" => 16,  // msoThemeColorLight2
				"accent1" => 5,  // msoThemeColorAccent1
				"accent2" => 6,  // msoThemeColorAccent2
				"accent3" => 7,  // msoThemeColorAccent3
				"accent4" => 8,  // msoThemeColorAccent4
				"accent5" => 9,  // msoThemeColorAccent5
				"accent6" => 10, // msoThemeColorAccent6
				"hyperlink" => 11,
				"followedhyperlink" => 12,
				_ => null
			};
		}

		/// <summary>
		/// 获取主题颜色的默认 RGB 值（近似值，用于边框等）
		/// </summary>
		public static int GetThemeColorRgb(string themeColorName)
		{
			if(string.IsNullOrEmpty(themeColorName))
				return 0x000000;

			// 返回常见主题色的近似 RGB 值
			return themeColorName.ToLower() switch
			{
				"dark1" => 0x000000,    // 黑色
				"dark2" => 0x44546A,    // 深灰蓝
				"light1" => 0xFFFFFF,   // 白色
				"light2" => 0xE7E6E6,   // 浅灰
				"accent1" => 0x4472C4,  // 蓝色
				"accent2" => 0xED7D31,  // 橙色
				"accent3" => 0xA5A5A5,  // 灰色
				"accent4" => 0xFFC000,  // 金色
				"accent5" => 0x5B9BD5,  // 淡蓝
				"accent6" => 0x70AD47,  // 绿色
				_ => 0x000000
			};
		}
	}
}
