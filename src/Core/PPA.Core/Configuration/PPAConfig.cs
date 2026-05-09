// placeholder

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using System.Xml.Linq;
using PPA.Core.Abstraction;

namespace PPA.Core.Configuration
{
	[XmlRoot("PPAConfig")]
	public class PPAConfig
	{
		/// <summary>幻灯片尺寸等全局兜底（对齐参考、适配器读宽失败时等）。</summary>
		[XmlElement("Defaults")]
		public DefaultsConfig Defaults { get; set; }

		[XmlElement("Table")]
		public TableConfig Table { get; set; }

		[XmlElement("Text")]
		public TextConfig Text { get; set; }

		[XmlElement("Chart")]
		public ChartConfig Chart { get; set; }

		[XmlElement("GlassCard")]
		public GlassCardConfig GlassCard { get; set; }

		[XmlElement("Logging")]
		public LoggingConfig Logging { get; set; }

		/// <summary>矩阵/线性复制默认值（对话框初始值）。</summary>
		[XmlElement("Duplicate")]
		public DuplicateConfig Duplicate { get; set; }

		public static PPAConfig LoadOrCreate(string configPath)
		{
			if (string.IsNullOrWhiteSpace(configPath))
			{
				throw new ArgumentNullException(nameof(configPath));
			}

			var directory = Path.GetDirectoryName(configPath);
			if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
			{
				Directory.CreateDirectory(directory);
			}

			if (!File.Exists(configPath))
			{
				File.WriteAllText(configPath, GetDefaultXmlContent());
			}

			try
			{
				var config = TryLoadWithXDocument(configPath);
				if (config != null)
				{
					return config;
				}
			}
			catch (Exception ex)
			{
				TryLogConfigError(configPath, ex);
			}

			// 到这里说明：文件存在但解析失败（返回 null 或抛异常）
			// 使用默认内容重写配置文件，再尝试加载一次
			try
			{
				File.WriteAllText(configPath, GetDefaultXmlContent());

				var fallback = TryLoadWithXDocument(configPath);
				return fallback ?? LoadFromDefaultXmlString() ?? new PPAConfig();
			}
			catch (Exception ex)
			{
				// 即使重写默认配置或再次解析失败，也要记录日志，最终退回到模板默认值或空对象
				TryLogConfigError(configPath, ex);
				return LoadFromDefaultXmlString() ?? new PPAConfig();
			}
		}

		/// <summary>内存解析默认 XML 字符串，与磁盘模板内容一致（极端失败路径的最后兜底）。</summary>
		private static PPAConfig LoadFromDefaultXmlString()
		{
			try
			{
				return TryLoadWithXDocument(XDocument.Parse(GetDefaultXmlContent()));
			}
			catch
			{
				return null;
			}
		}

		private static string GetDefaultXmlContent()
		{
			return "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
				   "<PPAConfig>" +
				   "  <Defaults SlideWidthFallback=\"960\" SlideHeightFallback=\"540\" />" +
				   "  <Table StyleId=\"{2D5ABB26-0587-4C30-8999-92F81FD0307C}\" DataRowBorderWidth=\"1\" HeaderRowBorderWidth=\"1.75\" FinalRowBorderWidth=\"1.75\" " +
				   "         DataRowBorderColorIndex=\"13\" HeaderRowBorderColorIndex=\"13\" FinalRowBorderColorIndex=\"13\" " +
				   "         AutoNumberFormat=\"true\" DecimalPlaces=\"0\" NegativeTextColor=\"255\">" +
				   "    <DataRowFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"15\" Bold=\"false\" ThemeColorIndex=\"13\" />" +
				   "    <HeaderRowFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"15\" Bold=\"true\" ThemeColorIndex=\"13\" />" +
				   "    <TableSettings FirstRow=\"true\" FirstCol=\"false\" LastRow=\"false\" LastCol=\"false\" HorizBanding=\"false\" VertBanding=\"false\" />" +
				   "  </Table>" +
				   "  <Text LeftIndent=\"1\">" +
				   "    <Margins Top=\"0.2\" Bottom=\"0.2\" Left=\"0.5\" Right=\"0.5\" />" +
				   "    <Font Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"16\" Bold=\"true\" ThemeColor=\"Accent2\" />" +
				   "    <Paragraph Alignment=\"Justify\" WordWrap=\"true\" SpaceBefore=\"0\" SpaceAfter=\"0\" SpaceWithin=\"1.25\" FarEastLineBreakControl=\"true\" HangingPunctuation=\"true\" />" +
				   "    <Bullet Type=\"Unnumbered\" Character=\"9632\" FontName=\"+mn-lt\" RelativeSize=\"1\" ThemeColor=\"Dark1\" />" +
				   "  </Text>" +
				   "  <Chart>" +
				   "    <RegularFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"8\" Bold=\"false\" ThemeColor=\"Dark1\" />" +
				   "    <TitleFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"11\" Bold=\"true\" ThemeColor=\"Dark1\" />" +
				   "    <LegendFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"8\" Bold=\"false\" ThemeColor=\"Dark1\" />" +
				   "  </Chart>" +
				   "  <GlassCard BorderColorIndex=\"13\" BorderWidth=\"1.5\" CornerRadius=\"0.3\" " +
				   "             DefaultWidthRatio=\"0.6\" DefaultHeightRatio=\"0.25\" " +
				   "             PaddingTop=\"0.5\" PaddingBottom=\"0.5\" PaddingLeft=\"0.5\" PaddingRight=\"0.5\" " +
				   "             GradientDirection=\"45\" BlurRadius=\"10\">" +
				   "    <GradientStops>" +
				   "      <Stop Position=\"0\"   Opacity=\"0\"   />" +
				   "      <Stop Position=\"45\"  Opacity=\"80\"  />" +
				   "      <Stop Position=\"55\"  Opacity=\"90\"  />" +
				   "      <Stop Position=\"100\" Opacity=\"0\"   />" +
				   "    </GradientStops>" +
				   "    <TextStyle Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"16\" Bold=\"true\" ThemeColorIndex=\"13\" />" +
				   "  </GlassCard>" +
				   "  <Duplicate MatrixRows=\"3\" MatrixColumns=\"3\" MatrixRowSpacing=\"20\" MatrixColumnSpacing=\"20\" LinearCopyCount=\"5\" LinearSpacing=\"20\" LinearDirection=\"Horizontal\" />" +
				   "  <Logging EnableFileLogging=\"true\" MaxLogFiles=\"10\" MaxLogAgeDays=\"7\" MinimumLogLevel=\"Information\" RollingFileSizeMB=\"50\" />" +
				   "</PPAConfig>";
		}

		/// <summary>
		/// 使用 XDocument 从 XML 中解析完整的 PPAConfig；
		/// 解析错误将通过异常抛出，由调用方统一处理。
		/// </summary>
		private static PPAConfig TryLoadWithXDocument(string configPath)
		{
			try
			{
				var doc = XDocument.Load(configPath);
				return TryLoadWithXDocument(doc);
			}
			catch
			{
				return null;
			}
		}

		private static PPAConfig TryLoadWithXDocument(XDocument doc)
		{
			var root = doc?.Root;
			if (root == null || !string.Equals(root.Name.LocalName, "PPAConfig", StringComparison.OrdinalIgnoreCase))
			{
				return null;
			}

			var result = new PPAConfig();

			// 解析 Table 配置
			var tableElement = root.Element("Table");
			if (tableElement != null)
			{
				var table = new TableConfig();

				string GetAttr(XElement e, string name) => (string)e.Attribute(name);

				float ParseFloat(string v, float fallback)
				{
					if (string.IsNullOrWhiteSpace(v)) return fallback;
					return float.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var f)
						? f
						: fallback;
				}

				int ParseInt(string v, int fallback)
				{
					if (string.IsNullOrWhiteSpace(v)) return fallback;
					return int.TryParse(v, out var i) ? i : fallback;
				}

				bool ParseBool(string v, bool fallback)
				{
					if (string.IsNullOrWhiteSpace(v)) return fallback;
					return bool.TryParse(v, out var b) ? b : fallback;
				}

				table.StyleId = GetAttr(tableElement, "StyleId");
				table.DataRowBorderWidth = ParseFloat(GetAttr(tableElement, "DataRowBorderWidth"), table.DataRowBorderWidth);
				table.HeaderRowBorderWidth = ParseFloat(GetAttr(tableElement, "HeaderRowBorderWidth"), table.HeaderRowBorderWidth);
				table.FinalRowBorderWidth = ParseFloat(GetAttr(tableElement, "FinalRowBorderWidth"), table.FinalRowBorderWidth);
				table.DataRowBorderColorIndex = ParseInt(GetAttr(tableElement, "DataRowBorderColorIndex"), table.DataRowBorderColorIndex ?? 0);
				table.HeaderRowBorderColorIndex = ParseInt(GetAttr(tableElement, "HeaderRowBorderColorIndex"), table.HeaderRowBorderColorIndex ?? 0);
				table.FinalRowBorderColorIndex = ParseInt(GetAttr(tableElement, "FinalRowBorderColorIndex"), table.FinalRowBorderColorIndex ?? 0);
				table.AutoNumberFormat = ParseBool(GetAttr(tableElement, "AutoNumberFormat"), table.AutoNumberFormat);
				table.DecimalPlaces = ParseInt(GetAttr(tableElement, "DecimalPlaces"), table.DecimalPlaces);
				table.NegativeTextColor = ParseInt(GetAttr(tableElement, "NegativeTextColor"), table.NegativeTextColor);

				// 子节点：DataRowFont / HeaderRowFont
				var dataRowFontElement = tableElement.Element("DataRowFont");
				if (dataRowFontElement != null)
				{
					table.DataRowFont = ParseFontElement(dataRowFontElement);
				}

				var headerRowFontElement = tableElement.Element("HeaderRowFont");
				if (headerRowFontElement != null)
				{
					table.HeaderRowFont = ParseFontElement(headerRowFontElement);
				}

				// 子节点：TableSettings
				var tableSettingsElement = tableElement.Element("TableSettings");
				if (tableSettingsElement != null)
				{
					var settings = new TableSettingsConfig
					{
						FirstRow = ParseBool(GetAttr(tableSettingsElement, "FirstRow"), false),
						FirstCol = ParseBool(GetAttr(tableSettingsElement, "FirstCol"), false),
						LastRow = ParseBool(GetAttr(tableSettingsElement, "LastRow"), false),
						LastCol = ParseBool(GetAttr(tableSettingsElement, "LastCol"), false),
						HorizBanding = ParseBool(GetAttr(tableSettingsElement, "HorizBanding"), false),
						VertBanding = ParseBool(GetAttr(tableSettingsElement, "VertBanding"), false)
					};

					table.TableSettings = settings;
				}

				result.Table = table;
			}

			// 解析 Chart 配置
			var chartElement = root.Element("Chart");
			if (chartElement != null)
			{
				var chart = new ChartConfig();

				var regularFontElement = chartElement.Element("RegularFont");
				if (regularFontElement != null)
				{
					chart.RegularFont = ParseFontElement(regularFontElement);
				}

				var titleFontElement = chartElement.Element("TitleFont");
				if (titleFontElement != null)
				{
					chart.TitleFont = ParseFontElement(titleFontElement);
				}

				var legendFontElement = chartElement.Element("LegendFont");
				if (legendFontElement != null)
				{
					chart.LegendFont = ParseFontElement(legendFontElement);
				}

				result.Chart = chart;
			}

			// 解析 GlassCard 配置
			var glassCardElementRoot = root.Element("GlassCard");
			if (glassCardElementRoot != null)
			{
				var gc = ParseGlassCardElement(glassCardElementRoot);
				if (gc != null)
				{
					result.GlassCard = gc;
				}
			}

			// 解析 Duplicate（矩阵/线性复制默认值）
			var dupEl = root.Element("Duplicate");
			if (dupEl != null)
			{
				float PF(string v, float fb)
				{
					if (string.IsNullOrWhiteSpace(v)) return fb;
					return float.TryParse(v, System.Globalization.NumberStyles.Float,
						System.Globalization.CultureInfo.InvariantCulture, out var f) ? f : fb;
				}

				int PI(string v, int fb)
				{
					if (string.IsNullOrWhiteSpace(v)) return fb;
					return int.TryParse(v, out var i) ? i : fb;
				}

				var dir = (string)dupEl.Attribute("LinearDirection");
				var dupFb = new DuplicateConfig();
				if (string.IsNullOrWhiteSpace(dir)) dir = dupFb.LinearDirection;

				result.Duplicate = new DuplicateConfig
				{
					MatrixRows = PI((string)dupEl.Attribute("MatrixRows"), dupFb.MatrixRows),
					MatrixColumns = PI((string)dupEl.Attribute("MatrixColumns"), dupFb.MatrixColumns),
					MatrixRowSpacing = PF((string)dupEl.Attribute("MatrixRowSpacing"), dupFb.MatrixRowSpacing),
					MatrixColumnSpacing = PF((string)dupEl.Attribute("MatrixColumnSpacing"), dupFb.MatrixColumnSpacing),
					LinearCopyCount = PI((string)dupEl.Attribute("LinearCopyCount"), dupFb.LinearCopyCount),
					LinearSpacing = PF((string)dupEl.Attribute("LinearSpacing"), dupFb.LinearSpacing),
					LinearDirection = dir.Trim()
				};
			}

			var defaultsEl = root.Element("Defaults");
			if (defaultsEl != null)
			{
				float DF(string v, float fb)
				{
					if (string.IsNullOrWhiteSpace(v)) return fb;
					return float.TryParse(v, System.Globalization.NumberStyles.Float,
						System.Globalization.CultureInfo.InvariantCulture, out var f) ? f : fb;
				}

				var w = DF((string)defaultsEl.Attribute("SlideWidthFallback"), PpaConfigTemplateFallbacks.SlideWidthFallback);
				var h = DF((string)defaultsEl.Attribute("SlideHeightFallback"), PpaConfigTemplateFallbacks.SlideHeightFallback);
				result.Defaults = new DefaultsConfig
				{
					SlideWidthFallback = w > 0 ? w : PpaConfigTemplateFallbacks.SlideWidthFallback,
					SlideHeightFallback = h > 0 ? h : PpaConfigTemplateFallbacks.SlideHeightFallback
				};
			}

			var textRoot = root.Element("Text");
			if (textRoot != null)
			{
				result.Text = ParseTextConfigElement(textRoot);
			}

			var loggingRoot = root.Element("Logging");
			if (loggingRoot != null)
			{
				result.Logging = ParseLoggingConfigElement(loggingRoot);
			}

			return result;
		}

		private static TextConfig ParseTextConfigElement(XElement textElement)
		{
			string Ga(XElement e, string name) => (string)e.Attribute(name);

			float PF(string v, float fb)
			{
				if (string.IsNullOrWhiteSpace(v)) return fb;
				return float.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var f)
					? f
					: fb;
			}

			bool PB(string v, bool fb)
			{
				if (string.IsNullOrWhiteSpace(v)) return fb;
				return bool.TryParse(v, out var b) ? b : fb;
			}

			var text = new TextConfig
			{
				LeftIndent = PF(Ga(textElement, "LeftIndent"), 0f)
			};

			var marginsEl = textElement.Element("Margins");
			if (marginsEl != null)
			{
				text.Margins = new MarginsConfig
				{
					Top = PF(Ga(marginsEl, "Top"), 0f),
					Bottom = PF(Ga(marginsEl, "Bottom"), 0f),
					Left = PF(Ga(marginsEl, "Left"), 0f),
					Right = PF(Ga(marginsEl, "Right"), 0f)
				};
			}

			var fontEl = textElement.Element("Font");
			if (fontEl != null)
			{
				text.Font = ParseFontElement(fontEl);
			}

			var paraEl = textElement.Element("Paragraph");
			if (paraEl != null)
			{
				text.Paragraph = new ParagraphConfig
				{
					Alignment = Ga(paraEl, "Alignment"),
					WordWrap = PB(Ga(paraEl, "WordWrap"), true),
					SpaceBefore = PF(Ga(paraEl, "SpaceBefore"), 0f),
					SpaceAfter = PF(Ga(paraEl, "SpaceAfter"), 0f),
					SpaceWithin = PF(Ga(paraEl, "SpaceWithin"), 0f),
					FarEastLineBreakControl = PB(Ga(paraEl, "FarEastLineBreakControl"), true),
					HangingPunctuation = PB(Ga(paraEl, "HangingPunctuation"), true)
				};
			}

			var bulletEl = textElement.Element("Bullet");
			if (bulletEl != null)
			{
				int PI(string v, int fb)
				{
					if (string.IsNullOrWhiteSpace(v)) return fb;
					return int.TryParse(v, out var i) ? i : fb;
				}

				var bullet = new BulletConfig
				{
					Type = Ga(bulletEl, "Type"),
					Character = PI(Ga(bulletEl, "Character"), 0),
					FontName = Ga(bulletEl, "FontName"),
					RelativeSize = PF(Ga(bulletEl, "RelativeSize"), 1f)
				};
				var btci = Ga(bulletEl, "ThemeColorIndex");
				if (!string.IsNullOrWhiteSpace(btci) && int.TryParse(btci, out var bix))
					bullet.ThemeColorIndex = bix;
				else
				{
					var btc = ThemeColorIndexHelper.TryParse(Ga(bulletEl, "ThemeColor"));
					if (btc.HasValue)
						bullet.ThemeColorIndex = btc.Value;
				}

				text.Bullet = bullet;
			}

			return text;
		}

		private static LoggingConfig ParseLoggingConfigElement(XElement logElement)
		{
			string Ga(XElement e, string name) => (string)e.Attribute(name);

			bool PB(string v, bool fb)
			{
				if (string.IsNullOrWhiteSpace(v)) return fb;
				return bool.TryParse(v, out var b) ? b : fb;
			}

			int PI(string v, int fb)
			{
				if (string.IsNullOrWhiteSpace(v)) return fb;
				return int.TryParse(v, out var i) ? i : fb;
			}

			return new LoggingConfig
			{
				EnableFileLogging = PB(Ga(logElement, "EnableFileLogging"), true),
				MaxLogFiles = PI(Ga(logElement, "MaxLogFiles"), 10),
				MaxLogAgeDays = PI(Ga(logElement, "MaxLogAgeDays"), 7),
				MinimumLogLevel = Ga(logElement, "MinimumLogLevel") ?? "Information",
				RollingFileSizeMB = PI(Ga(logElement, "RollingFileSizeMB"), 50)
			};
		}

		/// <summary>
		/// 从 GlassCard XML 节点解析 GlassCardConfig。
		/// </summary>
		private static GlassCardConfig ParseGlassCardElement(XElement glassCardElement)
		{
			if (glassCardElement == null) return null;

			var gc = new GlassCardConfig();

			float ParseFloat(string v, float fallback)
			{
				if (string.IsNullOrWhiteSpace(v)) return fallback;
				return float.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var f)
					? f
					: fallback;
			}

			int ParseInt(string v, int fallback)
			{
				if (string.IsNullOrWhiteSpace(v)) return fallback;
				return int.TryParse(v, out var i) ? i : fallback;
			}

			// 读取 GlassCard 属性
			gc.BorderColorIndex = ParseInt((string)glassCardElement.Attribute("BorderColorIndex"), gc.BorderColorIndex);
			gc.BorderWidth = ParseFloat((string)glassCardElement.Attribute("BorderWidth"), gc.BorderWidth);
			gc.CornerRadius = ParseFloat((string)glassCardElement.Attribute("CornerRadius"), gc.CornerRadius);
			gc.DefaultWidthRatio = ParseFloat((string)glassCardElement.Attribute("DefaultWidthRatio"), gc.DefaultWidthRatio);
			gc.DefaultHeightRatio = ParseFloat((string)glassCardElement.Attribute("DefaultHeightRatio"), gc.DefaultHeightRatio);
			gc.PaddingTop = ParseFloat((string)glassCardElement.Attribute("PaddingTop"), gc.PaddingTop);
			gc.PaddingBottom = ParseFloat((string)glassCardElement.Attribute("PaddingBottom"), gc.PaddingBottom);
			gc.PaddingLeft = ParseFloat((string)glassCardElement.Attribute("PaddingLeft"), gc.PaddingLeft);
			gc.PaddingRight = ParseFloat((string)glassCardElement.Attribute("PaddingRight"), gc.PaddingRight);
			gc.GradientDirection = ParseFloat((string)glassCardElement.Attribute("GradientDirection"), gc.GradientDirection);
			gc.BlurRadius = ParseFloat((string)glassCardElement.Attribute("BlurRadius"), gc.BlurRadius);

			// 渐变停靠点
			var gradientStopsElement = glassCardElement.Element("GradientStops");
			if (gradientStopsElement != null)
			{
				var stops = new System.Collections.Generic.List<GlassGradientStopConfig>();
				foreach (var stopElement in gradientStopsElement.Elements("Stop"))
				{
					var stop = new GlassGradientStopConfig
					{
						Position = ParseFloat((string)stopElement.Attribute("Position"), 0f),
						Opacity = ParseFloat((string)stopElement.Attribute("Opacity"), 0f)
					};
					stops.Add(stop);
				}

				if (stops.Count > 0)
				{
					gc.GradientStops = stops.ToArray();
				}
			}

			// 文本样式：仅解析与玻璃卡片相关的基本字体配置
			var textStyleElement = glassCardElement.Element("TextStyle");
			if (textStyleElement != null)
			{
				var font = new FontConfig
				{
					Name = NormalizeThemeFontName((string)textStyleElement.Attribute("Name")),
					NameFarEast = NormalizeThemeFarEastFontName((string)textStyleElement.Attribute("NameFarEast")),
					Size = ParseFloat((string)textStyleElement.Attribute("Size"), 16f),
					Bold = bool.TryParse((string)textStyleElement.Attribute("Bold"), out var b) && b
				};

				var themeColorIndexAttr = (string)textStyleElement.Attribute("ThemeColorIndex");
				if (!string.IsNullOrWhiteSpace(themeColorIndexAttr) && int.TryParse(themeColorIndexAttr, out var themeIndex))
				{
					font.ThemeColorIndex = themeIndex;
				}
				else
				{
					var tc = ThemeColorIndexHelper.TryParse((string)textStyleElement.Attribute("ThemeColor"));
					if (tc.HasValue)
						font.ThemeColorIndex = tc.Value;
				}

				gc.TextStyle = font;
			}

			return gc;
		}

		/// <summary>
		/// 从字体节点解析 FontConfig（用于 Table.DataRowFont / HeaderRowFont 等）。
		/// </summary>
		private static FontConfig ParseFontElement(XElement fontElement)
		{
			if (fontElement == null) return null;

			float ParseFloat(string v, float fallback)
			{
				if (string.IsNullOrWhiteSpace(v)) return fallback;
				return float.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var f)
					? f
					: fallback;
			}

			bool ParseBool(string v, bool fallback)
			{
				if (string.IsNullOrWhiteSpace(v)) return fallback;
				return bool.TryParse(v, out var b) ? b : fallback;
			}

			int? ParseNullableInt(string v)
			{
				if (string.IsNullOrWhiteSpace(v)) return null;
				return int.TryParse(v, out var i) ? i : (int?)null;
			}

			var font = new FontConfig
			{
				Name = NormalizeThemeFontName((string)fontElement.Attribute("Name")),
				NameFarEast = NormalizeThemeFarEastFontName((string)fontElement.Attribute("NameFarEast")),
				Size = ParseFloat((string)fontElement.Attribute("Size"), 0f),
				Bold = ParseBool((string)fontElement.Attribute("Bold"), false)
			};

			var themeColorIndexAttr = (string)fontElement.Attribute("ThemeColorIndex");
			var parsedThemeIndex = ParseNullableInt(themeColorIndexAttr);
			if (parsedThemeIndex.HasValue)
			{
				font.ThemeColorIndex = parsedThemeIndex.Value;
			}
			else
			{
				var themeColorName = (string)fontElement.Attribute("ThemeColor");
				var fromName = ThemeColorIndexHelper.TryParse(themeColorName);
				if (fromName.HasValue)
					font.ThemeColorIndex = fromName.Value;
			}

			return font;
		}

		/// <summary>
		/// 记录配置加载相关错误到本地日志文件，避免影响主流程。
		/// </summary>
		private static void TryLogConfigError(string configPath, Exception ex)
		{
			try
			{
				var baseDir = Path.GetDirectoryName(configPath) ?? AppDomain.CurrentDomain.BaseDirectory;
				var errorLogPath = Path.Combine(baseDir, "PPAConfig.load-error.log");
				File.AppendAllText(errorLogPath,
					$"[{DateTime.Now:O}] Failed to load PPAConfig from '{configPath}': {ex}\r\n");
			}
			catch
			{
				// 忽略日志写入失败，避免影响主流程
			}
		}

		private static string NormalizeThemeFontName(string value)
		{
			if (string.IsNullOrWhiteSpace(value)) return value;
			var v = value.Trim();

			if (string.Equals(v, "Arial", StringComparison.OrdinalIgnoreCase) ||
				string.Equals(v, "+正文", StringComparison.OrdinalIgnoreCase))
			{
				return "+mn-lt";
			}

			return v;
		}

		private static string NormalizeThemeFarEastFontName(string value)
		{
			if (string.IsNullOrWhiteSpace(value)) return value;
			var v = value.Trim();

			if (string.Equals(v, "微软雅黑", StringComparison.OrdinalIgnoreCase) ||
				string.Equals(v, "+中文正文", StringComparison.OrdinalIgnoreCase))
			{
				return "+mn-ea";
			}

			return v;
		}
	}

	/// <summary>
	/// 全局兜底参数（无法从宿主读到演示文稿页面大小时使用）。
	/// </summary>
	public class DefaultsConfig
	{
		[XmlAttribute]
		public float SlideWidthFallback { get; set; } = 960f;

		[XmlAttribute]
		public float SlideHeightFallback { get; set; } = 540f;
	}

	public class TableConfig
	{
		[XmlAttribute]
		public string StyleId { get; set; }

		[XmlAttribute]
		public float DataRowBorderWidth { get; set; } = 1.0f;

		[XmlAttribute]
		public float HeaderRowBorderWidth { get; set; } = 1.75f;

		[XmlAttribute]
		public float FinalRowBorderWidth { get; set; } = 1.75f;

		[XmlAttribute("DataRowBorderColorIndex")]
		public int? DataRowBorderColorIndex { get; set; }

		[XmlAttribute("HeaderRowBorderColorIndex")]
		public int? HeaderRowBorderColorIndex { get; set; }

		[XmlAttribute("FinalRowBorderColorIndex")]
		public int? FinalRowBorderColorIndex { get; set; }

		[XmlAttribute("DataRowBorderColor")]
		public string LegacyDataRowBorderColor
		{
			get => null;
			set
			{
				var parsed = ThemeColorIndexHelper.TryParse(value);
				if (parsed.HasValue && !DataRowBorderColorIndex.HasValue)
				{
					DataRowBorderColorIndex = parsed.Value;
				}
			}
		}

		[XmlAttribute("HeaderRowBorderColor")]
		public string LegacyHeaderRowBorderColor
		{
			get => null;
			set
			{
				var parsed = ThemeColorIndexHelper.TryParse(value);
				if (parsed.HasValue && !HeaderRowBorderColorIndex.HasValue)
				{
					HeaderRowBorderColorIndex = parsed.Value;
				}
			}
		}

		[XmlAttribute("FinalRowBorderColor")]
		public string LegacyFinalRowBorderColor
		{
			get => null;
			set
			{
				var parsed = ThemeColorIndexHelper.TryParse(value);
				if (parsed.HasValue && !FinalRowBorderColorIndex.HasValue)
				{
					FinalRowBorderColorIndex = parsed.Value;
				}
			}
		}

		[XmlAttribute]
		public bool AutoNumberFormat { get; set; } = true;

		[XmlAttribute]
		public int DecimalPlaces { get; set; } = 0;

		[XmlAttribute]
		public int NegativeTextColor { get; set; } = 255;

		[XmlElement("DataRowFont")]
		public FontConfig DataRowFont { get; set; }

		[XmlElement("HeaderRowFont")]
		public FontConfig HeaderRowFont { get; set; }

		[XmlElement("TableSettings")]
		public TableSettingsConfig TableSettings { get; set; }
	}

	public class TextConfig
	{
		[XmlAttribute]
		public float LeftIndent { get; set; }

		[XmlElement("Margins")]
		public MarginsConfig Margins { get; set; }

		[XmlElement("Font")]
		public FontConfig Font { get; set; }

		[XmlElement("Paragraph")]
		public ParagraphConfig Paragraph { get; set; }

		[XmlElement("Bullet")]
		public BulletConfig Bullet { get; set; }
	}

	public class ChartConfig
	{
		[XmlElement("RegularFont")]
		public FontConfig RegularFont { get; set; }

		[XmlElement("TitleFont")]
		public FontConfig TitleFont { get; set; }

		[XmlElement("LegendFont")]
		public FontConfig LegendFont { get; set; }
	}

	public class GlassCardConfig
	{
		[XmlAttribute]
		public int BorderColorIndex { get; set; } = 13;

		[XmlAttribute]
		public float BorderWidth { get; set; } = 1.5f;

		/// <summary>圆角半径，单位：厘米（解释由具体适配器决定）</summary>
		[XmlAttribute]
		public float CornerRadius { get; set; } = 0.3f;

		/// <summary>默认宽度占页面宽度比例（无选中形状时使用）</summary>
		[XmlAttribute]
		public float DefaultWidthRatio { get; set; } = 0.6f;

		/// <summary>默认高度占页面高度比例（无选中形状时使用）</summary>
		[XmlAttribute]
		public float DefaultHeightRatio { get; set; } = 0.25f;

		[XmlAttribute]
		public float PaddingTop { get; set; } = 0.5f;

		[XmlAttribute]
		public float PaddingBottom { get; set; } = 0.5f;

		[XmlAttribute]
		public float PaddingLeft { get; set; } = 0.5f;

		[XmlAttribute]
		public float PaddingRight { get; set; } = 0.5f;

		/// <summary>渐变方向角度，单位：度</summary>
		[XmlAttribute]
		public float GradientDirection { get; set; } = 45f;

		/// <summary>模糊半径，仅在 PowerPoint 下尝试使用</summary>
		[XmlAttribute]
		public float BlurRadius { get; set; } = 10f;

		[XmlArray("GradientStops")]
		[XmlArrayItem("Stop")]
		public GlassGradientStopConfig[] GradientStops { get; set; }

		[XmlElement("TextStyle")]
		public FontConfig TextStyle { get; set; }
	}

	public class GlassGradientStopConfig
	{
		/// <summary>位置（0-100）</summary>
		[XmlAttribute]
		public float Position { get; set; }

		/// <summary>不透明度（0-100）</summary>
		[XmlAttribute]
		public float Opacity { get; set; }
	}

	/// <summary>
	/// 矩阵复制 / 线性复制对话框的默认值（来自 PPAConfig.xml）。
	/// </summary>
	public class DuplicateConfig
	{
		[XmlAttribute]
		public int MatrixRows { get; set; } = 3;

		[XmlAttribute]
		public int MatrixColumns { get; set; } = 3;

		[XmlAttribute]
		public float MatrixRowSpacing { get; set; } = 20f;

		[XmlAttribute]
		public float MatrixColumnSpacing { get; set; } = 20f;

		[XmlAttribute]
		public int LinearCopyCount { get; set; } = 5;

		[XmlAttribute]
		public float LinearSpacing { get; set; } = 20f;

		/// <summary>Horizontal 或 Vertical（不区分大小写）</summary>
		[XmlAttribute]
		public string LinearDirection { get; set; } = "Horizontal";
	}

	/// <summary>
	/// 日志相关配置
	/// </summary>
	public class LoggingConfig
	{
		/// <summary>
		/// 是否启用文件日志
		/// true：写入本地日志文件；false：仅使用默认控制台日志
		/// </summary>
		[XmlAttribute]
		public bool EnableFileLogging { get; set; }

		/// <summary>
		/// 最多保留的日志文件数量
		/// 超过该数量时会优先删除最早的日志文件
		/// </summary>
		[XmlAttribute]
		public int MaxLogFiles { get; set; }

		/// <summary>
		/// 日志文件最大保留天数
		/// 大于 0 时，会删除早于当前时间 N 天之前创建的日志文件
		/// 小于等于 0 时，不按时间限制，只按数量限制
		/// </summary>
		[XmlAttribute]
		public int MaxLogAgeDays { get; set; }

		/// <summary>
		/// 最小日志级别
		/// 取值为 Debug、Information、Warning、Error 等，对应 <see cref="LogLevel"/>
		/// 小于该级别的日志不会写入文件
		/// </summary>
		[XmlAttribute]
		public string MinimumLogLevel { get; set; }

		/// <summary>
		/// 单个日志文件的最大大小（单位：MB）
		/// 大于 0 时，以该值为滚动阈值；小于等于 0 时使用默认 50MB
		/// </summary>
		[XmlAttribute]
		public int RollingFileSizeMB { get; set; }
	}

	public class FontConfig
	{
		[XmlAttribute]
		public string Name { get; set; }

		[XmlAttribute]
		public string NameFarEast { get; set; }

		[XmlAttribute]
		public float Size { get; set; }

		[XmlAttribute]
		public bool Bold { get; set; }

		[XmlAttribute("ThemeColorIndex")]
		public int? ThemeColorIndex { get; set; }

		[XmlAttribute("ThemeColor")]
		public string LegacyThemeColor
		{
			get => null;
			set
			{
				var parsed = ThemeColorIndexHelper.TryParse(value);
				if (parsed.HasValue && !ThemeColorIndex.HasValue)
				{
					ThemeColorIndex = parsed.Value;
				}
			}
		}

		/// <summary>
		/// 映射为 <see cref="FontStyle"/>；空名称或字号≤0 时使用主题占位字体与 15pt（与默认模板中表格数据行字号一致）。
		/// </summary>
		public FontStyle ToFontStyle()
		{
			const float fallbackSize = 15f;
			const string fallbackLatin = "+mn-lt";
			const string fallbackEastAsia = "+mn-ea";
			return new FontStyle
			{
				Name = string.IsNullOrWhiteSpace(Name) ? fallbackLatin : Name,
				NameFarEast = string.IsNullOrWhiteSpace(NameFarEast) ? fallbackEastAsia : NameFarEast,
				Size = Size > 0 ? Size : fallbackSize,
				Bold = Bold,
				ThemeColorIndex = ThemeColorIndex
			};
		}
	}

	public class TableSettingsConfig
	{
		[XmlAttribute]
		public bool FirstRow { get; set; }

		[XmlAttribute]
		public bool FirstCol { get; set; }

		[XmlAttribute]
		public bool LastRow { get; set; }

		[XmlAttribute]
		public bool LastCol { get; set; }

		[XmlAttribute]
		public bool HorizBanding { get; set; }

		[XmlAttribute]
		public bool VertBanding { get; set; }
	}

	public class MarginsConfig
	{
		[XmlAttribute]
		public float Top { get; set; }

		[XmlAttribute]
		public float Bottom { get; set; }

		[XmlAttribute]
		public float Left { get; set; }

		[XmlAttribute]
		public float Right { get; set; }
	}

	public class ParagraphConfig
	{
		[XmlAttribute]
		public string Alignment { get; set; }

		[XmlAttribute]
		public bool WordWrap { get; set; }

		[XmlAttribute]
		public float SpaceBefore { get; set; }

		[XmlAttribute]
		public float SpaceAfter { get; set; }

		[XmlAttribute]
		public float SpaceWithin { get; set; }

		[XmlAttribute]
		public bool FarEastLineBreakControl { get; set; }

		[XmlAttribute]
		public bool HangingPunctuation { get; set; }
	}

	public class BulletConfig
	{
		[XmlAttribute]
		public string Type { get; set; }

		[XmlAttribute]
		public int Character { get; set; }

		[XmlAttribute]
		public string FontName { get; set; }

		[XmlAttribute]
		public float RelativeSize { get; set; }

		[XmlAttribute("ThemeColorIndex")]
		public int? ThemeColorIndex { get; set; }

		[XmlAttribute("ThemeColor")]
		public string LegacyThemeColor
		{
			get => null;
			set
			{
				var parsed = ThemeColorIndexHelper.TryParse(value);
				if (parsed.HasValue && !ThemeColorIndex.HasValue)
				{
					ThemeColorIndex = parsed.Value;
				}
			}
		}
	}

	/// <summary>
	/// 与 <see cref="PPAConfig.GetDefaultXmlContent"/> 中模板一致，供「节点缺失 / 解析失败 / 代码兜底」单点引用，避免魔法数与 XML 漂移。
	/// </summary>
	public static class PpaConfigTemplateFallbacks
	{
		public const float SlideWidthFallback = 960f;
		public const float SlideHeightFallback = 540f;

		/// <summary>默认模板中 <c>Text/Font</c>（Ribbon「文本字体」在 Text 整节缺失时的兜底）。</summary>
		public static FontStyle TextBoxRibbonFontStyle()
		{
			return new FontConfig
			{
				Name = "+mn-lt",
				NameFarEast = "+mn-ea",
				Size = 16,
				Bold = true,
				ThemeColorIndex = ThemeColorIndexHelper.TryParse("Accent2")
			}.ToFontStyle();
		}

		/// <summary>默认模板中 <c>Chart/TitleFont</c>。</summary>
		public static FontStyle ChartTitleFontStyle()
		{
			return new FontConfig
			{
				Name = "+mn-lt",
				NameFarEast = "+mn-ea",
				Size = 11,
				Bold = true,
				ThemeColorIndex = ThemeColorIndexHelper.TryParse("Dark1")
			}.ToFontStyle();
		}

		/// <summary>默认模板中 <c>Chart/LegendFont</c>。</summary>
		public static FontStyle ChartLegendFontStyle()
		{
			return new FontConfig
			{
				Name = "+mn-lt",
				NameFarEast = "+mn-ea",
				Size = 8,
				Bold = false,
				ThemeColorIndex = ThemeColorIndexHelper.TryParse("Dark1")
			}.ToFontStyle();
		}
	}

	internal static class ThemeColorIndexHelper
	{
		private static readonly Dictionary<string, int> ThemeColorIndexMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
		{
			["Dark1"] = 0,
			["Light1"] = 1,
			["Dark2"] = 2,
			["Light2"] = 3,
			["Accent1"] = 4,
			["Accent2"] = 5,
			["Accent3"] = 6,
			["Accent4"] = 7,
			["Accent5"] = 8,
			["Accent6"] = 9,
			["Hyperlink"] = 10,
			["FollowedHyperlink"] = 11,
			["Text1"] = 12,
			["Background1"] = 13,
			["Text2"] = 14,
			["Background2"] = 15
		};

		public static int? TryParse(string value)
		{
			if (string.IsNullOrWhiteSpace(value))
				return null;

			if (ThemeColorIndexMap.TryGetValue(value.Trim(), out var index))
			{
				return index;
			}

			if (int.TryParse(value, out var numeric))
			{
				return numeric;
			}

			return null;
		}
	}
}
