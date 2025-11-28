// placeholder

using System;
using System.IO;
using System.Xml.Serialization;

namespace PPA.Core.Configuration
{
    [XmlRoot("PPAConfig")]
    public class PPAConfig
    {
        [XmlElement("Table")]
        public TableConfig Table { get; set; }

        [XmlElement("Text")]
        public TextConfig Text { get; set; }

        [XmlElement("Chart")]
        public ChartConfig Chart { get; set; }

        [XmlElement("Shortcuts")]
        public ShortcutsConfig Shortcuts { get; set; }

        [XmlElement("Logging")]
        public LoggingConfig Logging { get; set; }

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
                var serializer = new XmlSerializer(typeof(PPAConfig));
                using (var stream = File.OpenRead(configPath))
                {
                    return (PPAConfig)serializer.Deserialize(stream);
                }
            }
            catch
            {
                // 如果解析失败，返回一个新的配置实例，避免影响主流程
                return new PPAConfig();
            }
        }

        private static string GetDefaultXmlContent()
        {
            return "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                   "<PPAConfig>" +
                   "  <Table StyleId=\"{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}\" DataRowBorderWidth=\"1\" HeaderRowBorderWidth=\"1.75\" DataRowBorderColor=\"Accent2\" HeaderRowBorderColor=\"Accent1\" AutoNumberFormat=\"true\" DecimalPlaces=\"0\" NegativeTextColor=\"255\">" +
                   "    <DataRowFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"9\" Bold=\"false\" ThemeColor=\"Dark1\" />" +
                   "    <HeaderRowFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"10\" Bold=\"true\" ThemeColor=\"Dark1\" />" +
                   "    <TableSettings FirstRow=\"true\" FirstCol=\"false\" LastRow=\"false\" LastCol=\"false\" HorizBanding=\"false\" VertBanding=\"false\" />" +
                   "  </Table>" +
                   "  <Text LeftIndent=\"1\">" +
                   "    <Margins Top=\"0.2\" Bottom=\"0.2\" Left=\"0.5\" Right=\"0.5\" />" +
                   "    <Font Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"16\" Bold=\"true\" ThemeColor=\"Accent2\" />" +
                   "    <Paragraph Alignment=\"Justify\" WordWrap=\"true\" SpaceBefore=\"0\" SpaceAfter=\"0\" SpaceWithin=\"1.25\" FarEastLineBreakControl=\"true\" HangingPunctuation=\"true\" />" +
                   "    <Bullet Type=\"Unnumbered\" Character=\"9632\" FontName=\"Arial\" RelativeSize=\"1\" ThemeColor=\"Dark1\" />" +
                   "  </Text>" +
                   "  <Chart>" +
                   "    <RegularFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"8\" Bold=\"false\" ThemeColor=\"Dark1\" />" +
                   "    <TitleFont Name=\"+mn-lt\" NameFarEast=\"+mn-ea\" Size=\"11\" Bold=\"true\" ThemeColor=\"Dark1\" />" +
                   "  </Chart>" +
                   "  <Shortcuts FormatTables=\"1\" FormatText=\"2\" FormatChart=\"3\" CreateBoundingBox=\"4\" />" +
                   "  <Logging EnableFileLogging=\"true\" MaxLogFiles=\"10\" MaxLogAgeDays=\"7\" MinimumLogLevel=\"Information\" RollingFileSizeMB=\"50\" />" +
                   "</PPAConfig>";
        }
    }

    public class TableConfig
    {
        [XmlAttribute]
        public string StyleId { get; set; }

        [XmlAttribute]
        public float DataRowBorderWidth { get; set; }

        [XmlAttribute]
        public float HeaderRowBorderWidth { get; set; }

        [XmlAttribute]
        public string DataRowBorderColor { get; set; }

        [XmlAttribute]
        public string HeaderRowBorderColor { get; set; }

        [XmlAttribute]
        public bool AutoNumberFormat { get; set; }

        [XmlAttribute]
        public int DecimalPlaces { get; set; }

        [XmlAttribute]
        public int NegativeTextColor { get; set; }

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
    }

    public class ShortcutsConfig
    {
        [XmlAttribute]
        public int FormatTables { get; set; }

        [XmlAttribute]
        public int FormatText { get; set; }

        [XmlAttribute]
        public int FormatChart { get; set; }

        [XmlAttribute]
        public int CreateBoundingBox { get; set; }
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

        [XmlAttribute]
        public string ThemeColor { get; set; }
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

        [XmlAttribute]
        public string ThemeColor { get; set; }
    }
}
