using System;
using System.Runtime.InteropServices;
using PPA.WPS;
using NETOP = NetOffice.PowerPointApi;
using NetOffice.OfficeApi.Enums;

namespace PPA.WPS.Debug
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== PPA WPS 功能适配测试 ===\n");
            Console.WriteLine("1. 表格格式化测试");
            Console.WriteLine("2. 文本格式化测试");
            Console.WriteLine("3. 图表格式化测试");
            Console.WriteLine("4. 形状对齐测试");
            Console.WriteLine("0. 退出");
            Console.Write("\n请选择: ");

            var key = Console.ReadKey();
            Console.WriteLine();

            if (key.KeyChar == '0') return;

            // 检测 WPS 运行状态
            try
            {
                var wpsApp = Marshal.GetActiveObject("KWpp.Application");
                Console.WriteLine("\nWPS 运行状态: 已连接");
            }
            catch
            {
                Console.WriteLine("\nWPS 未运行，请先打开 WPS 演示");
                Console.ReadKey();
                return;
            }

            // 初始化插件环境
            Console.WriteLine("初始化插件环境...");
            using var addIn = new WPSAddIn();
            addIn.StartupAuto();

            dynamic wpsApp2 = addIn.Bootstrapper.WPSApplication;

            switch (key.KeyChar)
            {
                case '1':
                    TestTableFormatting(wpsApp2);
                    break;
                case '2':
                    TestTextFormatting(wpsApp2);
                    break;
                case '3':
                    TestChartFormatting(wpsApp2);
                    break;
                case '4':
                    TestAlignOperations(wpsApp2);
                    break;
            }

            Console.WriteLine("\n测试结束，按任意键退出...");
            Console.ReadKey();
        }

        static void TestTableFormatting(dynamic wpsApp)
        {
            Console.WriteLine("\n=== 表格格式化测试 ===");
            Console.WriteLine("请选中包含表格的形状，然后按任意键...");
            Console.ReadKey(true);

            var config = new StubFormattingConfig();
            var helper = new PerformanceTableHelper(config);

            try
            {
                dynamic selection = wpsApp.ActiveWindow.Selection;
                if (selection.Type != 2) { Console.WriteLine("未选中形状!"); return; }

                dynamic shapes = selection.ShapeRange;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    try
                    {
                        dynamic shape = shapes[i];
                        if (shape.HasTable == -1)
                        {
                            var netTable = new NETOP.Table(null, shape.Table);
                            Console.Write($"格式化表格 ({netTable.Rows.Count}x{netTable.Columns.Count})... ");
                            var sw = System.Diagnostics.Stopwatch.StartNew();
                            helper.FormatTables(netTable);
                            sw.Stop();
                            Console.WriteLine($"完成! 耗时: {sw.ElapsedMilliseconds}ms");
                        }
                    }
                    catch (Exception ex) { Console.WriteLine($"出错: {ex.Message}"); }
                }
            }
            catch (Exception ex) { Console.WriteLine($"出错: {ex.Message}"); }
        }

        static void TestTextFormatting(dynamic wpsApp)
        {
            Console.WriteLine("\n=== 文本格式化测试 ===");
            Console.WriteLine("请选中包含文本的形状，然后按任意键...");
            Console.ReadKey(true);

            try
            {
                dynamic selection = wpsApp.ActiveWindow.Selection;
                if (selection.Type != 2) { Console.WriteLine("未选中形状!"); return; }

                dynamic shapes = selection.ShapeRange;
                Console.WriteLine($"选中形状数: {shapes.Count}");

                for (int i = 1; i <= shapes.Count; i++)
                {
                    try
                    {
                        dynamic shape = shapes[i];
                        int shapeType = shape.Type;
                        bool hasTable = false;
                        bool hasChart = false;
                        try { hasTable = shape.HasTable == -1; } catch { }
                        try { hasChart = shape.HasChart == -1; } catch { }
                        
                        string typeName = shapeType switch {
                            1 => "AutoShape", 3 => "Chart", 6 => "Group", 9 => "Line",
                            14 => "Placeholder", 17 => "TextBox", 19 => "Table", _ => $"Unknown({shapeType})"
                        };
                        Console.WriteLine($"\n--- 形状 {i} ({typeName}) ---");
                        
                        // 跳过不支持文本格式化的形状类型
                        if (hasTable) { Console.WriteLine("  [跳过] 表格需通过 Table.Cell 访问"); continue; }
                        if (hasChart) { Console.WriteLine("  [跳过] 图表需用 Chart API"); continue; }
                        if (shapeType == 9) { Console.WriteLine("  [跳过] 线条没有文本"); continue; }
                        
                        // 测试 TextFrame 属性
                        TestProperty(() => { var tf = shape.TextFrame; Console.WriteLine($"  TextFrame: OK"); });
                        TestProperty(() => { var tf = shape.TextFrame; tf.MarginTop = 5; Console.WriteLine($"  MarginTop: OK"); });
                        TestProperty(() => { var tf = shape.TextFrame; tf.MarginLeft = 5; Console.WriteLine($"  MarginLeft: OK"); });
                        
                        // 测试 Font 属性
                        TestProperty(() => { var f = shape.TextFrame.TextRange.Font; Console.WriteLine($"  Font: OK"); });
                        TestProperty(() => { shape.TextFrame.TextRange.Font.Name = "+mn-lt"; Console.WriteLine($"  Font.Name: OK"); });
                        TestProperty(() => { shape.TextFrame.TextRange.Font.Size = 11; Console.WriteLine($"  Font.Size: OK"); });
                        TestProperty(() => { shape.TextFrame.TextRange.Font.Bold = 0; Console.WriteLine($"  Font.Bold: OK"); });
                        TestProperty(() => { shape.TextFrame.TextRange.Font.Color.ObjectThemeColor = 13; Console.WriteLine($"  Font.Color: OK"); });
                        
                        // 测试段落属性
                        TestProperty(() => { var p = shape.TextFrame.TextRange.ParagraphFormat; Console.WriteLine($"  ParagraphFormat: OK"); });
                        TestProperty(() => { shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2; Console.WriteLine($"  Alignment: OK"); });
                        TestProperty(() => { shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0; Console.WriteLine($"  SpaceBefore: OK"); });
                        TestProperty(() => { shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0; Console.WriteLine($"  SpaceAfter: OK"); });
                    }
                    catch (Exception ex) { Console.WriteLine($"  形状出错: {ex.Message}"); }
                }
            }
            catch (Exception ex) { Console.WriteLine($"出错: {ex.Message}"); }
        }

        static void TestChartFormatting(dynamic wpsApp)
        {
            Console.WriteLine("\n=== 图表格式化测试 ===");
            Console.WriteLine("请选中包含图表的形状，然后按任意键...");
            Console.ReadKey(true);

            try
            {
                dynamic selection = wpsApp.ActiveWindow.Selection;
                if (selection.Type != 2) { Console.WriteLine("未选中形状!"); return; }

                dynamic shapes = selection.ShapeRange;
                Console.WriteLine($"选中形状数: {shapes.Count}");

                for (int i = 1; i <= shapes.Count; i++)
                {
                    try
                    {
                        dynamic shape = shapes[i];
                        Console.WriteLine($"\n--- 形状 {i} (Type: {shape.Type}) ---");
                        
                        // 检查是否是图表
                        bool hasChart = false;
                        TestProperty(() => { hasChart = shape.HasChart == -1; Console.WriteLine($"  HasChart: {hasChart}"); });
                        
                        if (!hasChart) { Console.WriteLine("  不是图表，跳过"); continue; }
                        
                        dynamic chart = shape.Chart;
                        
                        // 图表字体不支持主题占位符(+mn-lt)，需使用实际字体名
                        string chartFont = "微软雅黑"; // 或 "Arial"
                        
                        // 测试图表标题
                        TestProperty(() => { var t = chart.HasTitle; Console.WriteLine($"  HasTitle: {t}"); });
                        TestProperty(() => { if (chart.HasTitle) chart.ChartTitle.Font.Name = chartFont; Console.WriteLine($"  ChartTitle.Font.Name: OK ({chartFont})"); });
                        TestProperty(() => { if (chart.HasTitle) chart.ChartTitle.Font.Size = 14; Console.WriteLine($"  ChartTitle.Font.Size: OK"); });
                        
                        // 测试图例
                        TestProperty(() => { var l = chart.HasLegend; Console.WriteLine($"  HasLegend: {l}"); });
                        TestProperty(() => { if (chart.HasLegend) chart.Legend.Font.Name = chartFont; Console.WriteLine($"  Legend.Font.Name: OK ({chartFont})"); });
                        TestProperty(() => { if (chart.HasLegend) chart.Legend.Font.Size = 10; Console.WriteLine($"  Legend.Font.Size: OK"); });
                    }
                    catch (Exception ex) { Console.WriteLine($"  图表出错: {ex.Message}"); }
                }
            }
            catch (Exception ex) { Console.WriteLine($"出错: {ex.Message}"); }
        }

        static void TestProperty(Action action)
        {
            try { action(); }
            catch (Exception ex) { Console.WriteLine($"  FAILED: {ex.Message}"); }
        }

        static void TestAlignOperations(dynamic wpsApp)
        {
            Console.WriteLine("\n=== 形状对齐测试 ===");
            Console.WriteLine("请选中 2 个或更多形状，然后按任意键...");
            Console.ReadKey(true);

            try
            {
                dynamic selection = wpsApp.ActiveWindow.Selection;
                if (selection.Type != 2) { Console.WriteLine("未选中形状!"); return; }

                dynamic shapes = selection.ShapeRange;
                int count = shapes.Count;
                Console.WriteLine($"选中形状数: {count}");

                if (count < 2) { Console.WriteLine("需要至少选中 2 个形状!"); return; }

                // 记录原始位置
                float[] origLeft = new float[count + 1];
                float[] origTop = new float[count + 1];
                float[] origWidth = new float[count + 1];
                float[] origHeight = new float[count + 1];
                
                for (int i = 1; i <= count; i++)
                {
                    origLeft[i] = shapes[i].Left;
                    origTop[i] = shapes[i].Top;
                    origWidth[i] = shapes[i].Width;
                    origHeight[i] = shapes[i].Height;
                }
                Console.WriteLine("已记录原始位置\n");

                // 测试基本位置属性
                Console.WriteLine("--- 基本位置属性测试 ---");
                TestProperty(() => { var l = shapes[1].Left; Console.WriteLine($"  Read Left: OK ({l})"); });
                TestProperty(() => { var t = shapes[1].Top; Console.WriteLine($"  Read Top: OK ({t})"); });
                TestProperty(() => { var w = shapes[1].Width; Console.WriteLine($"  Read Width: OK ({w})"); });
                TestProperty(() => { var h = shapes[1].Height; Console.WriteLine($"  Read Height: OK ({h})"); });
                
                TestProperty(() => { shapes[2].Left = shapes[1].Left; Console.WriteLine($"  Set Left: OK"); });
                TestProperty(() => { shapes[2].Top = shapes[1].Top; Console.WriteLine($"  Set Top: OK"); });
                TestProperty(() => { shapes[2].Width = shapes[1].Width; Console.WriteLine($"  Set Width: OK"); });
                TestProperty(() => { shapes[2].Height = shapes[1].Height; Console.WriteLine($"  Set Height: OK"); });

                // 恢复位置
                for (int i = 1; i <= count; i++)
                {
                    shapes[i].Left = origLeft[i];
                    shapes[i].Top = origTop[i];
                    shapes[i].Width = origWidth[i];
                    shapes[i].Height = origHeight[i];
                }
                Console.WriteLine("\n已恢复原始位置");

                // 测试参考线访问
                Console.WriteLine("\n--- 参考线访问测试 ---");
                TestProperty(() => { 
                    var guides = wpsApp.ActivePresentation.Guides;
                    Console.WriteLine($"  Guides 集合: OK (Count: {guides.Count})");
                });

                // 测试形状填充和线条属性（SwapSize 用到）
                Console.WriteLine("\n--- 形状样式属性测试 ---");
                TestProperty(() => { var f = shapes[1].Fill; Console.WriteLine($"  Fill: OK"); });
                TestProperty(() => { var c = shapes[1].Fill.ForeColor.RGB; Console.WriteLine($"  Fill.ForeColor.RGB: OK ({c})"); });
                TestProperty(() => { var l = shapes[1].Line; Console.WriteLine($"  Line: OK"); });
                TestProperty(() => { var v = shapes[1].Line.Visible; Console.WriteLine($"  Line.Visible: OK ({v})"); });

                Console.WriteLine("\n✓ 对齐功能所需的基本属性测试完成");
            }
            catch (Exception ex) { Console.WriteLine($"出错: {ex.Message}"); }
        }
    }

    // --- 配置存根类 ---
    public class StubFormattingConfig : IDebugFormattingConfig
    {
        public IDebugTableConfig Table => new StubTableConfig();
    }

    public class StubTableConfig : IDebugTableConfig
    {
        public string StyleId => "{5940675A-B579-460E-94D1-54222C63F5DA}"; // WPS 无样式
        // 使用主题字体：+mn-lt (正文西文), +mn-ea (正文东亚文)
        public IDebugFontConfig HeaderRowFont => new StubFontConfig { Size = 12, Name = "+mn-lt", NameFarEast = "+mn-ea" };
        public IDebugFontConfig DataRowFont => new StubFontConfig { Size = 11, Name = "+mn-lt", NameFarEast = "+mn-ea" };
        
        public int HeaderRowBorderColor => 13; // dk1
        public int DataRowBorderColor => 13;
        
        public float HeaderRowBorderWidth => 1.5f;
        public float DataRowBorderWidth => 1.0f;
        
        public bool AutoNumberFormat => true;
        public int DecimalPlaces => 2;
        public int NegativeTextColor => 255; // Red

        public IDebugTableSettings TableSettings => new StubTableSettings();
    }

    public class StubFontConfig : IDebugFontConfig
    {
        public string Name { get; set; }
        public string NameFarEast { get; set; }
        public float Size { get; set; }
        public int ThemeColor { get; set; } = 13; // dk1
    }

    public class StubTableSettings : IDebugTableSettings
    {
        public bool FirstRow => true;
        public bool FirstCol => false;
        public bool LastRow => false;
        public bool LastCol => false;
        public bool HorizBanding => false;
        public bool VertBanding => false;
    }
}
