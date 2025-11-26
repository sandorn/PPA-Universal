using System;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Universal.Platform;
using PPA.WPS;

namespace PPA.WPS.Debug
{
    /// <summary>
    /// WPS 调试控制台程序
    /// 用于测试 WPS 集成功能
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== PPA WPS 调试工具 ===\n");

            // 1. 检测平台
            Console.WriteLine("[1] 检测平台状态...");
            var platformInfo = PlatformDetector.Detect();
            Console.WriteLine($"    PowerPoint: 安装={platformInfo.PowerPointInstalled}, 运行={platformInfo.PowerPointRunning}");
            Console.WriteLine($"    WPS: 安装={platformInfo.WPSInstalled}, 运行={platformInfo.WPSRunning}");
            Console.WriteLine($"    当前平台: {platformInfo.ActivePlatform}\n");

            if (!platformInfo.WPSRunning)
            {
                Console.WriteLine("[!] 请先打开 WPS 演示 (wpp.exe)，然后按任意键继续...");
                Console.ReadKey();
                platformInfo = PlatformDetector.Redetect();
                
                if (!platformInfo.WPSRunning)
                {
                    Console.WriteLine("[错误] 未检测到运行中的 WPS 演示");
                    Console.ReadKey();
                    return;
                }
            }

            // 2. 初始化 WPS 插件
            Console.WriteLine("[2] 初始化 WPS 插件...");
            using var addIn = new WPSAddIn();
            
            try
            {
                addIn.StartupAuto();
                Console.WriteLine("    插件初始化成功\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"    [错误] 初始化失败: {ex.Message}");
                Console.ReadKey();
                return;
            }

            // 3. 测试服务
            Console.WriteLine("[3] 测试服务获取...");
            var tableService = addIn.ServiceProvider.GetService(typeof(ITableFormatService)) as ITableFormatService;
            Console.WriteLine($"    ITableFormatService: {(tableService != null ? "OK" : "未注册")}\n");

            // 4. 测试表格操作
            Console.WriteLine("[4] 测试表格操作...");
            TestTableFormatting(addIn);

            Console.WriteLine("\n=== 调试完成，按任意键退出 ===");
            Console.ReadKey();
        }

        static void TestTableFormatting(WPSAddIn addIn)
        {
            try
            {
                dynamic wpsApp = addIn.Bootstrapper.WPSApplication;
                
                // 检查是否有活动演示文稿
                if (wpsApp.Presentations.Count == 0)
                {
                    Console.WriteLine("    [!] 请打开一个包含表格的演示文稿");
                    return;
                }

                dynamic presentation = wpsApp.ActivePresentation;
                Console.WriteLine($"    当前文档: {presentation.Name}");

                // 检查选区
                dynamic selection = wpsApp.ActiveWindow.Selection;
                int selectionType = selection.Type;
                Console.WriteLine($"    选区类型: {selectionType}");

                if (selectionType == 2) // ShapeRange
                {
                    dynamic shapes = selection.ShapeRange;
                    Console.WriteLine($"    选中形状数: {shapes.Count}");

                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        dynamic shape = shapes[i];
                        bool hasTable = shape.HasTable;
                        Console.WriteLine($"    形状[{i}]: HasTable={hasTable}");

                        if (hasTable)
                        {
                            dynamic table = shape.Table;
                            Console.WriteLine($"        表格: {table.Rows.Count} 行 x {table.Columns.Count} 列");

                            // 使用新架构格式化表格
                            var tableService = addIn.ServiceProvider.GetService(typeof(ITableFormatService)) as ITableFormatService;
                            if (tableService != null)
                            {
                                var factory = new AdapterFactory();
                                var tableContext = factory.CreateTableContext(table, PlatformType.WPS);
                                
                                // 创建默认格式化选项
                                var options = new TableFormatOptions
                                {
                                    ApplyTableStyle = false,
                                    Settings = new TableSettings
                                    {
                                        FirstRow = true,
                                        HorizBanding = true
                                    }
                                };

                                Console.WriteLine("        开始格式化表格...");
                                tableService.FormatTable(tableContext, options);
                                Console.WriteLine("        格式化完成!");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("    [!] 请选中一个表格后重试");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"    [错误] {ex.Message}");
                Console.WriteLine($"    {ex.StackTrace}");
            }
        }
    }
}
