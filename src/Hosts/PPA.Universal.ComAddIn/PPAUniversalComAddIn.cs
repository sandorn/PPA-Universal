using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Extensibility;
using Microsoft.Office.Core;
using PPA.Core.Abstraction;
using PPA.Logging;
using PPA.Universal.Integration;
using Microsoft.Win32;

namespace PPA.Universal.ComAddIn
{
    /// <summary>
    /// 让 PowerPoint/WPS 识别的 COM 加载项入口
    /// 在 OnConnection 中初始化通用架构
    /// 实现 IRibbonExtensibility 提供自定义 Ribbon UI
    /// </summary>
    [ComVisible(true)]
    [Guid("C1BE96F0-86DF-4C00-9E51-09C989249C58")]
    [ProgId("PPA.Universal.ComAddIn")]
    public class PPAUniversalComAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private RibbonCallbacks _ribbonCallbacks;
        
        // PowerPoint 注册表路径
        private static readonly string[] PowerPointAddinKeys =
        {
            @"Software\Microsoft\Office\PowerPoint\Addins\PPA.Universal.ComAddIn"
        };

        // WPS 演示 注册表路径
        private static readonly string[] WPSAddinKeys =
        {
            @"Software\kingsoft\Office\WPP\Addins\PPA.Universal.ComAddIn"
        };

        // WPS 白名单路径
        private const string WPSWhitelistKey = @"Software\kingsoft\Office\WPP\AddinsWL";
        private const string WPSWhitelistValue = "PPA.Universal.ComAddIn";

        // 需要清理的错误路径
        private static readonly string[] ObsoleteKeys =
        {
            @"Software\kingsoft\Office\addins"
        };

        private const string FriendlyName = "PPA Universal";
        private const string Description = "PPA Universal COM Add-in";

        static PPAUniversalComAddIn()
        {
        }

        public PPAUniversalComAddIn()
        {
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                UniversalIntegration.Initialize(application);
                UniversalIntegration.Logger?.LogInformation($"Initialize success. Platform={UniversalIntegration.Platform}");
            }
            catch (Exception ex)
            {
                UniversalIntegration.Logger?.LogError($"Initialize failed: {ex}", ex);
                throw;
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                UniversalIntegration.Logger?.LogInformation("Disconnect invoked, cleaning up.");
                UniversalIntegration.Cleanup();
            }
            catch (Exception ex)
            {
                UniversalIntegration.Logger?.LogError($"Cleanup failed: {ex}", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        #region IRibbonExtensibility 实现

        /// <summary>
        /// 获取自定义 Ribbon UI XML
        /// </summary>
        public string GetCustomUI(string ribbonId)
        {
            try
            {
                // 检测平台（如果 UniversalIntegration 已初始化）
                PlatformType platform = PlatformType.Unknown;
                try
                {
                    platform = UniversalIntegration.Platform;
                }
                catch
                {
                    // UniversalIntegration 可能还未初始化，尝试通过 ribbonId 判断
                    if (ribbonId != null && (ribbonId.Contains("WPS") || ribbonId.Contains("Kingsoft")))
                    {
                        platform = PlatformType.WPS;
                    }
                }
                
                UniversalIntegration.Logger?.LogInformation($"GetCustomUI called with ribbonId: {ribbonId}, Platform: {platform}");
                
                // 从嵌入资源读取 Ribbon XML
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "PPA.Universal.ComAddIn.Resources.PPARibbon.xml";
                
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        UniversalIntegration.Logger?.LogWarning($"Ribbon resource not found: {resourceName}");
                        // 尝试列出所有资源名称用于调试
                        var names = assembly.GetManifestResourceNames();
                        UniversalIntegration.Logger?.LogDebug($"Available resources: {string.Join(", ", names)}");
                        return null;
                    }
                    
                    // 使用 UTF-8 编码读取，确保 WPS 和 PowerPoint 都能正确显示中文
                    // WPS 对编码更敏感，必须明确指定 UTF-8
                    using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
                    {
                        var xml = reader.ReadToEnd();
                        
                        // WPS 不支持 PowerPoint 的 imageMso 图标，需要移除以避免显示问号
                        if (platform == PlatformType.WPS)
                        {
                            UniversalIntegration.Logger?.LogInformation("WPS 平台检测到，移除 imageMso 属性以避免图标显示问题");
                            // 使用正则表达式移除所有 imageMso 属性
                            xml = System.Text.RegularExpressions.Regex.Replace(
                                xml,
                                @"\s+imageMso=""[^""]+""",
                                "",
                                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        }
                        
                        // 验证 XML 内容是否包含中文字符（用于调试）
                        if (xml.Contains("对齐") || xml.Contains("分布"))
                        {
                            UniversalIntegration.Logger?.LogInformation($"Ribbon XML loaded successfully, length: {xml.Length}, contains Chinese characters");
                        }
                        else
                        {
                            UniversalIntegration.Logger?.LogWarning($"Ribbon XML loaded but may have encoding issues, length: {xml.Length}");
                        }
                        
                        return xml;
                    }
                }
            }
            catch (Exception ex)
            {
                UniversalIntegration.Logger?.LogError($"GetCustomUI failed: {ex}", ex);
                return null;
            }
        }

        #endregion

        #region Ribbon 回调方法

        /// <summary>
        /// Ribbon 加载时回调
        /// </summary>
        public void Ribbon_OnLoad(IRibbonUI ribbon)
        {
            _ribbonCallbacks = new RibbonCallbacks();
            _ribbonCallbacks.Ribbon_OnLoad(ribbon);
            UniversalIntegration.Logger?.LogInformation("Ribbon_OnLoad completed");
        }

        // 对齐操作
        public void OnAlignLeft(IRibbonControl control) => _ribbonCallbacks?.OnAlignLeft(control);
        public void OnAlignRight(IRibbonControl control) => _ribbonCallbacks?.OnAlignRight(control);
        public void OnAlignTop(IRibbonControl control) => _ribbonCallbacks?.OnAlignTop(control);
        public void OnAlignBottom(IRibbonControl control) => _ribbonCallbacks?.OnAlignBottom(control);
        public void OnAlignCenterH(IRibbonControl control) => _ribbonCallbacks?.OnAlignCenterH(control);
        public void OnAlignCenterV(IRibbonControl control) => _ribbonCallbacks?.OnAlignCenterV(control);

        // 分布操作
        public void OnDistributeH(IRibbonControl control) => _ribbonCallbacks?.OnDistributeH(control);
        public void OnDistributeV(IRibbonControl control) => _ribbonCallbacks?.OnDistributeV(control);

        // 尺寸操作
        public void OnEqualWidth(IRibbonControl control) => _ribbonCallbacks?.OnEqualWidth(control);
        public void OnEqualHeight(IRibbonControl control) => _ribbonCallbacks?.OnEqualHeight(control);
        public void OnEqualSize(IRibbonControl control) => _ribbonCallbacks?.OnEqualSize(control);

        // 表格操作
        public void OnFormatThreeLineTable(IRibbonControl control) => _ribbonCallbacks?.OnFormatThreeLineTable(control);

        // 参考选项
        public void OnAlignRefChanged(IRibbonControl control, string selectedId, int selectedIndex)
            => _ribbonCallbacks?.OnAlignRefChanged(control, selectedId, selectedIndex);

        public int GetAlignRefIndex(IRibbonControl control)
            => _ribbonCallbacks?.GetAlignRefIndex(control) ?? 0;

        #endregion

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            // 清理错误的注册表项
            CleanupObsoleteKeys();

            // 注册 PowerPoint
            foreach (var path in PowerPointAddinKeys)
            {
                RegisterOfficeAddinKey(Registry.CurrentUser, path);
            }

            // 注册 WPS
            foreach (var path in WPSAddinKeys)
            {
                RegisterOfficeAddinKey(Registry.CurrentUser, path);
            }

            // 注册 WPS 白名单
            RegisterWPSWhitelist();
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            // 注销 PowerPoint
            foreach (var path in PowerPointAddinKeys)
            {
                Registry.CurrentUser.DeleteSubKeyTree(path, false);
            }

            // 注销 WPS
            foreach (var path in WPSAddinKeys)
            {
                Registry.CurrentUser.DeleteSubKeyTree(path, false);
            }

            // 移除 WPS 白名单
            UnregisterWPSWhitelist();

            // 清理错误的注册表项
            CleanupObsoleteKeys();
        }

        /// <summary>
        /// 注册 WPS 白名单
        /// </summary>
        private static void RegisterWPSWhitelist()
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(WPSWhitelistKey);
                key?.SetValue(WPSWhitelistValue, "", RegistryValueKind.String);
            }
            catch
            {
                // 忽略白名单注册失败
            }
        }

        /// <summary>
        /// 移除 WPS 白名单
        /// </summary>
        private static void UnregisterWPSWhitelist()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(WPSWhitelistKey, writable: true);
                key?.DeleteValue(WPSWhitelistValue, throwOnMissingValue: false);
            }
            catch
            {
                // 忽略白名单移除失败
            }
        }

        /// <summary>
        /// 清理错误创建的注册表项
        /// </summary>
        private static void CleanupObsoleteKeys()
        {
            foreach (var path in ObsoleteKeys)
            {
                try
                {
                    Registry.CurrentUser.DeleteSubKeyTree(path, throwOnMissingSubKey: false);
                }
                catch
                {
                    // 忽略清理失败
                }
            }
        }

        private static RegistryKey GetRegistryKey(RegistryHive hive)
        {
            try
            {
                return hive switch
                {
                    RegistryHive.CurrentUser => Registry.CurrentUser,
                    RegistryHive.LocalMachine => Registry.LocalMachine,
                    _ => null
                };
            }
            catch
            {
                return null;
            }
        }

        private static void RegisterOfficeAddinKey(RegistryKey root, string keyPath)
        {
            if (root == null || string.IsNullOrWhiteSpace(keyPath))
                return;

            try
            {
                using var key = root.CreateSubKey(keyPath);
                if (key == null) return;

                key.SetValue("FriendlyName", FriendlyName);
                key.SetValue("Description", Description);
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);

                TryRegisterWhitelistValue(root, keyPath);
            }
            catch
            {
                // 忽略没有权限写入 HKLM 的情况
            }
        }

        private static void TryRegisterWhitelistValue(RegistryKey root, string keyPath)
        {
            var marker = "AddinsWL\\";
            var index = keyPath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (index < 0) return;

            var parentPath = keyPath.Substring(0, index + marker.Length - 1);
            var valueName = keyPath.Substring(index + marker.Length);
            if (string.IsNullOrWhiteSpace(parentPath) || string.IsNullOrWhiteSpace(valueName)) return;

            try
            {
                using var wlKey = root.CreateSubKey(parentPath);
                wlKey?.SetValue(valueName, "1", RegistryValueKind.String);
            }
            catch
            {
                // 忽略白名单写入失败
            }
        }


    }
}

