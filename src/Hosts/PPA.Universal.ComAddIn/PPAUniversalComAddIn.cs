using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
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
        private static readonly string[] PowerPointAddinKeys =
        {
            @"Software\Microsoft\Office\PowerPoint\Addins\PPA.Universal.ComAddIn"
        };

        private const string FriendlyName = "PPA Universal";
        private const string Description = "PPA Universal COM Add-in";
        private static readonly string LogFilePath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PPA.Universal",
            "ComAddIn.log");
        private const string FallbackLogPath = @"C:\PPAUniversalComAddIn.log";

        static PPAUniversalComAddIn()
        {
            try
            {
                var traceLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Static ctor executed{Environment.NewLine}";
                System.IO.File.AppendAllText(@"C:\PPAUniversalTrace.txt", traceLine);
            }
            catch
            {
                // 忽略静态构造日志失败
            }
        }

        public PPAUniversalComAddIn()
        {
            try
            {
                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Instance ctor executed{Environment.NewLine}";
                System.IO.File.AppendAllText(@"C:\PPAUniversalTrace.txt", line);
            }
            catch
            {
                // ignore
            }
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                UniversalIntegration.Initialize(application);
                Log($"Initialize success. Platform={UniversalIntegration.Platform}");
            }
            catch (Exception ex)
            {
                Log($"Initialize failed: {ex}");
                WriteFallbackLog($"Initialize failed: {ex}");
                throw;
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                Log("Disconnect invoked, cleaning up.");
                UniversalIntegration.Cleanup();
            }
            catch (Exception ex)
            {
                Log($"Cleanup failed: {ex}");
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
                Log($"GetCustomUI called with ribbonId: {ribbonId}");
                
                // 从嵌入资源读取 Ribbon XML
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "PPA.Universal.ComAddIn.Resources.PPARibbon.xml";
                
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        Log($"Ribbon resource not found: {resourceName}");
                        // 尝试列出所有资源名称用于调试
                        var names = assembly.GetManifestResourceNames();
                        Log($"Available resources: {string.Join(", ", names)}");
                        return null;
                    }
                    
                    using (var reader = new StreamReader(stream))
                    {
                        var xml = reader.ReadToEnd();
                        Log($"Ribbon XML loaded, length: {xml.Length}");
                        return xml;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"GetCustomUI failed: {ex}");
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
            Log("Ribbon_OnLoad completed");
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

        // 参考选项
        public void OnAlignRefChanged(IRibbonControl control, string selectedId, int selectedIndex)
            => _ribbonCallbacks?.OnAlignRefChanged(control, selectedId, selectedIndex);

        public int GetAlignRefIndex(IRibbonControl control)
            => _ribbonCallbacks?.GetAlignRefIndex(control) ?? 0;

        #endregion

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            foreach (var path in PowerPointAddinKeys)
            {
                RegisterOfficeAddinKey(Registry.CurrentUser, path);
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            foreach (var path in PowerPointAddinKeys)
            {
                Registry.CurrentUser.DeleteSubKeyTree(path, false);
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

        private static void Log(string message)
        {
            try
            {
                var directory = System.IO.Path.GetDirectoryName(LogFilePath);
                if (!System.IO.Directory.Exists(directory))
                {
                    System.IO.Directory.CreateDirectory(directory);
                }

                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                System.IO.File.AppendAllText(LogFilePath, line);
            }
            catch
            {
                // 忽略日志写入异常
            }
        }

        private static void WriteFallbackLog(string message)
        {
            try
            {
                var directory = System.IO.Path.GetDirectoryName(FallbackLogPath);
                if (!System.IO.Directory.Exists(directory))
                {
                    System.IO.Directory.CreateDirectory(directory);
                }

                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                System.IO.File.AppendAllText(FallbackLogPath, line);
            }
            catch
            {
                // 忽略
            }
        }

    }
}

