using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using PPA.Adapter.WPS;
using PPA.Core.Abstraction;

namespace PPA.Universal.Platform
{
    /// <summary>
    /// 运行时平台检测器
    /// 自动检测当前运行的是 PowerPoint 还是 WPS
    /// </summary>
    public static class PlatformDetector
    {
        /// <summary>
        /// 检测结果缓存
        /// </summary>
        private static PlatformInfo _cachedInfo;

        /// <summary>
        /// 检测当前平台
        /// </summary>
        public static PlatformInfo Detect()
        {
            if (_cachedInfo != null)
                return _cachedInfo;

            _cachedInfo = DetectInternal();
            return _cachedInfo;
        }

        /// <summary>
        /// 强制重新检测
        /// </summary>
        public static PlatformInfo Redetect()
        {
            _cachedInfo = null;
            return Detect();
        }

        private static PlatformInfo DetectInternal()
        {
            var info = new PlatformInfo();

            // 1. 检测 PowerPoint
            info.PowerPointInstalled = IsPowerPointInstalled();
            info.PowerPointRunning = IsPowerPointRunning();

            // 2. 检测 WPS
            info.WPSInstalled = WPSHelper.IsWPSInstalled();
            info.WPSRunning = IsWPSRunning();

            // 3. 确定当前活跃平台
            info.ActivePlatform = DetermineActivePlatform(info);

            return info;
        }

        /// <summary>
        /// 检测 PowerPoint 是否安装
        /// </summary>
        public static bool IsPowerPointInstalled()
        {
            try
            {
                var type = Type.GetTypeFromProgID("PowerPoint.Application");
                return type != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 检测 PowerPoint 是否正在运行
        /// </summary>
        public static bool IsPowerPointRunning()
        {
            try
            {
                var processes = Process.GetProcessesByName("POWERPNT");
                return processes.Length > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 检测 WPS 是否正在运行
        /// </summary>
        public static bool IsWPSRunning()
        {
            try
            {
                // WPS 演示的进程名可能是 wpp 或 wps
                var wppProcesses = Process.GetProcessesByName("wpp");
                var wpsProcesses = Process.GetProcessesByName("wps");
                return wppProcesses.Length > 0 || wpsProcesses.Length > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 尝试获取正在运行的 PowerPoint 实例
        /// </summary>
        public static object GetRunningPowerPoint()
        {
            try
            {
                return Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 尝试获取正在运行的 WPS 实例
        /// </summary>
        public static object GetRunningWPS()
        {
            return WPSHelper.GetRunningWPSApplication();
        }

        /// <summary>
        /// 确定当前活跃平台
        /// </summary>
        private static PlatformType DetermineActivePlatform(PlatformInfo info)
        {
            // 优先检测正在运行的应用
            if (info.PowerPointRunning)
                return PlatformType.PowerPoint;

            if (info.WPSRunning)
                return PlatformType.WPS;

            // 如果都没运行，根据安装情况决定
            if (info.PowerPointInstalled)
                return PlatformType.PowerPoint;

            if (info.WPSInstalled)
                return PlatformType.WPS;

            return PlatformType.Unknown;
        }

        /// <summary>
        /// 从应用程序对象检测平台类型
        /// </summary>
        public static PlatformType DetectFromApplication(object app)
        {
            if (app == null)
                return PlatformType.Unknown;

            try
            {
                // 尝试获取应用程序名称
                dynamic dynApp = app;
                string name = dynApp.Name;

                if (string.IsNullOrEmpty(name))
                    return PlatformType.Unknown;

                if (name.Contains("PowerPoint") || name.Contains("Microsoft"))
                    return PlatformType.PowerPoint;

                if (name.Contains("WPS") || name.Contains("Kingsoft") || name.Contains("金山"))
                    return PlatformType.WPS;
            }
            catch
            {
                // 尝试通过类型判断
                var typeName = app.GetType().FullName ?? string.Empty;

                if (typeName.Contains("PowerPoint"))
                    return PlatformType.PowerPoint;
            }

            return PlatformType.Unknown;
        }
    }

    /// <summary>
    /// 平台检测信息
    /// </summary>
    public class PlatformInfo
    {
        /// <summary>PowerPoint 是否安装</summary>
        public bool PowerPointInstalled { get; set; }

        /// <summary>PowerPoint 是否正在运行</summary>
        public bool PowerPointRunning { get; set; }

        /// <summary>WPS 是否安装</summary>
        public bool WPSInstalled { get; set; }

        /// <summary>WPS 是否正在运行</summary>
        public bool WPSRunning { get; set; }

        /// <summary>当前活跃平台</summary>
        public PlatformType ActivePlatform { get; set; }

        /// <summary>是否有可用平台</summary>
        public bool HasAvailablePlatform =>
            PowerPointInstalled || WPSInstalled;

        /// <summary>是否有正在运行的平台</summary>
        public bool HasRunningPlatform =>
            PowerPointRunning || WPSRunning;

        public override string ToString()
        {
            return $"Platform: {ActivePlatform}, " +
                   $"PPT(installed={PowerPointInstalled}, running={PowerPointRunning}), " +
                   $"WPS(installed={WPSInstalled}, running={WPSRunning})";
        }
    }
}
