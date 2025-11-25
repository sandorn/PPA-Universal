using System;
using System.Runtime.InteropServices;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS COM 互操作辅助类
    /// </summary>
    public static class WPSHelper
    {
        /// <summary>
        /// WPS 演示的 ProgID 列表（按优先级排序）
        /// </summary>
        public static readonly string[] WPSProgIDs = new[]
        {
            "KWPP.Application",      // WPS 演示 (Kingsoft)
            "Wpp.Application",       // WPS 演示 (另一种 ProgID)
            "WPS.Application",       // WPS 通用
            "KPresentation.Application" // WPS 国际版
        };

        /// <summary>
        /// 检测 WPS 是否已安装
        /// </summary>
        public static bool IsWPSInstalled()
        {
            foreach (var progId in WPSProgIDs)
            {
                var type = Type.GetTypeFromProgID(progId);
                if (type != null) return true;
            }
            return false;
        }

        /// <summary>
        /// 获取可用的 WPS ProgID
        /// </summary>
        public static string GetAvailableProgID()
        {
            foreach (var progId in WPSProgIDs)
            {
                var type = Type.GetTypeFromProgID(progId);
                if (type != null) return progId;
            }
            return null;
        }

        /// <summary>
        /// 创建 WPS 应用程序实例
        /// </summary>
        public static dynamic CreateWPSApplication()
        {
            var progId = GetAvailableProgID();
            if (progId == null)
            {
                throw new InvalidOperationException("WPS 演示未安装或无法访问");
            }

            var type = Type.GetTypeFromProgID(progId);
            return Activator.CreateInstance(type);
        }

        /// <summary>
        /// 获取正在运行的 WPS 应用程序实例
        /// </summary>
        public static dynamic GetRunningWPSApplication()
        {
            foreach (var progId in WPSProgIDs)
            {
                try
                {
                    return Marshal.GetActiveObject(progId);
                }
                catch
                {
                    // 继续尝试下一个 ProgID
                }
            }
            return null;
        }

        /// <summary>
        /// 安全释放 COM 对象
        /// </summary>
        public static void SafeRelease(object comObject)
        {
            if (comObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch
                {
                    // 忽略释放错误
                }
            }
        }

        /// <summary>
        /// 检测给定的应用程序对象是否为 WPS
        /// </summary>
        public static bool IsWPSApplication(dynamic app)
        {
            if (app == null) return false;

            try
            {
                string name = app.Name;
                return name != null && 
                    (name.Contains("WPS") || 
                     name.Contains("Kingsoft") || 
                     name.Contains("金山"));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取 MsoTriState 等价值（用于 WPS）
        /// </summary>
        public static class TriState
        {
            public const int True = -1;      // msoTrue
            public const int False = 0;       // msoFalse
            public const int Mixed = -2;      // msoTriStateMixed
            public const int Toggle = -3;     // msoTriStateToggle
        }

        /// <summary>
        /// 边框类型常量
        /// </summary>
        public static class BorderType
        {
            public const int Left = 1;
            public const int Top = 2;
            public const int Right = 3;
            public const int Bottom = 4;
            public const int DiagonalDown = 5;
            public const int DiagonalUp = 6;
        }

        /// <summary>
        /// 线条样式常量
        /// </summary>
        public static class LineStyle
        {
            public const int Solid = 1;
            public const int Dash = 4;
            public const int RoundDot = 3;
            public const int DashDot = 5;
            public const int DashDotDot = 6;
        }
    }
}
