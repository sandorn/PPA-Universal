using System;
using PPA.Core.Abstraction;
using PPA.Logging;
using PPA.Universal.Integration;
using PPA.Legacy.Bridge;
using NETOP = NetOffice.PowerPointApi;
#if VSTO40
using MSOP = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
#endif

namespace PPA.PowerPoint.VSTO
{
    /// <summary>
    /// PPA PowerPoint 插件入口类（新架构版本）
    /// </summary>
    public partial class ThisAddIn
    {
        #region Private Fields

        private NETOP.Application _netApp;
        private ILogger _logger;
        private bool _disposed;

        #endregion

        #region Properties

        /// <summary>
        /// NetOffice PowerPoint 应用程序实例
        /// </summary>
        public NETOP.Application NetApp => _netApp;

        /// <summary>
        /// 原生 PowerPoint 应用程序实例
        /// </summary>
        public MSOP.Application NativeApp => this.Application;

        /// <summary>
        /// 应用程序上下文
        /// </summary>
        public IApplicationContext Context => UniversalIntegration.Context;

        /// <summary>
        /// 当前平台类型
        /// </summary>
        public PlatformType Platform => UniversalIntegration.Platform;

        /// <summary>
        /// 服务提供者
        /// </summary>
        public IServiceProvider ServiceProvider => UniversalIntegration.ServiceProvider;

        #endregion

        #region VSTO Events

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                // 初始化 NetOffice 包装器
                _netApp = new NETOP.Application(null, this.Application);

                // 使用新架构初始化
                UniversalIntegration.Initialize(_netApp, PlatformType.PowerPoint);

                // 获取日志服务
                _logger = UniversalIntegration.Logger;
                _logger.LogInformation("PPA PowerPoint 插件启动（新架构）");

                // 初始化服务桥接
                LegacyServiceBridge.Initialize(UniversalIntegration.ServiceProvider);

                // TODO: 初始化 Ribbon 和其他 UI 组件
                // InitializeRibbon();
                // InitializeKeyboardShortcuts();

                _logger.LogInformation("PPA 插件初始化完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PPA 启动错误: {ex.Message}");
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                _logger?.LogInformation("PPA PowerPoint 插件正在关闭");

                // 清理服务桥接
                LegacyServiceBridge.Cleanup();

                // 清理新架构资源
                UniversalIntegration.Cleanup();

                // 释放 NetOffice 实例
                if (_netApp != null)
                {
                    try { _netApp.Dispose(); } catch { }
                    _netApp = null;
                }

                _disposed = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PPA 关闭错误: {ex.Message}");
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// 获取服务
        /// </summary>
        public T GetService<T>() where T : class
        {
            return UniversalIntegration.GetService<T>();
        }

        #endregion

        #region VSTO Generated Code

        /// <summary>
        /// 必需的方法 - 请勿修改
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
