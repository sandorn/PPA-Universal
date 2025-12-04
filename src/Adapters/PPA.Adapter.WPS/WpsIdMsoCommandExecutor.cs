using System;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Adapter.WPS
{
    /// <summary>
    /// WPS 平台下基于动态调用的 idMso 命令执行器实现。
    /// </summary>
    public class WpsIdMsoCommandExecutor : IIdMsoCommandExecutor
    {
        private readonly WPSContext _context;
        private readonly ILogger _logger;

        public WpsIdMsoCommandExecutor(IApplicationContext appContext, ILogger logger = null)
        {
            _context = appContext as WPSContext;
            _logger = logger ?? NullLogger.Instance;
        }

        public bool TryExecute(IApplicationContext appContext, string idMso)
        {
            var ctx = appContext as WPSContext ?? _context;
            if (ctx == null)
            {
                return false;
            }

            try
            {
                dynamic app = ctx.Application ?? ctx.NativeApplication;
                if (app == null)
                {
                    return false;
                }

                try
                {
                    // ClearMenu 作为抽象的“清除表格格式”命令，在 WPS 中直接映射为同名 idMso；
                    // 其它 idMso 原样透传。
                    app.CommandBars.ExecuteMso(idMso);
                    return true;
                }
                catch (Exception ex)
                {
                    _logger.LogDebug($"WPS ExecuteMso 失败: {idMso}, {ex.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug($"WPS ExecuteMso 访问 Application 失败: {idMso}, {ex.Message}");
                return false;
            }
        }
    }
}
