using System;
using PPA.Core.Abstraction;
using PPA.Logging;
using NETOP = NetOffice.PowerPointApi;
using NetOffice.OfficeApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 平台下基于 NetOffice 的 idMso 命令执行器实现。
    /// </summary>
    public class PowerPointIdMsoCommandExecutor : IIdMsoCommandExecutor
    {
        private readonly PowerPointContext _context;
        private readonly ILogger _logger;

        public PowerPointIdMsoCommandExecutor(IApplicationContext appContext, ILogger logger = null)
        {
            _context = appContext as PowerPointContext;
            _logger = logger ?? NullLogger.Instance;
        }

        public bool TryExecute(IApplicationContext appContext, string idMso)
        {
            // 仅在 PowerPoint 平台有效
            var ctx = appContext as PowerPointContext ?? _context;
            if (ctx == null)
            {
                return false;
            }

            try
            {
                var app = ctx.NetApplication;
                if (app == null)
                {
                    return false;
                }

                CommandBars commandBars = null;
                try
                {
                    commandBars = app.CommandBars;
                }
                catch
                {
                    // 部分环境可能不支持 CommandBars
                }

                if (commandBars == null)
                {
                    return false;
                }

                commandBars.ExecuteMso(idMso);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogDebug($"ExecuteMso 失败: {idMso}, {ex.Message}");
                return false;
            }
        }
    }
}
