using System;
using System.Runtime.CompilerServices;
using PPA.Logging;

namespace PPA.Legacy.Bridge
{
    /// <summary>
    /// 日志桥接器 - 将新架构日志接口适配到旧代码可用的形式
    /// </summary>
    public class LegacyLoggerBridge : ILogger
    {
        private readonly Action<string> _logAction;
        private readonly LogLevel _minLevel;

        public LegacyLoggerBridge(Action<string> logAction, LogLevel minLevel = LogLevel.Information)
        {
            _logAction = logAction ?? (_ => { });
            _minLevel = minLevel;
        }

        public void Log(LogLevel level, string message, Exception exception = null,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            if (level < _minLevel) return;

            var logMessage = $"[{DateTime.Now:HH:mm:ss}] [{level}] [{callerMethod}] {message}";
            if (exception != null)
            {
                logMessage += $"\n  Exception: {exception.Message}";
            }
            _logAction(logMessage);
        }

        public void LogInformation(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            Log(LogLevel.Information, message, null, callerMethod, callerFile);
        }

        public void LogWarning(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            Log(LogLevel.Warning, message, null, callerMethod, callerFile);
        }

        public void LogDebug(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            Log(LogLevel.Debug, message, null, callerMethod, callerFile);
        }

        public void LogError(string message, Exception exception = null,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            Log(LogLevel.Error, message, exception, callerMethod, callerFile);
        }
    }
}
