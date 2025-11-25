using System;
using System.Runtime.CompilerServices;

namespace PPA.Logging
{
    /// <summary>
    /// 空日志实现（不输出任何内容）
    /// </summary>
    public class NullLogger : ILogger
    {
        public static readonly NullLogger Instance = new NullLogger();

        private NullLogger() { }

        public void Log(LogLevel level, string message, Exception exception = null,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            // 不做任何事
        }

        public void LogInformation(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            // 不做任何事
        }

        public void LogWarning(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            // 不做任何事
        }

        public void LogDebug(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            // 不做任何事
        }

        public void LogError(string message, Exception exception = null,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            // 不做任何事
        }
    }
}
