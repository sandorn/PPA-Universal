using System;
using System.Runtime.CompilerServices;

namespace PPA.Logging
{
    /// <summary>
    /// 统一的日志接口
    /// </summary>
    public interface ILogger
    {
        /// <summary>
        /// 记录指定级别的日志
        /// </summary>
        void Log(LogLevel level, string message, Exception exception = null,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "");

        /// <summary>
        /// 记录信息级别的日志
        /// </summary>
        void LogInformation(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "");

        /// <summary>
        /// 记录警告级别的日志
        /// </summary>
        void LogWarning(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "");

        /// <summary>
        /// 记录调试级别的日志
        /// </summary>
        void LogDebug(string message,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "");

        /// <summary>
        /// 记录错误级别的日志
        /// </summary>
        void LogError(string message, Exception exception = null,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "");
    }
}
