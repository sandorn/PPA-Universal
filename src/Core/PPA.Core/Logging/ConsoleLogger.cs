using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace PPA.Logging
{
    /// <summary>
    /// 简单的控制台日志实现
    /// </summary>
    public class ConsoleLogger : ILogger
    {
        private readonly LogLevel _minLevel;

        public ConsoleLogger(LogLevel minLevel = LogLevel.Information)
        {
            _minLevel = minLevel;
        }

        public void Log(LogLevel level, string message, Exception exception = null,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            if (level < _minLevel) return;

            var fileName = Path.GetFileName(callerFile);
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var levelStr = level.ToString().ToUpper().PadRight(5);

            Console.WriteLine($"[{timestamp}] [{levelStr}] [{fileName}:{callerMethod}] {message}");

            if (exception != null)
            {
                Console.WriteLine($"  Exception: {exception.Message}");
                Console.WriteLine($"  StackTrace: {exception.StackTrace}");
            }
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
