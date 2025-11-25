using PPA.Logging;
using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace PPA.Core.Logging
{
	/// <summary>
	/// 基于 Profiler 的日志实现，实现新架构 ILogger 接口。
	/// </summary>
	public sealed class ProfilerLoggerAdapter : ILogger
	{
		public void Log(LogLevel level, string message, Exception exception = null,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
		{
			// 构建调用位置标识符（与 Profiler.LogMessage 的格式一致）
			string methodIdentifier = string.IsNullOrEmpty(callerFile)
				? callerMethod
				: $"{Path.GetFileNameWithoutExtension(callerFile)}.{callerMethod}";

			var logMessage = exception == null
				? message
				: $"{message} | Exception: {exception}";

			// 明确把 callerMethod/callerFile 传入 Profiler，避免 Profiler 自动采集到当前适配器方法名
			Profiler.LogMessage(logMessage,MapLevel(level),methodIdentifier,callerMethod,callerFile);
		}

		public void LogInformation(string message,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
			=> Log(LogLevel.Information,message,null,callerMethod,callerFile);

		public void LogWarning(string message,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
			=> Log(LogLevel.Warning,message,null,callerMethod,callerFile);

		public void LogDebug(string message,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
			=> Log(LogLevel.Debug,message,null,callerMethod,callerFile);

		public void LogError(string message,Exception exception = null,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
			=> Log(LogLevel.Error,message,exception,callerMethod,callerFile);

		private static string MapLevel(LogLevel level) => level switch
		{
			LogLevel.Debug => "DEBUG",
			LogLevel.Warning => "WARN",
			LogLevel.Error => "ERROR",
			LogLevel.Critical => "ERROR",
			_ => "INFO"
		};
	}
}
