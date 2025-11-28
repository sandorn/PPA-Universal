using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace PPA.Logging
{
    /// <summary>
    /// 文件日志实现
    /// 将日志写入到文件
    /// </summary>
    public class FileLogger : ILogger
    {
        private const long DefaultMaxLogFileSize = 50L * 1024 * 1024;
        private const int DefaultMaxLogFiles = 14;
        private const int CleanupBatchSize = 7;

        private readonly string _baseLogFilePath;
        private readonly string _logDirectory;
        private readonly string _baseFileName;
        private readonly string _fileExtension;
        private readonly LogLevel _minLevel;
        private readonly long _maxLogFileSize;
        private readonly int _maxLogFiles;
        private readonly int? _maxLogAgeDays;
        private readonly object _lock = new object();
        private string _currentLogFilePath;
        private DateTime _currentDate;
        private int _currentSequence;

        /// <summary>
        /// 创建文件日志实例
        /// </summary>
        /// <param name="logFilePath">日志文件路径</param>
        /// <param name="minLevel">最小日志级别</param>
        /// <param name="maxLogFiles">最多保留的日志文件数</param>
        /// <param name="maxLogAgeDays">日志文件最大保留天数（null 表示不按时间限制）</param>
        /// <param name="maxFileSizeBytes">单个日志文件的最大大小（字节），小于等于 0 时使用默认值</param>
        public FileLogger(string logFilePath, LogLevel minLevel = LogLevel.Information, int maxLogFiles = DefaultMaxLogFiles, int? maxLogAgeDays = null, long maxFileSizeBytes = DefaultMaxLogFileSize)
        {
            _baseLogFilePath = logFilePath ?? throw new ArgumentNullException(nameof(logFilePath));
            _minLevel = minLevel;
            _maxLogFiles = maxLogFiles > 0 ? maxLogFiles : DefaultMaxLogFiles;
            _maxLogAgeDays = maxLogAgeDays > 0 ? maxLogAgeDays : null;
            _maxLogFileSize = maxFileSizeBytes > 0 ? maxFileSizeBytes : DefaultMaxLogFileSize;
            _logDirectory = Path.GetDirectoryName(_baseLogFilePath) ?? AppDomain.CurrentDomain.BaseDirectory;
            _baseFileName = Path.GetFileNameWithoutExtension(_baseLogFilePath);
            _fileExtension = Path.GetExtension(_baseLogFilePath);

            // 确保日志目录存在
            try
            {
                if (!string.IsNullOrEmpty(_logDirectory) && !Directory.Exists(_logDirectory))
                {
                    Directory.CreateDirectory(_logDirectory);
                }
            }
            catch
            {
                // 忽略目录创建失败
            }

            CleanupOldLogs();
            InitializeLogFile();
        }

        public void Log(LogLevel level, string message, Exception exception = null,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            if (level < _minLevel) return;

            try
            {
                var fileName = Path.GetFileName(callerFile);
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var levelStr = level.ToString().ToUpper().PadRight(5);

                var logLine = $"[{timestamp}] [{levelStr}] [{fileName}:{callerMethod}] {message}";

                if (exception != null)
                {
                    logLine += $"{Environment.NewLine}  Exception: {exception.Message}";
                    if (!string.IsNullOrEmpty(exception.StackTrace))
                    {
                        logLine += $"{Environment.NewLine}  StackTrace: {exception.StackTrace}";
                    }
                }

                logLine += Environment.NewLine;

                lock (_lock)
                {
                    EnsureLogFile();
                    File.AppendAllText(_currentLogFilePath, logLine);
                }
            }
            catch
            {
                // 忽略日志写入失败，避免影响主流程
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

        private void InitializeLogFile()
        {
            _currentDate = DateTime.Now.Date;
            _currentSequence = GetAvailableSequence(_currentDate, 0);
            _currentLogFilePath = BuildLogFilePath(_currentDate, _currentSequence);
        }

        private void EnsureLogFile()
        {
            var today = DateTime.Now.Date;
            if (today != _currentDate)
            {
                _currentDate = today;
                _currentSequence = GetAvailableSequence(_currentDate, 0);
                _currentLogFilePath = BuildLogFilePath(_currentDate, _currentSequence);
                CleanupOldLogs();
                return;
            }

            if (GetFileSize(_currentLogFilePath) >= _maxLogFileSize)
            {
                _currentSequence = GetAvailableSequence(_currentDate, _currentSequence + 1);
                _currentLogFilePath = BuildLogFilePath(_currentDate, _currentSequence);
            }
        }

        private int GetAvailableSequence(DateTime date, int startIndex)
        {
            var index = startIndex;
            while (true)
            {
                var path = BuildLogFilePath(date, index);
                if (!File.Exists(path))
                {
                    return index;
                }

                var info = new FileInfo(path);
                if (info.Length < _maxLogFileSize)
                {
                    return index;
                }

                index++;
            }
        }

        private long GetFileSize(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path)) return 0;
                return new FileInfo(path).Length;
            }
            catch
            {
                return 0;
            }
        }

        private string BuildLogFilePath(DateTime date, int sequence)
        {
            var datePart = date.ToString("yyyyMMdd");
            var suffix = sequence > 0 ? $"_{sequence}" : string.Empty;
            var fileName = $"{_baseFileName}_{datePart}{suffix}{_fileExtension}";
            return Path.Combine(_logDirectory, fileName);
        }

        private void CleanupOldLogs()
        {
            try
            {
                if (!Directory.Exists(_logDirectory)) return;

                var searchPattern = $"{_baseFileName}_*{_fileExtension}";
                var files = Directory.GetFiles(_logDirectory, searchPattern, SearchOption.TopDirectoryOnly)
                    .Select(f => new FileInfo(f))
                    .OrderBy(f => f.CreationTimeUtc)
                    .ToList();

                if (files.Count == 0) return;

                int removed = 0;

                // 按时间限制清理
                if (_maxLogAgeDays.HasValue && _maxLogAgeDays.Value > 0)
                {
                    var threshold = DateTime.UtcNow.AddDays(-_maxLogAgeDays.Value);
                    foreach (var file in files.ToList())
                    {
                        if (removed >= CleanupBatchSize) break;

                        if (file.CreationTimeUtc < threshold)
                        {
                            try
                            {
                                file.Delete();
                                files.Remove(file);
                                removed++;
                            }
                            catch
                            {
                                // 忽略删除失败
                            }
                        }
                    }
                }

                // 按数量限制清理
                if (files.Count <= _maxLogFiles) return;

                var needRemove = Math.Min(CleanupBatchSize - removed, files.Count - _maxLogFiles);
                for (int i = 0; i < needRemove; i++)
                {
                    try
                    {
                        files[i].Delete();
                    }
                    catch
                    {
                        // 忽略删除失败
                    }
                }
            }
            catch
            {
                // 忽略清理失败
            }
        }
    }
}

