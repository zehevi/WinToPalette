using System;
using System.IO;

namespace WinToPalette.Logging
{
    public enum LogLevel
    {
        Debug = 0,
        Info = 1,
        Warning = 2,
        Error = 3
    }

    public interface ILogger
    {
        void LogInfo(string message);
        void LogWarning(string message);
        void LogError(string message);
        void LogDebug(string message);
    }

    public class FileLogger : ILogger
    {
        private readonly string _logFilePath;
        private readonly object _lockObject = new object();
        private readonly LogLevel _minimumLevel;

        public FileLogger(string logDirectory = null, LogLevel minimumLevel = LogLevel.Info)
        {
            if (string.IsNullOrEmpty(logDirectory))
            {
                logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "WinToPalette"
                );
            }

            Directory.CreateDirectory(logDirectory);
            _logFilePath = Path.Combine(logDirectory, $"WinToPalette_{DateTime.Now:yyyy_MM_dd}.log");
            _minimumLevel = minimumLevel;
        }

        public void LogInfo(string message) => Log(LogLevel.Info, "INFO", message);
        public void LogWarning(string message) => Log(LogLevel.Warning, "WARN", message);
        public void LogError(string message) => Log(LogLevel.Error, "ERROR", message);
        public void LogDebug(string message) => Log(LogLevel.Debug, "DEBUG", message);

        private void Log(LogLevel level, string levelLabel, string message)
        {
            if (level < _minimumLevel)
            {
                return;
            }

            try
            {
                lock (_lockObject)
                {
                    string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{levelLabel}] {message}";
                    File.AppendAllText(_logFilePath, logLine + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }
    }

    public class ConsoleLogger : ILogger
    {
        private readonly LogLevel _minimumLevel;

        public ConsoleLogger(LogLevel minimumLevel = LogLevel.Info)
        {
            _minimumLevel = minimumLevel;
        }

        public void LogInfo(string message) => Log(LogLevel.Info, "INFO", message, ConsoleColor.Green);
        public void LogWarning(string message) => Log(LogLevel.Warning, "WARN", message, ConsoleColor.Yellow);
        public void LogError(string message) => Log(LogLevel.Error, "ERROR", message, ConsoleColor.Red);
        public void LogDebug(string message) => Log(LogLevel.Debug, "DEBUG", message, ConsoleColor.Gray);

        private void Log(LogLevel level, string levelLabel, string message, ConsoleColor color)
        {
            if (level < _minimumLevel)
            {
                return;
            }

            var previousColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{levelLabel}] {message}");
            Console.ForegroundColor = previousColor;
        }
    }

    public class CompositeLogger : ILogger
    {
        private readonly ILogger[] _loggers;

        public CompositeLogger(params ILogger[] loggers)
        {
            _loggers = loggers ?? throw new ArgumentNullException(nameof(loggers));
        }

        public void LogInfo(string message)
        {
            foreach (var logger in _loggers)
                logger.LogInfo(message);
        }

        public void LogWarning(string message)
        {
            foreach (var logger in _loggers)
                logger.LogWarning(message);
        }

        public void LogError(string message)
        {
            foreach (var logger in _loggers)
                logger.LogError(message);
        }

        public void LogDebug(string message)
        {
            foreach (var logger in _loggers)
                logger.LogDebug(message);
        }
    }
}
