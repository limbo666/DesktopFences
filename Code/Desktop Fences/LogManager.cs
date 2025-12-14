using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Desktop_Fences
{
    public static class LogManager
    {
        public enum LogLevel { Debug, Info, Warn, Error }
        public enum LogCategory
        {
            General,
            FenceCreation,
            FenceUpdate,
            UI,
            IconHandling,
            Error,
            ImportExport,
            Settings,
            BackgroundValidation
        }

        private static readonly object _logLock = new object();
        private static string _logFilePath;
        private static int _rotationCheckCounter = 0; // Optimization counter

        static LogManager()
        {
            _logFilePath = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location),
                "Desktop_Fences.log");
        }

        public static void Log(LogLevel level, LogCategory category, string message)
        {
            try
            {
                if (!SettingsManager.IsLogEnabled) return;

                // Fast in-memory filtering
                if (level < SettingsManager.MinLogLevel) return;
                if (!SettingsManager.EnabledLogCategories.Contains(category)) return;

                lock (_logLock)
                {
                    // OPTIMIZATION: Only check disk for rotation every 500 logs
                    // This prevents checking FileInfo on every single mouse move event
                    _rotationCheckCounter++;
                    if (_rotationCheckCounter >= 500)
                    {
                        RotateLogIfNeeded();
                        _rotationCheckCounter = 0;
                    }

                    string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}][{category}] {message}\n";
                    System.IO.File.AppendAllText(_logFilePath, logMessage);
                }
            }
            catch (Exception ex)
            {
                // Silent fallback to prevent recursive crashes
                try
                {
                    string fallback = $"{DateTime.Now}: Log Error: {ex.Message}\n";
                    System.IO.File.AppendAllText(_logFilePath + ".err", fallback);
                }
                catch { }
            }
        }

        // Compatibility wrappers
        public static void DebugLog(string context, string identifier, string message)
        {
            Log(LogLevel.Debug, LogCategory.General, $"[{context}][{identifier ?? "UNKNOWN"}] {message}");
        }

        public static void DebugLog(string context, string identifier, string message, double p1, double p2, bool p3)
        {
            Log(LogLevel.Debug, LogCategory.General, $"[{context}][{identifier ?? "UNKNOWN"}] {message} | P1:{p1:F1} | P2:{p2:F1} | P3:{p3}");
        }

        private static void RotateLogIfNeeded()
        {
            const long maxFileSize = 5 * 1024 * 1024; // 5MB

            try
            {
                if (!System.IO.File.Exists(_logFilePath)) return;

                FileInfo fileInfo = new FileInfo(_logFilePath);
                if (fileInfo.Length <= maxFileSize) return;

                string archivePath = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(_logFilePath),
                    $"Desktop_Fences_{DateTime.Now:yyyyMMdd_HHmmss}.log");

                System.IO.File.Move(_logFilePath, archivePath);
                CleanupOldLogs();
            }
            catch { }
        }

        private static void CleanupOldLogs()
        {
            try
            {
                string logDirectory = System.IO.Path.GetDirectoryName(_logFilePath);
                var logFiles = Directory.GetFiles(logDirectory, "Desktop_Fences_*.log")
                    .OrderByDescending(f => System.IO.File.GetCreationTime(f))
                    .Skip(5); // Keep 5

                foreach (var oldLog in logFiles)
                {
                    try { System.IO.File.Delete(oldLog); } catch { }
                }
            }
            catch { }
        }

        public static string GetLogFilePath() => _logFilePath;
        public static bool IsLoggingEnabled() => SettingsManager.IsLogEnabled;

        public static void ForceLogRotation() { lock (_logLock) { RotateLogIfNeeded(); } }

        public static void LogError(LogCategory category, string message, Exception ex)
        {
            string msg = $"{message}: {ex.Message}";
            if (SettingsManager.MinLogLevel == LogLevel.Debug) msg += $"\n{ex.StackTrace}";
            Log(LogLevel.Error, category, msg);
        }

        public static void LogWithContext(LogLevel level, LogCategory category, string context, string message)
        {
            Log(level, category, $"[{context}] {message}");
        }
    }
}