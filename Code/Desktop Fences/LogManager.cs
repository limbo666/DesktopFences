using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Desktop_Fences
{

    // Centralized logging manager for Desktop Fences application.
    // Handles log file rotation, filtering, and maintains compatibility with existing logging calls.

    public static class LogManager
    {
        // Use the same enums from FenceManager for full compatibility
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
            Settings
        }

        private static readonly object _logLock = new object();
        private static string _logFilePath;


        // Static constructor to initialize log file path

        static LogManager()
        {
            _logFilePath = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location),
                "Desktop_Fences.log");
        }


        // Main logging method - maintains exact same signature as FenceManager.Log for easy replacement
        // LogLevel:
        // Debug: Verbose logs (e.g., UI updates, icon caching, position changes).
        // Info: General operations (e.g., fence creation, property updates).
        // Warn: Non-critical issues (e.g., missing files, invalid settings).
        // Error: Critical failures (e.g., exceptions).
        // LogCategory:
        // General: Miscellaneous or uncategorized logs.
        // FenceCreation: Logs related to creating or initializing fences.
        // FenceUpdate: Logs for updating fence properties or content.
        // UI: Logs for UI interactions (e.g., mouse events, context menus).
        // IconHandling: Logs for icon loading, updating, or removal.
        // Error: Logs for errors or exceptions.
        // ImportExport: Logs for importing/exporting fences.
        // Settings: Logs for settings changes.
        // <param name="level">The severity level of the log message</param>
        // <param name="category">The functional category of the log message</param>
        // <param name="message">The log message content</param>
        public static void Log(LogLevel level, LogCategory category, string message)
        {
            try
            {
                // Check if logging is enabled via SettingsManager (maintains existing behavior)
                if (!SettingsManager.IsLogEnabled)
                    return;

                // Apply filtering based on SettingsManager settings (maintains existing behavior)
                LogLevel minLevel = SettingsManager.MinLogLevel;
                List<LogCategory> enabledCategories = SettingsManager.EnabledLogCategories;

                if (level < minLevel || !enabledCategories.Contains(category))
                    return;

                // Thread-safe logging with file rotation
                lock (_logLock)
                {
                    // Implement log rotation (maintains existing behavior)
                    RotateLogIfNeeded();

                    // Write log with same format as original
                    string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}][{category}] {message}\n";
                    System.IO.File.AppendAllText(_logFilePath, logMessage);
                }
            }
            catch (Exception ex)
            {
                // Fallback logging to prevent logging failures from breaking the application
                try
                {
                    string fallbackPath = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(_logFilePath),
                        "Desktop_Fences_Fallback.log");
                    string fallbackMessage = $"{DateTime.Now}: LogManager error: {ex.Message} | Original: [{level}][{category}] {message}\n";
                    System.IO.File.AppendAllText(fallbackPath, fallbackMessage);
                }
                catch
                {
                    // If even fallback fails, silently continue to prevent application crashes
                }
            }
        }


        // Debug logging method - placeholder for any existing DebugLog calls
        // Can be expanded based on your specific DebugLog requirements
        // <param name="context">Debug context information</param>
        // <param name="identifier">Object or fence identifier</param>
        // <param name="message">Debug message</param>
        public static void DebugLog(string context, string identifier, string message)
        {
            // Map to regular Log call with Debug level and appropriate category
            Log(LogLevel.Debug, LogCategory.General, $"[{context}][{identifier ?? "UNKNOWN"}] {message}");
        }


        // Overload for DebugLog with additional parameters for compatibility
        // <param name="context">Debug context information</param>
        // <param name="identifier">Object or fence identifier</param>
        // <param name="message">Debug message</param>
        // <param name="param1">Additional parameter 1</param>
        // <param name="param2">Additional parameter 2</param>
        // <param name="param3">Additional parameter 3</param>
        public static void DebugLog(string context, string identifier, string message, double param1, double param2, bool param3)
        {
            string detailedMessage = $"[{context}][{identifier ?? "UNKNOWN"}] {message} | P1:{param1:F1} | P2:{param2:F1} | P3:{param3}";
            Log(LogLevel.Debug, LogCategory.General, detailedMessage);
        }


        // Implements log file rotation with cleanup - maintains exact same behavior as original

        private static void RotateLogIfNeeded()
        {
            const long maxFileSize = 5 * 1024 * 1024; // 5MB - same as original

            if (!System.IO.File.Exists(_logFilePath))
                return;

            FileInfo fileInfo = new FileInfo(_logFilePath);
            if (fileInfo.Length <= maxFileSize)
                return;

            try
            {
                // Create archive with timestamp (same format as original)
                string archivePath = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(_logFilePath),
                    $"Desktop_Fences_{DateTime.Now:yyyyMMdd_HHmmss}.log");

                System.IO.File.Move(_logFilePath, archivePath);

                // Clean up old logs (keep last 5) - same logic as original
                CleanupOldLogs();
            }
            catch (Exception ex)
            {
                // Log rotation failure shouldn't break the application
                // Create fallback log entry
                try
                {
                    string fallbackPath = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(_logFilePath),
                        "Desktop_Fences_Fallback.log");
                    System.IO.File.AppendAllText(fallbackPath,
                        $"{DateTime.Now}: Log rotation failed: {ex.Message}\n");
                }
                catch
                {
                    // If even fallback fails, continue silently
                }
            }
        }


        // Cleans up old log files, keeping only the most recent 5 - maintains original behavior

        private static void CleanupOldLogs()
        {
            try
            {
                string logDirectory = System.IO.Path.GetDirectoryName(_logFilePath);
                var logFiles = Directory.GetFiles(logDirectory, "Desktop_Fences_*.log")
                    .OrderByDescending(f => System.IO.File.GetCreationTime(f))
                    .Skip(5); // Keep last 5, same as original

                foreach (var oldLog in logFiles)
                {
                    try
                    {
                        System.IO.File.Delete(oldLog);
                    }
                    catch (Exception ex)
                    {
                        // Log deletion failure to fallback log (same as original behavior)
                        string fallbackPath = System.IO.Path.Combine(logDirectory, "Desktop_Fences_Fallback.log");
                        System.IO.File.AppendAllText(fallbackPath,
                            $"{DateTime.Now}: Failed to delete old log {oldLog}: {ex.Message}\n");
                    }
                }
            }
            catch
            {
                // Cleanup failure shouldn't break the application
            }
        }


        // Gets the current log file path - utility method for external access
        // <returns>Full path to the current log file</returns>
        public static string GetLogFilePath()
        {
            return _logFilePath;
        }


        // Checks if logging is currently enabled based on settings
        // <returns>True if logging is enabled, false otherwise</returns>
        public static bool IsLoggingEnabled()
        {
            return SettingsManager.IsLogEnabled;
        }


        // Forces a log file rotation - utility method for maintenance

        public static void ForceLogRotation()
        {
            lock (_logLock)
            {
                RotateLogIfNeeded();
            }
        }


        // Convenience method for logging errors with exception details
        // <param name="category">The log category</param>
        // <param name="message">Error message</param>
        // <param name="ex">Exception details</param>
        public static void LogError(LogCategory category, string message, Exception ex)
        {
            string detailedMessage = $"{message}: {ex.Message}";
            if (SettingsManager.MinLogLevel == LogLevel.Debug)
            {
                detailedMessage += $"\nStackTrace: {ex.StackTrace}";
            }
            Log(LogLevel.Error, category, detailedMessage);
        }


        // Convenience method for logging with context information

        // <param name="level">Log level</param>
        // <param name="category">Log category</param>
        // <param name="context">Context information (e.g., method name, fence ID)</param>
        // <param name="message">Log message</param>
        public static void LogWithContext(LogLevel level, LogCategory category, string context, string message)
        {
            Log(level, category, $"[{context}] {message}");
        }
    }
}