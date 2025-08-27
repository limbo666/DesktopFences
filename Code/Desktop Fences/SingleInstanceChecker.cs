using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Desktop_Fences
{
    /// <summary>
    /// Single Instance Checker for Desktop Fences application
    /// Detects duplicate instances and coordinates with registry trigger system
    /// Compatible with existing project structure and logging patterns
    /// </summary>
    public static class SingleInstanceChecker
    {
        #region Private Fields

        private static string _currentProcessName;
        private static string _currentExecutablePath;

        #endregion

        #region Public Methods

        /// <summary>
        /// Initializes the single instance checker with current process information
        /// Should be called once during application startup, before any major initialization
        /// </summary>
        public static void Initialize()
        {
            try
            {
                Process currentProcess = Process.GetCurrentProcess();
                _currentProcessName = Path.GetFileNameWithoutExtension(currentProcess.ProcessName);
                _currentExecutablePath = Assembly.GetEntryAssembly()?.Location ?? currentProcess.MainModule?.FileName;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"SingleInstanceChecker: Initialized - Process: {_currentProcessName}, Path: {_currentExecutablePath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"SingleInstanceChecker: Error during initialization: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Checks if another instance of the application is already running
        /// Uses process name and executable path for reliable detection
        /// </summary>
        /// <returns>True if another instance is running, false if this is the first instance</returns>
        public static bool IsAnotherInstanceRunning()
        {
            try
            {
                if (string.IsNullOrEmpty(_currentProcessName))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                        "SingleInstanceChecker: Not initialized, calling Initialize() automatically");
                    Initialize();
                }

                Process currentProcess = Process.GetCurrentProcess();
                int currentProcessId = currentProcess.Id;

                // Get all processes with the same name
                Process[] processes = Process.GetProcessesByName(_currentProcessName);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"SingleInstanceChecker: Found {processes.Length} processes with name '{_currentProcessName}'");

                foreach (Process process in processes)
                {
                    try
                    {
                        // Skip our own process
                        if (process.Id == currentProcessId)
                            continue;

                        // Check if process is still running (not a zombie)
                        if (process.HasExited)
                            continue;

                        // Try to get the executable path for comparison
                        string processPath = null;
                        try
                        {
                            processPath = process.MainModule?.FileName;
                        }
                        catch (Exception)
                        {
                            // Access denied or other issue - skip detailed comparison
                            // But still count it as a potential duplicate since name matches
                        }

                        // If we can compare paths, do so for more reliable detection
                        if (!string.IsNullOrEmpty(processPath) && !string.IsNullOrEmpty(_currentExecutablePath))
                        {
                            bool isSameExecutable = string.Equals(processPath, _currentExecutablePath,
                                StringComparison.OrdinalIgnoreCase);

                            if (isSameExecutable)
                            {
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                    $"SingleInstanceChecker: Found another instance running - PID: {process.Id}, Path: {processPath}");
                                return true;
                            }
                        }
                        else
                        {
                            // Can't compare paths, but same process name - likely another instance
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                $"SingleInstanceChecker: Found potential duplicate process - PID: {process.Id} (path comparison unavailable)");
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                            $"SingleInstanceChecker: Error checking process {process.Id}: {ex.Message}");
                        // Continue checking other processes
                    }
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    "SingleInstanceChecker: No other instances found - this is the first instance");
                return false;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"SingleInstanceChecker: Error checking for other instances: {ex.Message}");
                // On error, assume no other instances (safer to continue than exit)
                return false;
            }
        }

        /// <summary>
        /// Handles the case where another instance is already running
        /// Writes trigger to registry and exits the current (duplicate) instance
        /// </summary>
        /// <returns>Always returns false (indicates this instance should exit)</returns>
        public static bool HandleDuplicateInstance()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    "SingleInstanceChecker: Handling duplicate instance - writing registry trigger");

                // Write trigger to activate effect in existing instance
                bool triggerWritten = RegistryHelper.WriteTrigger();

                if (triggerWritten)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "SingleInstanceChecker: Registry trigger written successfully - exiting duplicate instance");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                        "SingleInstanceChecker: Failed to write registry trigger - still exiting duplicate instance");
                }

                // Give the registry write a moment to complete
                System.Threading.Thread.Sleep(100);

                return false; // Always return false to indicate this instance should exit
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"SingleInstanceChecker: Error handling duplicate instance: {ex.Message}");
                return false; // Still exit on error
            }
        }

        /// <summary>
        /// Complete single instance check and handling
        /// Call this at the very beginning of application startup
        /// </summary>
        /// <returns>True if application should continue, false if it should exit</returns>
        public static bool CheckAndHandle()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    "SingleInstanceChecker: Starting single instance check");

                // Initialize if not already done
                if (string.IsNullOrEmpty(_currentProcessName))
                {
                    Initialize();
                }

                // Check for other instances
                bool anotherInstanceRunning = IsAnotherInstanceRunning();

                if (anotherInstanceRunning)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "SingleInstanceChecker: Another instance detected - handling as duplicate");
                    return HandleDuplicateInstance();
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "SingleInstanceChecker: This is the first instance - continuing startup");
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"SingleInstanceChecker: Critical error in CheckAndHandle: {ex.Message}");
                // On critical error, allow application to continue (safer than unexpected exit)
                return true;
            }
        }

        /// <summary>
        /// Gets information about current process for debugging
        /// </summary>
        /// <returns>Process information string</returns>
        public static string GetProcessInfo()
        {
            try
            {
                Process currentProcess = Process.GetCurrentProcess();
                return $"PID: {currentProcess.Id}, Name: {_currentProcessName}, Path: {_currentExecutablePath}";
            }
            catch (Exception ex)
            {
                return $"Error getting process info: {ex.Message}";
            }
        }

        /// <summary>
        /// Test method to simulate duplicate instance scenario
        /// Used for development/testing purposes only
        /// </summary>
        /// <returns>Test results</returns>
        public static string TestSingleInstanceLogic()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    "SingleInstanceChecker: Running test simulation");

                string results = $"Current Process Info: {GetProcessInfo()}\n";

                Process[] allInstances = Process.GetProcessesByName(_currentProcessName);
                results += $"Found {allInstances.Length} processes with same name\n";

                foreach (var proc in allInstances)
                {
                    try
                    {
                        results += $"  - PID: {proc.Id}, HasExited: {proc.HasExited}\n";
                    }
                    catch (Exception ex)
                    {
                        results += $"  - PID: {proc.Id}, Error: {ex.Message}\n";
                    }
                }

                return results;
            }
            catch (Exception ex)
            {
                return $"Test failed: {ex.Message}";
            }
        }

        #endregion
    }
}