using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Desktop_Fences
{
    /// <summary>
    /// Registry Helper Class for Single Instance Management
    /// Handles all registry operations for the instance trigger system
    /// Compatible with existing project structure and follows established patterns
    /// </summary>
    public static class RegistryHelper
    {
        #region Constants

        // Registry path for our trigger system
        private static readonly string REGISTRY_KEY_PATH = @"SOFTWARE\Desktop_Fences_Plus\InstanceTrigger";
        private static readonly string TRIGGER_VALUE_NAME = "TriggerEffect";

        // Registry path for program management values
        private static readonly string PROGRAM_REGISTRY_KEY_PATH = @"SOFTWARE\Desktop_Fences_Plus\ProgramManagement";

        #endregion

        #region Public Methods

        /// <summary>
        /// Writes a trigger value to registry to signal effect activation
        /// Used by: Single instance checker when duplicate instance found
        /// </summary>
        /// <returns>True if write successful, false otherwise</returns>
        public static bool WriteTrigger()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(REGISTRY_KEY_PATH))
                {
                    if (key != null)
                    {
                        // Use current timestamp to ensure uniqueness and prevent duplicate triggers
                        string triggerValue = DateTime.Now.Ticks.ToString();
                        key.SetValue(TRIGGER_VALUE_NAME, triggerValue, RegistryValueKind.String);

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                            $"RegistryHelper: Wrote trigger value: {triggerValue}");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryHelper: Error writing trigger: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Checks for trigger value in registry
        /// Used by: Registry monitor timer to detect when effect should be activated
        /// </summary>
        /// <returns>Trigger value if found, null if not found or error</returns>
        public static string CheckForTrigger()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(REGISTRY_KEY_PATH))
                {
                    if (key != null)
                    {
                        object value = key.GetValue(TRIGGER_VALUE_NAME);
                        if (value != null)
                        {
                            string triggerValue = value.ToString();
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                $"RegistryHelper: Found trigger value: {triggerValue}");
                            return triggerValue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryHelper: Error checking trigger: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Deletes the trigger value from registry
        /// Used by: Registry monitor after effect is triggered to prevent re-triggering
        /// </summary>
        /// <returns>True if delete successful or value didn't exist, false on error</returns>
        public static bool DeleteTrigger()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(REGISTRY_KEY_PATH, writable: true))
                {
                    if (key != null)
                    {
                        // Check if value exists before trying to delete
                        object value = key.GetValue(TRIGGER_VALUE_NAME);
                        if (value != null)
                        {
                            key.DeleteValue(TRIGGER_VALUE_NAME, throwOnMissingValue: false);
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                $"RegistryHelper: Deleted trigger value: {value}");
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryHelper: Error deleting trigger: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Cleans up registry key (removes entire key path)
        /// Used by: Application shutdown or cleanup operations (optional)
        /// </summary>
        /// <returns>True if cleanup successful, false otherwise</returns>
        public static bool CleanupRegistry()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(@"SOFTWARE\Desktop_Fences_Plus\InstanceTrigger", throwOnMissingSubKey: false);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    "RegistryHelper: Cleaned up registry key");
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryHelper: Error cleaning registry: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Test method to verify registry operations work correctly
        /// Used by: Development/testing purposes only
        /// </summary>
        /// <returns>True if all operations work correctly</returns>
        public static bool TestRegistryOperations()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    "RegistryHelper: Starting registry operations test");

                // Test write
                if (!WriteTrigger())
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                        "RegistryHelper: Test failed - WriteTrigger returned false");
                    return false;
                }

                // Test read
                string triggerValue = CheckForTrigger();
                if (string.IsNullOrEmpty(triggerValue))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                        "RegistryHelper: Test failed - CheckForTrigger returned null/empty");
                    return false;
                }

                // Test delete
                if (!DeleteTrigger())
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                        "RegistryHelper: Test failed - DeleteTrigger returned false");
                    return false;
                }

                // Verify deletion
                string afterDelete = CheckForTrigger();
                if (!string.IsNullOrEmpty(afterDelete))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                        "RegistryHelper: Test failed - Trigger still exists after deletion");
                    return false;
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    "RegistryHelper: All registry operations test passed successfully");
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryHelper: Test failed with exception: {ex.Message}");
                return false;
            }
        }

        #endregion
        #region Program Management Methods

        /// <summary>
        /// Sets or updates program management values in registry
        /// Updates values based on specified rules (some update on each run, others only if not exist)
        /// </summary>
        /// <param name="programVersion">Current program version</param>
        /// <returns>True if operations successful, false otherwise</returns>
        public static bool SetProgramManagementValues(string programVersion)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(PROGRAM_REGISTRY_KEY_PATH))
                {
                    if (key != null)
                    {
                        // Get current executable path
                        string currentProgramPath = Assembly.GetEntryAssembly()?.Location ?? "";

                        // ProgramPath: Updates on each run
                        key.SetValue("ProgramPath", currentProgramPath, RegistryValueKind.String);

                        // FirstRunDate: Updates only if value doesn't exist
                        if (key.GetValue("FirstRunDate") == null)
                        {
                            key.SetValue("FirstRunDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), RegistryValueKind.String);
                        }

                        // FirstVersion: Updates only if value doesn't exist
                        if (key.GetValue("FirstVersion") == null)
                        {
                            key.SetValue("FirstVersion", programVersion, RegistryValueKind.String);
                        }

                        // CurrentVersion: Updates on each run if value is greater than existing
                        var existingVersionObj = key.GetValue("CurrentVersion");
                        if (existingVersionObj == null || IsVersionGreater(programVersion, existingVersionObj.ToString()))
                        {
                            key.SetValue("CurrentVersion", programVersion, RegistryValueKind.String);
                        }

                        // DVC: Updates only if value doesn't exist
                        if (key.GetValue("DVC") == null)
                        {
                            key.SetValue("DVC", "893579621b01f56b6f508bdc0e6c34f84a2c9ec2dbd0aa72b02d94a3708d3e9c", RegistryValueKind.String);
                        }

                        // UnitID: Updates only if value doesn't exist
                        if (key.GetValue("UnitID") == null)
                        {
                            key.SetValue("UnitID", "df001", RegistryValueKind.String);
                        }

                        // EnPU: Updates only if value doesn't exist
                        if (key.GetValue("EnPU") == null)
                        {
                            key.SetValue("EnPU", 0, RegistryValueKind.DWord);
                        }

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                            "RegistryHelper: Program management values set successfully");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryHelper: Error setting program management values: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Retrieves all program management values from registry
        /// </summary>
        /// <returns>Dictionary containing all registry values, empty if error</returns>
        public static Dictionary<string, object> GetProgramManagementValues()
        {
            var values = new Dictionary<string, object>();

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(PROGRAM_REGISTRY_KEY_PATH))
                {
                    if (key != null)
                    {
                        values["ProgramPath"] = key.GetValue("ProgramPath") ?? "";
                        values["FirstRunDate"] = key.GetValue("FirstRunDate") ?? "";
                        values["FirstVersion"] = key.GetValue("FirstVersion") ?? "";
                        values["CurrentVersion"] = key.GetValue("CurrentVersion") ?? "";
                        values["DVC"] = key.GetValue("DVC") ?? "";
                        values["UnitID"] = key.GetValue("UnitID") ?? "";
                        values["EnPU"] = key.GetValue("EnPU") ?? 0;

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                            "RegistryHelper: Program management values retrieved successfully");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryHelper: Error retrieving program management values: {ex.Message}");
            }

            return values;
        }

        /// <summary>
        /// Exports all program management values to a text file in program directory
        /// </summary>
        /// <returns>True if export successful, false otherwise</returns>
        public static bool ExportProgramManagementValues()
        {
            try
            {
                var values = GetProgramManagementValues();
                string programPath = Assembly.GetEntryAssembly()?.Location ?? "";
                string programDir = System.IO.Path.GetDirectoryName(programPath) ?? "";
                string exportFilePath = System.IO.Path.Combine(programDir, "Desktop Fences + Registry Values.txt");

                using (var writer = new System.IO.StreamWriter(exportFilePath))
                {
                    writer.WriteLine("Desktop Fences Plus - Registry Values Export");
                    writer.WriteLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    writer.WriteLine(new string('-', 50));
                    writer.WriteLine();

                    foreach (var kvp in values)
                    {
                        writer.WriteLine($"{kvp.Key}: {kvp.Value}");
                    }

                    writer.WriteLine();
                    writer.WriteLine(new string('-', 50));
                    writer.WriteLine("End of Registry Values");
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    $"RegistryHelper: Registry values exported to: {exportFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryHelper: Error exporting registry values: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Helper method to compare version strings
        /// </summary>
        /// <param name="newVersion">New version to compare</param>
        /// <param name="existingVersion">Existing version to compare against</param>
        /// <returns>True if newVersion is greater than existingVersion</returns>
        private static bool IsVersionGreater(string newVersion, string existingVersion)
        {
            try
            {
                // Simple version comparison - you can enhance this based on your version format
                if (string.IsNullOrEmpty(existingVersion)) return true;
                if (string.IsNullOrEmpty(newVersion)) return false;

                // Try to parse as Version objects for proper comparison
                if (Version.TryParse(newVersion, out Version newVer) &&
                    Version.TryParse(existingVersion, out Version existingVer))
                {
                    return newVer > existingVer;
                }

                // Fallback to string comparison if not valid version format
                return string.Compare(newVersion, existingVersion, StringComparison.OrdinalIgnoreCase) > 0;
            }
            catch
            {
                return false;
            }
        }

        #endregion
      
    }
}