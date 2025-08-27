using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Desktop_Fences
{
    /// <summary>
    /// Manages copy/paste operations for fence items, similar to BackupManager patterns
    /// Allows copying items between fences with temporary storage and cleanup
    /// </summary>
    public static class CopyPasteManager
    {
        #region Private Fields - State Management (BackupManager pattern)
        // Manage copied item state - similar to _lastDeletedFence pattern
        private static dynamic _copiedItem;
        private static bool _isCopyAvailable;
        private static string _copiedItemFolderPath;
        #endregion

        #region Send to Desktop Operations - Send shortcuts to user's desktop
        /// <summary>
        /// Sends a shortcut to the user's desktop with duplicate handling
        /// Uses BackupManager patterns for file operations and desktop detection
        /// </summary>
        /// <param name="icon">The icon/item to send to desktop</param>
        public static void SendToDesktop(dynamic icon)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, "Starting send to desktop operation");

                // Extract item properties - similar to Copy functionality
                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ?
                    dict : ((JObject)icon).ToObject<IDictionary<string, object>>();

                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                string displayName = iconDict.ContainsKey("DisplayName") ? (string)iconDict["DisplayName"] :
                    System.IO.Path.GetFileNameWithoutExtension(filePath);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"Sending to desktop: {displayName} from {filePath}");

                // Verify source shortcut exists
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate,
                        $"Source file not found or invalid: {filePath}");
                    return;
                }

                // Get desktop path - BackupManager environment detection pattern
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"Desktop path detected: {desktopPath}");

                if (string.IsNullOrEmpty(desktopPath) || !Directory.Exists(desktopPath))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                        "Could not detect or access desktop folder");
                    return;
                }

                // Generate unique filename for desktop - BackupManager duplicate handling pattern
                string originalFileName = Path.GetFileName(filePath);
                string newFileName = originalFileName;
                string desktopFilePath = Path.Combine(desktopPath, newFileName);

                // Handle duplicate filenames - BackupManager counter pattern
                int counter = 1;
                while (File.Exists(desktopFilePath))
                {
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(originalFileName);
                    string extension = Path.GetExtension(originalFileName);
                    newFileName = $"{nameWithoutExt} ({counter++}){extension}";
                    desktopFilePath = Path.Combine(desktopPath, newFileName);
                }

                // Copy shortcut to desktop - BackupManager file copy pattern
                File.Copy(filePath, desktopFilePath, true);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"Successfully copied shortcut to desktop: {newFileName}");

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                    $"Item '{displayName}' sent to desktop successfully");

                // No user notification required per user specification
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                    $"Error sending item to desktop: {ex.Message}");
                // Silent failure - no user notification required
            }
        }
        #endregion

        #region Public Properties - State Access
        /// <summary>
        /// Gets whether a copy operation is available for pasting
        /// Similar to BackupManager.IsRestoreAvailable pattern
        /// </summary>
        public static bool IsCopyAvailable => _isCopyAvailable;

        /// <summary>
        /// Gets the path to the copied item backup folder
        /// Similar to BackupManager.LastDeletedFolderPath pattern
        /// </summary>
        public static string CopiedItemFolderPath => _copiedItemFolderPath;
        #endregion

        #region Copy Operations - Adapted from BackupManager.BackupDeletedFence
        /// <summary>
        /// Copies an item to temporary storage for pasting to other fences
        /// Uses BackupManager patterns for file operations and JSON serialization
        /// </summary>
        /// <param name="icon">The icon/item to copy</param>
        /// <param name="fence">The source fence</param>
        public static void CopyItem(dynamic icon, dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, "Starting item copy operation");

                // Extract item properties - similar to BackupManager item handling
                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ?
                    dict : ((JObject)icon).ToObject<IDictionary<string, object>>();

                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                string displayName = iconDict.ContainsKey("DisplayName") ? (string)iconDict["DisplayName"] :
                    System.IO.Path.GetFileNameWithoutExtension(filePath);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"Copying item: {displayName} from {filePath}");

                // Set up copy folder path - similar to BackupManager._lastDeletedFolderPath pattern
                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                _copiedItemFolderPath = Path.Combine(exeDir, "CopiedItem");

                // Ensure the copy folder exists - BackupManager pattern
                if (!Directory.Exists(_copiedItemFolderPath))
                {
                    Directory.CreateDirectory(_copiedItemFolderPath);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        $"Created CopiedItem directory: {_copiedItemFolderPath}");
                }

                // Clear previous copied files - BackupManager cleanup pattern
                foreach (var file in Directory.GetFiles(_copiedItemFolderPath))
                {
                    File.Delete(file);
                }
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    "Cleared previous copied item files");

                // Copy the shortcut file if it exists - BackupManager file copy pattern
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    string sourceFileName = Path.GetFileName(filePath);
                    string destPath = Path.Combine(_copiedItemFolderPath, sourceFileName);

                    File.Copy(filePath, destPath, true);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        $"Copied shortcut file: {sourceFileName}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate,
                        $"Source file not found or invalid: {filePath}");
                }

                // Store item data - BackupManager JSON serialization pattern
                _copiedItem = icon;
                _isCopyAvailable = true;

                // Save item info to JSON for complete restoration - BackupManager JSON export pattern
                string itemJsonPath = Path.Combine(_copiedItemFolderPath, "CopiedItem.json");
                File.WriteAllText(itemJsonPath, JsonConvert.SerializeObject(icon, Formatting.Indented));

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                    $"Item '{displayName}' copied successfully for pasting");

                // Show auto-closing success notification - MessageDialogs pattern
                MessageBoxesManager.ShowAutoClosingMessageBoxForm($"Item '{displayName}' copied and ready to paste.", "Copy Item", 2000);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                    $"Error copying item: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error copying item: {ex.Message}", "Copy Error");

                // Reset state on error
                _isCopyAvailable = false;
                _copiedItem = null;
            }
        }
        #endregion

        #region Paste Operations - Adapted from BackupManager.RestoreLastDeletedFence
        /// <summary>
        /// Pastes the copied item to a target fence
        /// Uses BackupManager patterns for file operations and fence data updates
        /// </summary>
        /// <param name="targetFence">The fence to paste the item into</param>
        public static void PasteItem(dynamic targetFence)
        {
            try
            {
                if (!_isCopyAvailable || _copiedItem == null || string.IsNullOrEmpty(_copiedItemFolderPath))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                        "No item available to paste");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("No item to paste", "Paste Item");
                    return;
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                    $"Starting paste operation to fence: {targetFence.Title}");

                // Read copied item data - BackupManager JSON import pattern
                string itemJsonPath = Path.Combine(_copiedItemFolderPath, "CopiedItem.json");
                if (!File.Exists(itemJsonPath))
                {
                    throw new FileNotFoundException("Copied item data not found");
                }

                string jsonContent = File.ReadAllText(itemJsonPath);
                dynamic pastedItem = JsonConvert.DeserializeObject<JObject>(jsonContent);

                // Extract item properties for file operations
                string originalFilePath = pastedItem["Filename"]?.ToString();
                if (string.IsNullOrEmpty(originalFilePath))
                {
                    throw new InvalidOperationException("Invalid copied item - no filename");
                }

                string originalFileName = Path.GetFileName(originalFilePath);
                string displayName = pastedItem["DisplayName"]?.ToString() ??
                    Path.GetFileNameWithoutExtension(originalFileName);

                // Generate unique filename for the new shortcut - BackupManager duplicate handling pattern
                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string shortcutsDir = Path.Combine(exeDir, "Shortcuts");

                if (!Directory.Exists(shortcutsDir))
                {
                    Directory.CreateDirectory(shortcutsDir);
                }

                string newFileName = originalFileName;
                string newFilePath = Path.Combine(shortcutsDir, newFileName);

                // Handle duplicate filenames - BackupManager counter pattern
                int counter = 1;
                while (File.Exists(newFilePath))
                {
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(originalFileName);
                    string extension = Path.GetExtension(originalFileName);
                    newFileName = $"{nameWithoutExt} ({counter++}){extension}";
                    newFilePath = Path.Combine(shortcutsDir, newFileName);
                }

                // Copy shortcut from temp folder to Shortcuts folder - BackupManager file copy pattern
                string sourceFilePath = Path.Combine(_copiedItemFolderPath, originalFileName);
                if (File.Exists(sourceFilePath))
                {
                    File.Copy(sourceFilePath, newFilePath, true);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        $"Copied shortcut file: {newFileName}");
                }
                else
                {
                    throw new FileNotFoundException($"Source shortcut file not found: {sourceFilePath}");
                }

                // Update item data for the new location - BackupManager data update pattern
                pastedItem["Filename"] = Path.Combine("Shortcuts", newFileName);

                // Keep the original display name clean - don't add counter to displayed title
                // The filename will have (1), (2) etc. for uniqueness, but the displayed name stays clean
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"Pasted item - Filename: '{newFileName}', DisplayName: '{displayName}' (kept clean)");

                // Note: We intentionally do NOT update DisplayName with counter
                // This keeps the displayed name clean while the filename ensures uniqueness

                // Add item to target fence - BackupManager fence data update pattern
                var targetItems = targetFence.Items as JArray;
                if (targetItems == null)
                {
                    targetItems = new JArray();
                    targetFence.Items = targetItems;
                }

                targetItems.Add(pastedItem);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"Added item to fence '{targetFence.Title}'");

                // Save fence data - BackupManager save pattern
                FenceDataManager.SaveFenceData();

                // Refresh fence display - Find the fence window and refresh
                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                var targetWindow = windows.FirstOrDefault(w => w.Tag?.ToString() == targetFence.Id?.ToString());

                if (targetWindow != null)
                {
                    // Refresh the fence content to show the new item using the public method
                    FenceManager.RefreshFenceUsingFormApproach(targetWindow, targetFence);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        "Refreshed fence display after paste");
                }

                // Clean up copied item - BackupManager cleanup pattern
                CleanCopiedItem();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                    $"Item pasted successfully to fence '{targetFence.Title}'");

                // No success message for paste - keep it silent for better UX
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                    $"Error pasting item: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error pasting item: {ex.Message}", "Paste Error");
            }
        }
        #endregion

        #region Cleanup Operations - Adapted from BackupManager.CleanLastDeletedFolder
        /// <summary>
        /// Cleans the copied item folder and resets availability state
        /// Uses BackupManager cleanup patterns
        /// </summary>
        public static void CleanCopiedItem()
        {
            try
            {
                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                _copiedItemFolderPath = Path.Combine(exeDir, "CopiedItem");

                if (Directory.Exists(_copiedItemFolderPath))
                {
                    // Clear all files in the copy folder - BackupManager cleanup pattern
                    foreach (var file in Directory.GetFiles(_copiedItemFolderPath))
                    {
                        File.Delete(file);
                    }

                    // Remove the directory itself
                    Directory.Delete(_copiedItemFolderPath);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        "Cleaned copied item folder");
                }

                // Reset state - BackupManager reset pattern
                _isCopyAvailable = false;
                _copiedItem = null;
                _copiedItemFolderPath = null;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    "Reset copy/paste state");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                    $"Error cleaning copied item: {ex.Message}");
            }
        }

        /// <summary>
        /// Checks if there's a copied item available for pasting
        /// Uses BackupManager availability checking patterns
        /// </summary>
        public static bool HasCopiedItem()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"Checking copied item availability - _isCopyAvailable: {_isCopyAvailable}, _copiedItem: {(_copiedItem != null ? "exists" : "null")}");

                if (!_isCopyAvailable || _copiedItem == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        "No copied item available - flags not set");
                    return false;
                }

                // Verify files still exist - BackupManager validation pattern
                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string copiedItemPath = Path.Combine(exeDir, "CopiedItem");
                string jsonPath = Path.Combine(copiedItemPath, "CopiedItem.json");

                bool folderExists = Directory.Exists(copiedItemPath);
                bool jsonExists = File.Exists(jsonPath);
                bool hasValidCopy = folderExists && jsonExists;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"File validation - Folder exists: {folderExists}, JSON exists: {jsonExists}, Path: {copiedItemPath}");

                if (!hasValidCopy)
                {
                    // Reset state if files are missing - BackupManager consistency pattern
                    _isCopyAvailable = false;
                    _copiedItem = null;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        "Reset copy state - files no longer exist");
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                    $"HasCopiedItem result: {hasValidCopy}");
                return hasValidCopy;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                    $"Error checking copied item availability: {ex.Message}");
                return false;
            }
        }
        #endregion
    }
}