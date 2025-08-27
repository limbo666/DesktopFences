using IWshRuntimeLibrary;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Desktop_Fences
{
    /// <summary>
    /// Centralized file path and shortcut utilities with Unicode support
    /// Extracted from FenceManager for better code organization and reusability
    /// Handles complex Unicode path scenarios and shortcut resolution
    /// </summary>
    public static class FilePathUtilities
    {

        #region Unicode Shortcut Resolution - Used by: FenceManager (UpdateIcon, ClickEventAdder, LaunchItem)
        /// <summary>
        /// Enhanced shortcut target resolution with Unicode support
        /// Handles both direct shortcuts and explorer.exe-based folder shortcuts
        /// Used by: FenceManager.UpdateIcon, FenceManager.ClickEventAdder, FenceManager.LaunchItem
        /// Category: Unicode Path Resolution
        /// Moved from: FenceManager.GetShortcutTargetUnicodeSafe
        /// </summary>
        public static string GetShortcutTargetUnicodeSafe(string shortcutPath)
        {
            try
            {
                if (SettingsManager.EnableBackgroundValidationLogging)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Attempting Unicode-safe shortcut resolution for: {shortcutPath}");
                }
                // Verify the shortcut file exists
                if (!System.IO.File.Exists(shortcutPath))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                        $"Shortcut file not found: {shortcutPath}");
                    return string.Empty;
                }

                // Method 1: Try WshShell with enhanced Unicode folder detection
                try
                {
                    WshShell shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                    string targetPath = shortcut.TargetPath?.Trim();
                    string arguments = shortcut.Arguments?.Trim();

                    // Check if this is our Unicode folder shortcut (explorer.exe + folder argument)
                    if (!string.IsNullOrEmpty(targetPath) &&
                        targetPath.ToLower().EndsWith("explorer.exe") &&
                        !string.IsNullOrEmpty(arguments))
                    {
                        // Extract folder path from arguments (remove quotes)
                        string folderPath = arguments.Trim('"', ' ', '\t');
                        if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                                $"WshShell resolved Unicode folder shortcut: {shortcutPath} -> {folderPath}");
                            return folderPath;
                        }
                    }

                    // Regular shortcut - return target if valid
                    if (!string.IsNullOrEmpty(targetPath) && (System.IO.File.Exists(targetPath) || Directory.Exists(targetPath)))
                    {
                        if (SettingsManager.EnableBackgroundValidationLogging)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.BackgroundValidation,
                                $"WshShell successfully resolved regular shortcut: {shortcutPath} -> {targetPath}");
                        }
                        return targetPath;
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                        $"WshShell method failed for {shortcutPath}: {ex.Message}");
                }

                // Method 2: Binary parsing for Unicode shortcuts (advanced fallback)
                try
                {
                    string targetPath = ParseShortcutBinary(shortcutPath);
                    if (!string.IsNullOrEmpty(targetPath))
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                            $"Binary parsing resolved Unicode shortcut: {shortcutPath} -> {targetPath}");
                        return targetPath;
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                        $"Binary parsing failed for Unicode shortcut {shortcutPath}: {ex.Message}");
                }

                // Method 3: Final fallback - try original Utility method
                try
                {
                    string fallbackPath = Utility.GetShortcutTarget(shortcutPath);
                    if (!string.IsNullOrEmpty(fallbackPath))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Original method worked for Unicode shortcut: {shortcutPath} -> {fallbackPath}");
                        return fallbackPath;
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                        $"All methods failed for Unicode shortcut {shortcutPath}: {ex.Message}");
                }

                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Unicode shortcut resolution failed completely for: {shortcutPath}");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Critical error in Unicode shortcut resolution for {shortcutPath}: {ex.Message}");
                return string.Empty;
            }
        }
        #endregion

        /// <summary>
        /// Checks if the actual folder exists for folder shortcuts (including Unicode)
        /// Used by: Icon extraction logic to determine if folder-WhiteX.png should be used
        /// Category: Folder Validation
        /// </summary>
        public static bool DoesFolderExist(string shortcutPath, bool isFolder)
        {
            try
            {
                if (!isFolder)
                {
                    // Not a folder, return false
                    return false;
                }

                // Check if it's a shortcut first
                if (Path.GetExtension(shortcutPath).ToLower() != ".lnk")
                {
                    // Direct folder path
                    return Directory.Exists(shortcutPath);
                }

                // It's a shortcut - get the real folder path
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                string targetPath = shortcut.TargetPath?.Trim();
                string arguments = shortcut.Arguments?.Trim();

                // Check if this is a Unicode folder shortcut (explorer.exe + folder argument)
                if (!string.IsNullOrEmpty(targetPath) &&
                    targetPath.ToLower().EndsWith("explorer.exe") &&
                    !string.IsNullOrEmpty(arguments))
                {
                    // Extract the actual folder path from arguments
                    string folderPath = arguments.Trim('"', ' ', '\t');
                    if (!string.IsNullOrEmpty(folderPath))
                    {
                        bool exists = Directory.Exists(folderPath);
                        if (SettingsManager.EnableBackgroundValidationLogging)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.BackgroundValidation,
                                $"Unicode folder existence check: {folderPath} -> {exists}");
                        }
                        return exists;
                    }
                }
                else if (!string.IsNullOrEmpty(targetPath))
                {
                    // Regular folder shortcut
                    bool exists = Directory.Exists(targetPath);
                    if (SettingsManager.EnableBackgroundValidationLogging)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.BackgroundValidation,
                            $"Regular folder existence check: {targetPath} -> {exists}");
                    }
                    return exists;
                }

                return false;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error checking folder existence for {shortcutPath}: {ex.Message}");
                return false;
            }
        }

        #region Unicode Folder Validation - Used by: FenceManager, Portal Fence operations
        /// <summary>
        /// Unicode-safe folder validation and processing
        /// Handles folders with Unicode characters in their names
        /// Used by: FenceManager.CreateFence, Portal fence operations
        /// Category: Folder Validation
        /// Moved from: FenceManager.ValidateUnicodeFolderPath
        /// </summary>
        public static bool ValidateUnicodeFolderPath(string folderPath, out string sanitizedPath)
        {
            sanitizedPath = folderPath;

            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                    $"Validating Unicode folder path: {folderPath}");

                // Method 1: Direct validation
                if (Directory.Exists(folderPath))
                {
                    // Get the full path to normalize it
                    sanitizedPath = Path.GetFullPath(folderPath);
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                        $"Unicode folder validation successful: {folderPath} -> {sanitizedPath}");
                    return true;
                }

                // Method 2: Try with different encoding approaches
                try
                {
                    var dirInfo = new DirectoryInfo(folderPath);
                    if (dirInfo.Exists)
                    {
                        sanitizedPath = dirInfo.FullName;
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                            $"Unicode folder validation via DirectoryInfo: {folderPath} -> {sanitizedPath}");
                        return true;
                    }
                }
                catch (Exception dirEx)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                        $"DirectoryInfo validation failed for {folderPath}: {dirEx.Message}");
                }

                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                    $"Unicode folder validation failed: {folderPath}");
                return false;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Critical error validating Unicode folder {folderPath}: {ex.Message}");
                sanitizedPath = folderPath; // Keep original on error
                return false;
            }
        }
        #endregion

        #region Path Validation Helpers - Used by: File operations throughout application

        #endregion

        #region Dead Shortcut Management - Used by: FenceManager context menu
        /// <summary>
        /// Clears all dead shortcuts from a data fence (excludes Portal fences and web links)
        /// Used by: FenceManager context menu "Clear Dead Shortcuts"
        /// Category: Dead Shortcut Management
        /// </summary>
        /// <summary>
        /// Clears all dead shortcuts from a data fence (excludes Portal fences and web links)
        /// Used by: FenceManager context menu "Clear Dead Shortcuts"
        /// Category: Dead Shortcut Management
        /// </summary>
        public static int ClearDeadShortcutsFromFence(dynamic fence)
        {
            try
            {
                // Exclude Portal fences - they work differently
                if (fence.ItemsType?.ToString() != "Data")
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        $"Skipping dead shortcut cleanup for non-Data fence '{fence.Title}' (Type: {fence.ItemsType})");
                    return 0;
                }

                int removedCount = 0;
                bool fenceModified = false;

                // Handle tabbed fences
                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

                if (tabsEnabled)
                {
                    var tabs = fence.Tabs as JArray ?? new JArray();

                    foreach (var tab in tabs.Cast<JObject>())
                    {
                        var items = tab["Items"] as JArray ?? new JArray();
                        removedCount += RemoveDeadItemsFromArray(items);
                        if (removedCount > 0) fenceModified = true;
                    }
                }
                else
                {
                    // Handle regular fence (non-tabbed)
                    var items = fence.Items as JArray ?? new JArray();
                    removedCount = RemoveDeadItemsFromArray(items);
                    if (removedCount > 0) fenceModified = true;
                }

                // Save changes if any items were removed
                if (fenceModified)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                        $"Removed {removedCount} dead shortcuts from fence '{fence.Title}'");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        $"No dead shortcuts found in fence '{fence.Title}'");
                }

                return removedCount;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                    $"Error clearing dead shortcuts from fence '{fence.Title}': {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Helper function to remove dead items from a JArray of items
        /// Used by: ClearDeadShortcutsFromFence
        /// Category: Dead Shortcut Management
        /// </summary>
        private static int RemoveDeadItemsFromArray(JArray items)
        {
            int removedCount = 0;

            // Work backwards to safely remove items during iteration
            for (int i = items.Count - 1; i >= 0; i--)
            {
                var item = items[i] as JObject;
                if (item == null) continue;

                // Get item properties
                string filename = item["Filename"]?.ToString();
                bool isLink = item["IsLink"]?.ToString().ToLower() == "true";
                bool isFolder = item["IsFolder"]?.ToString().ToLower() == "true";

                // Skip web links - they don't have local file targets
                if (isLink)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        $"Skipping web link item: {filename}");
                    continue;
                }

                // Skip if filename is invalid
                if (string.IsNullOrEmpty(filename))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate,
                        "Found item with empty filename, removing");
                    items.RemoveAt(i);
                    removedCount++;
                    continue;
                }

                // Check if the target exists
                bool targetExists = false;

                try
                {
                    bool isShortcut = Path.GetExtension(filename).ToLower() == ".lnk";

                    if (isShortcut)
                    {
                        // For shortcuts, check the target
                        if (isFolder)
                        {
                            // Use existing folder checking logic
                            targetExists = DoesFolderExist(filename, true);
                        }
                        else
                        {
                            // For file shortcuts, resolve target and check existence
                            string targetPath = GetShortcutTargetUnicodeSafe(filename);
                            targetExists = !string.IsNullOrEmpty(targetPath) &&
                                         (System.IO.File.Exists(targetPath) || Directory.Exists(targetPath));
                        }
                    }
                    else
                    {
                        // Direct file/folder reference
                        targetExists = isFolder ? Directory.Exists(filename) : System.IO.File.Exists(filename);
                    }

                    // Remove if target doesn't exist
                    if (!targetExists)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                            $"Removing dead {(isFolder ? "folder" : "file")} shortcut: {filename}");

                        // Delete the actual shortcut file from Shortcuts folder (same as manual Remove)
                        DeleteShortcutFile(filename);

                        items.RemoveAt(i);
                        removedCount++;
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                        $"Error checking item '{filename}': {ex.Message}. Removing item.");
                    items.RemoveAt(i);
                    removedCount++;
                }
            }

            return removedCount;
        }

        /// <summary>
        /// Deletes the actual shortcut file from Shortcuts folder and backup location
        /// Used by: ClearDeadShortcutsFromFence
        /// Category: File Management
        /// </summary>
        private static void DeleteShortcutFile(string filename)
        {
            try
            {
                // Delete main shortcut file (same logic as manual Remove)
                string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
                string shortcutPath = System.IO.Path.Combine(exeDir, "Shortcuts", System.IO.Path.GetFileName(filename));

                if (System.IO.File.Exists(shortcutPath))
                {
                    try
                    {
                        System.IO.File.Delete(shortcutPath);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                            $"Deleted shortcut file: {shortcutPath}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                            $"Failed to delete shortcut {shortcutPath}: {ex.Message}");
                    }
                }

                // Delete backup shortcut if it exists (same logic as manual Remove)
                string tempShortcutsDir = System.IO.Path.Combine(exeDir, "Temp Shortcuts");
                string backupPath = System.IO.Path.Combine(tempShortcutsDir, System.IO.Path.GetFileName(filename));

                if (System.IO.File.Exists(backupPath))
                {
                    try
                    {
                        System.IO.File.Delete(backupPath);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                            $"Deleted backup shortcut: {backupPath}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                            $"Failed to delete backup shortcut {backupPath}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                    $"Error deleting shortcut files for {filename}: {ex.Message}");
            }
        }
        #endregion

        #region Binary Shortcut Parsing - Internal Helper
        /// <summary>
        /// Advanced binary parsing for Unicode shortcuts
        /// Used by: GetShortcutTargetUnicodeSafe as fallback method
        /// Category: Binary File Parsing
        /// </summary>
        private static string ParseShortcutBinary(string shortcutPath)
        {
            try
            {
                byte[] fileContent = System.IO.File.ReadAllBytes(shortcutPath);

                // Convert to both Unicode and ANSI for comprehensive searching
                string fileContentUnicode = System.Text.Encoding.Unicode.GetString(fileContent);
                string fileContentAnsi = System.Text.Encoding.Default.GetString(fileContent);
                string content = fileContentUnicode;

                // Look for explorer.exe with folder arguments (Unicode folder shortcuts)
                var explorerMatches = Regex.Matches(content,
                    @"explorer\.exe\s*""([^""]+)""",
                    RegexOptions.IgnoreCase);

                foreach (Match match in explorerMatches)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                        $"Found explorer.exe match: '{match.Value}', Groups: {match.Groups.Count}");

                    if (match.Groups.Count > 1)
                    {
                        string candidatePath = match.Groups[1].Value.Trim('\0', ' ', '\t', '\r', '\n');
                        if (!string.IsNullOrEmpty(candidatePath) && Directory.Exists(candidatePath))
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                                $"Found Unicode folder path via explorer.exe method: {candidatePath}");
                            return candidatePath;
                        }
                    }
                }

                // Try ANSI encoding for explorer.exe shortcuts
                explorerMatches = Regex.Matches(fileContentAnsi,
                    @"explorer\.exe\s*""([^""]+)""",
                    RegexOptions.IgnoreCase);

                foreach (Match match in explorerMatches)
                {
                    if (match.Groups.Count > 1)
                    {
                        string candidatePath = match.Groups[1].Value.Trim('\0', ' ', '\t', '\r', '\n');
                        if (!string.IsNullOrEmpty(candidatePath) && Directory.Exists(candidatePath))
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                                $"Found folder path via explorer.exe method (ANSI): {candidatePath}");
                            return candidatePath;
                        }
                    }
                }

                // Look for regular folder paths
                var folderMatches = Regex.Matches(content,
                    @"[A-Za-z]:[\\\/][^<>:""|?*\x00-\x1f\x20]+",
                    RegexOptions.IgnoreCase);

                foreach (Match match in folderMatches)
                {
                    string candidatePath = match.Value.Trim('\0', ' ', '\t', '\r', '\n');
                    candidatePath = candidatePath.Replace("\0", "").Trim();

                    if (candidatePath.Length > 5 && candidatePath.Contains("\\"))
                    {
                        try
                        {
                            if (Directory.Exists(candidatePath))
                            {
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                    $"Found valid folder path in binary: {candidatePath}");
                                return candidatePath;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                // Look for executable paths as final fallback
                var pathMatches = Regex.Matches(fileContentUnicode,
                    @"[A-Za-z]:\\[^<>:""|?*\x00-\x1f]*\.exe",
                    RegexOptions.IgnoreCase);

                foreach (Match match in pathMatches)
                {
                    string candidatePath = match.Value.Trim('\0', ' ', '\t', '\r', '\n');
                    if (System.IO.File.Exists(candidatePath))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Found valid executable path in LNK binary: {candidatePath}");
                        return candidatePath;
                    }
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error parsing shortcut binary {shortcutPath}: {ex.Message}");
                return string.Empty;
            }
        }
        #endregion

        #region URL File Processing - Used by: FenceUtilities (can be enhanced)
        /// <summary>
        /// Enhanced URL extraction from .url files with error handling
        /// Used by: FenceUtilities.ExtractUrlFromFile (can replace existing implementation)
        /// Category: URL Processing
        /// </summary>
        public static string ExtractUrlFromFileAdvanced(string urlFilePath)
        {
            try
            {
                if (!CoreUtilities.SafeFileExists(urlFilePath))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                        $"URL file does not exist: {urlFilePath}");
                    return null;
                }

                string[] lines = System.IO.File.ReadAllLines(urlFilePath);
                foreach (string line in lines)
                {
                    if (line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                    {
                        string url = line.Substring(4).Trim();

                        // Basic URL validation
                        if (Uri.TryCreate(url, UriKind.Absolute, out Uri validUri) &&
                            (validUri.Scheme == Uri.UriSchemeHttp || validUri.Scheme == Uri.UriSchemeHttps))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                $"Extracted and validated URL from {urlFilePath}: {url}");
                            return url;
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                                $"Invalid URL format extracted from {urlFilePath}: {url}");
                            return url; // Return anyway, let caller decide
                        }
                    }
                }

                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                    $"No URL= line found in .url file: {urlFilePath}");
                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"Error reading .url file {urlFilePath}: {ex.Message}");
                return null;
            }
        }
        #endregion
    }
}