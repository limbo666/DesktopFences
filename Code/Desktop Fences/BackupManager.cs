using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;

namespace Desktop_Fences
{
    /// <summary>
    /// Manages backup operations for Desktop Fences data, including the fences configuration file
    /// and associated shortcut files. Backups are stored in timestamped folders for easy recovery.
    /// Refactored to centralize all backup/restore functionality and support Multi-Profile.
    /// </summary>
    public static class BackupManager
    {
        #region Private Fields
        // Track active legendary effects to allow cleanup/reversion
        private static readonly Dictionary<string, Storyboard> _legendaryEffects = new Dictionary<string, Storyboard>();
        private static readonly Dictionary<string, Brush> _originalBorders = new Dictionary<string, Brush>();
        private static readonly Dictionary<string, Thickness> _originalBorderThicknesses = new Dictionary<string, Thickness>();
        private static readonly Dictionary<string, Effect> _originalEffects = new Dictionary<string, Effect>();

        // Manage last deleted fence restoration
        private static string _lastDeletedFolderPath;
        private static dynamic _lastDeletedFence;
        private static bool _isRestoreAvailable;
        private static System.Windows.Threading.DispatcherTimer _autoBackupTimer;
        #endregion

        #region Public Properties
        // Gets whether a restore operation is available for the last deleted fence
        public static bool IsRestoreAvailable => _isRestoreAvailable;

        // Gets the path to the last deleted fence backup folder
        public static string LastDeletedFolderPath => _lastDeletedFolderPath;
        #endregion

        #region Public UI Helpers (NEW)

        /// <summary>
        /// Opens the "Backups" folder for the CURRENT ACTIVE PROFILE in File Explorer.
        /// </summary>
        public static void OpenBackupsFolder()
        {
            try
            {
                string path = ProfileManager.GetProfileFilePath("Backups");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                Process.Start("explorer.exe", path);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Failed to open backups folder: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns the full path to the "Backups" folder for the CURRENT ACTIVE PROFILE.
        /// Useful for initializing OpenFileDialogs.
        /// </summary>
        public static string GetBackupsFolderPath()
        {
            string path = ProfileManager.GetProfileFilePath("Backups");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }

        #endregion

        #region Existing Backup Method (Enhanced)

        // Replaces existing BackupData to use the new shared helper
        public static void BackupData()
        {
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Starting manual backup");
            // Standard format: [datetime]_backup
            string backupName = DateTime.Now.ToString("yyMMddHHmm") + "_backup";
            CreateBackup(backupName, silent: false);
        }

        #endregion

        #region Restore Operations

        public static void RestoreLastDeletedFence()
        {
            try
            {
                if (!_isRestoreAvailable || string.IsNullOrEmpty(_lastDeletedFolderPath) || _lastDeletedFence == null)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "No fence available to restore");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("No fence to restore", "Restore");
                    return;
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Restoring last deleted fence: {_lastDeletedFence.Title}");

                // Restore shortcuts if they exist
                if (Directory.Exists(_lastDeletedFolderPath))
                {
                    var shortcutFiles = Directory.GetFiles(_lastDeletedFolderPath, "*.lnk");

                    // FIX: Use Profile Path for Shortcuts
                    string shortcutsDir = ProfileManager.GetProfileFilePath("Shortcuts");

                    // Ensure shortcuts directory exists
                    if (!Directory.Exists(shortcutsDir))
                    {
                        Directory.CreateDirectory(shortcutsDir);
                    }

                    foreach (var shortcutFile in shortcutFiles)
                    {
                        string destinationPath = Path.Combine(shortcutsDir, Path.GetFileName(shortcutFile));
                        File.Copy(shortcutFile, destinationPath, true);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Restored shortcut: {Path.GetFileName(shortcutFile)}");
                    }
                }

                // Restore fence data to the main fence collection
                var fenceData = FenceManager.GetFenceData();
                fenceData.Add(_lastDeletedFence);
                FenceDataManager.SaveFenceData();

                // Create the fence UI
                FenceManager.CreateFence(_lastDeletedFence, new TargetChecker(1000));

                // Clear backup state
                _lastDeletedFence = null;
                _isRestoreAvailable = false;

                // Update heart context menus to reflect restored state
                FenceManager.UpdateAllHeartContextMenus();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Last deleted fence restored successfully");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Error restoring last deleted fence: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error restoring fence: {ex.Message}", "Restore Error");
            }
        }
        #endregion

        #region Fence Export/Import Operations

        public static void ExportFence(dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Starting export of fence: {fence.Title}");

                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string fenceTitle = fence.Title.ToString();

                // Sanitize folder name
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    fenceTitle = fenceTitle.Replace(c, '_');
                }

                // Exports go to GLOBAL "Exports" folder (shared between profiles)
                string exportFolder = Path.Combine(exeDir, "Exports", fenceTitle);
                string fencePath = Path.Combine(exeDir, "Exports", $"{fenceTitle}.fence");

                // Ensure exports directory exists
                string exportsDir = Path.Combine(exeDir, "Exports");
                if (!Directory.Exists(exportsDir)) Directory.CreateDirectory(exportsDir);

                // Cleanup previous runs
                if (Directory.Exists(exportFolder)) Directory.Delete(exportFolder, true);
                if (File.Exists(fencePath)) File.Delete(fencePath);

                Directory.CreateDirectory(exportFolder);

                // 1. Save Metadata
                string fenceJson = JsonConvert.SerializeObject(fence, Formatting.Indented);
                File.WriteAllText(Path.Combine(exportFolder, "fence.json"), fenceJson);

                // 2. Copy Shortcuts (TAB-AWARE FIX)
                if (fence.ItemsType?.ToString() == "Data")
                {
                    string shortcutsDestDir = Path.Combine(exportFolder, "Shortcuts");
                    Directory.CreateDirectory(shortcutsDestDir);
                    int copiedShortcuts = 0;

                    // Helper to copy a list of items
                    void CopyItems(JArray items)
                    {
                        if (items == null) return;
                        foreach (var item in items)
                        {
                            string filename = item["Filename"]?.ToString();
                            if (!string.IsNullOrEmpty(filename))
                            {
                                // FIX: Resolve source path from Profile
                                string sourcePath = Path.Combine(ProfileManager.CurrentProfileDir, filename);

                                if (File.Exists(sourcePath))
                                {
                                    string destName = Path.GetFileName(filename);
                                    string destPath = Path.Combine(shortcutsDestDir, destName);

                                    // Avoid duplicate copy crashes
                                    if (!File.Exists(destPath))
                                    {
                                        File.Copy(sourcePath, destPath);
                                        copiedShortcuts++;
                                    }
                                }
                            }
                        }
                    }

                    // A. Copy Main Items
                    var mainItems = fence.Items as JArray;
                    if (mainItems != null) CopyItems(mainItems);

                    // B. Copy Tab Items (The Fix)
                    bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                    if (tabsEnabled)
                    {
                        var tabs = fence.Tabs as JArray;
                        if (tabs != null)
                        {
                            foreach (var tab in tabs)
                            {
                                var tabItems = tab["Items"] as JArray;
                                if (tabItems != null) CopyItems(tabItems);
                            }
                        }
                    }

                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Copied {copiedShortcuts} total shortcuts to export");
                }

                // 3. Zip It
                ZipFile.CreateFromDirectory(exportFolder, fencePath);
                Directory.Delete(exportFolder, true);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Fence exported successfully: {fencePath}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Fence exported to:\n{fencePath}", "Export Successful");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Export failed: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Export failed: {ex.Message}", "Error");
            }
        }

        public static void ImportFence()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Starting fence import process");

                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string exportsDir = Path.Combine(exeDir, "Exports");

                var openDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Fence Files|*.fence",
                    DefaultExt = ".fence",
                    InitialDirectory = Directory.Exists(exportsDir) ? exportsDir : exeDir,
                    Title = "Select Fence Export File"
                };

                if (openDialog.ShowDialog() != true) return;

                string selectedFile = openDialog.FileName;
                string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);

                try
                {
                    ZipFile.ExtractToDirectory(selectedFile, tempDir);

                    string fenceJsonPath = Path.Combine(tempDir, "fence.json");
                    if (!File.Exists(fenceJsonPath)) throw new FileNotFoundException("Invalid export: missing fence.json");

                    string jsonContent = File.ReadAllText(fenceJsonPath);
                    dynamic importedFence = JsonConvert.DeserializeObject<JObject>(jsonContent);
                    if (importedFence == null) throw new InvalidDataException("Invalid JSON data");

                    // New ID to avoid conflicts
                    string newId = Guid.NewGuid().ToString();
                    importedFence["Id"] = newId;

                    // Handle Shortcuts (TAB-AWARE FIX)
                    if (importedFence.ItemsType?.ToString() == "Data")
                    {
                        string sourceShortcuts = Path.Combine(tempDir, "Shortcuts");

                        // FIX: Target current Profile Shortcuts folder
                        string destShortcuts = ProfileManager.GetProfileFilePath("Shortcuts");

                        if (Directory.Exists(sourceShortcuts))
                        {
                            if (!Directory.Exists(destShortcuts)) Directory.CreateDirectory(destShortcuts);

                            foreach (string srcPath in Directory.GetFiles(sourceShortcuts))
                            {
                                string fileName = Path.GetFileName(srcPath);
                                string destPath = Path.Combine(destShortcuts, fileName);

                                // Handle collisions
                                int counter = 1;
                                while (File.Exists(destPath))
                                {
                                    string tempName = $"{Path.GetFileNameWithoutExtension(fileName)} ({counter++}){Path.GetExtension(fileName)}";
                                    destPath = Path.Combine(destShortcuts, tempName);
                                }

                                File.Copy(srcPath, destPath);
                                string finalFileName = Path.GetFileName(destPath);

                                // Update References Helper
                                void UpdateReferences(JArray items)
                                {
                                    if (items == null) return;
                                    // Find items that matched the ORIGINAL filename (fileName)
                                    // Update them to the NEW filename (finalFileName)
                                    foreach (var item in items)
                                    {
                                        string itemFile = Path.GetFileName(item["Filename"]?.ToString() ?? "");
                                        if (itemFile.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            item["Filename"] = Path.Combine("Shortcuts", finalFileName);
                                        }
                                    }
                                }

                                // 1. Update Main Items
                                UpdateReferences(importedFence.Items as JArray);

                                // 2. Update Tab Items (The Fix)
                                var tabs = importedFence.Tabs as JArray;
                                if (tabs != null)
                                {
                                    foreach (var tab in tabs)
                                    {
                                        UpdateReferences(tab["Items"] as JArray);
                                    }
                                }
                            }
                        }
                    }

                    // Add and Create
                    var fenceData = FenceManager.GetFenceData();
                    fenceData.Add(importedFence);
                    FenceManager.CreateFence(importedFence, new TargetChecker(1000));
                    FenceDataManager.SaveFenceData();

                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Fence imported successfully!", "Import Complete");
                }
                finally
                {
                    if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Import failed: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to import fence: {ex.Message}", "Import Error");
            }
        }

        #endregion

        #region Deletion Backup Management

        public static void BackupDeletedFence(dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Creating deletion backup for fence: {fence.Title}");

                // FIX: Use Profile Path
                _lastDeletedFolderPath = ProfileManager.GetProfileFilePath("Last Fence Deleted");

                if (!Directory.Exists(_lastDeletedFolderPath))
                {
                    Directory.CreateDirectory(_lastDeletedFolderPath);
                }

                // Clear previous backup files
                foreach (var file in Directory.GetFiles(_lastDeletedFolderPath))
                {
                    File.Delete(file);
                }

                _lastDeletedFence = fence;
                _isRestoreAvailable = true;

                // Backup shortcuts for Data fences
                if (fence.ItemsType?.ToString() == "Data")
                {
                    int backedUpShortcuts = 0;

                    void BackupItems(JArray items)
                    {
                        if (items == null) return;
                        foreach (var item in items)
                        {
                            string itemFilePath = item["Filename"]?.ToString();
                            if (!string.IsNullOrEmpty(itemFilePath))
                            {
                                string fullSourcePath = Path.IsPathRooted(itemFilePath)
                                    ? itemFilePath
                                    : Path.Combine(ProfileManager.CurrentProfileDir, itemFilePath);

                                if (File.Exists(fullSourcePath))
                                {
                                    string destPath = Path.Combine(_lastDeletedFolderPath, Path.GetFileName(itemFilePath));
                                    if (!File.Exists(destPath))
                                    {
                                        File.Copy(fullSourcePath, destPath, true);
                                        backedUpShortcuts++;
                                    }
                                }
                            }
                        }
                    }

                    // 1. Backup Main Items
                    var mainItems = fence.Items as JArray;
                    if (mainItems != null) BackupItems(mainItems);

                    // 2. Backup Tab Items
                    bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                    if (tabsEnabled)
                    {
                        var tabs = fence.Tabs as JArray;
                        if (tabs != null)
                        {
                            foreach (var tab in tabs)
                            {
                                var tabItems = tab["Items"] as JArray;
                                if (tabItems != null) BackupItems(tabItems);
                            }
                        }
                    }
                }

                string fenceJsonPath = Path.Combine(_lastDeletedFolderPath, "fence.json");
                File.WriteAllText(fenceJsonPath, JsonConvert.SerializeObject(fence, Formatting.Indented));

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Deletion backup completed for fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Error creating deletion backup: {ex.Message}");
            }
        }

        public static void CleanLastDeletedFolder()
        {
            try
            {
                _lastDeletedFolderPath = ProfileManager.GetProfileFilePath("Last Fence Deleted");

                if (Directory.Exists(_lastDeletedFolderPath))
                {
                    foreach (var file in Directory.GetFiles(_lastDeletedFolderPath))
                    {
                        File.Delete(file);
                    }
                }
                else
                {
                    Directory.CreateDirectory(_lastDeletedFolderPath);
                }

                _isRestoreAvailable = false;
                _lastDeletedFence = null;
                FenceManager.UpdateAllHeartContextMenus();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Error cleaning last deleted folder: {ex.Message}");
            }
        }

        #endregion

        #region Helper Methods

        public static void CopyDirectory(string sourceDir, string destDir)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(sourceDir);
                DirectoryInfo[] dirs = dir.GetDirectories();

                Directory.CreateDirectory(destDir);

                foreach (FileInfo file in dir.GetFiles())
                {
                    string targetFilePath = Path.Combine(destDir, file.Name);
                    file.CopyTo(targetFilePath, false);
                }

                foreach (DirectoryInfo subDir in dirs)
                {
                    string newDestDir = Path.Combine(destDir, subDir.Name);
                    CopyDirectory(subDir.FullName, newDestDir);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Error copying directory {sourceDir}: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Auto-Backup & Core Logic

        // Called at startup to schedule the daily auto-backup
        public static void InitializeAutoBackup()
        {
            if (!SettingsManager.EnableAutoBackup) return;

            // Check if already ran today FOR THE CURRENT PROFILE
            if (SettingsManager.LastAutoBackupDate.Date == DateTime.Today)
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Auto-backup skipped: Already ran today.");
                return;
            }

            _autoBackupTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(5)
            };
            _autoBackupTimer.Tick += (s, e) =>
            {
                _autoBackupTimer.Stop();
                PerformAutoBackup();
            };
            _autoBackupTimer.Start();
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Auto-backup scheduled for 5 minutes from now.");
        }

        private static void PerformAutoBackup()
        {
            try
            {
                // Double-check setting (in case it changed since startup)
                if (!SettingsManager.EnableAutoBackup) return;

                string timestamp = DateTime.Now.ToString("yyMMddHHmm");
                string backupFolderName = $"{timestamp}_backup_auto";

                // This calls CreateBackup which resolves Profile Path dynamically
                CreateBackup(backupFolderName, silent: true);

                // Update last run date (saves to Active Profile or Master)
                SettingsManager.LastAutoBackupDate = DateTime.Now;
                SettingsManager.SaveSettings();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Auto-backup completed: {backupFolderName}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Auto-backup failed: {ex.Message}");
            }
        }

        // Refactored Helper: Centralizes the actual backup work
        public static void CreateBackup(string folderName, bool silent = false)
        {
            try
            {
                // SOURCE: Profile Directory (Dynamic)
                string jsonFilePath = ProfileManager.GetProfileFilePath("fences.json");
                string optionsFilePath = ProfileManager.GetProfileFilePath("options.json");
                string shortcutsFolderPath = ProfileManager.GetProfileFilePath("Shortcuts");

                // DEST: Profile Directory -> Backups
                string backupsFolderPath = ProfileManager.GetProfileFilePath("Backups");
                string backupFolderPath = Path.Combine(backupsFolderPath, folderName);

                if (!Directory.Exists(backupsFolderPath))
                {
                    Directory.CreateDirectory(backupsFolderPath);
                }
                Directory.CreateDirectory(backupFolderPath);

                // 1. Copy fences.json
                string backupJsonFilePath = Path.Combine(backupFolderPath, "fences.json");
                if (File.Exists(jsonFilePath))
                {
                    File.Copy(jsonFilePath, backupJsonFilePath, true);
                }

                // 2. Copy options.json
                string backupOptionsFilePath = Path.Combine(backupFolderPath, "options.json");
                if (File.Exists(optionsFilePath))
                {
                    File.Copy(optionsFilePath, backupOptionsFilePath, true);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Backed up options.json");
                }

                // 3. Copy Shortcuts Folder
                string backupShortcutsFolderPath = Path.Combine(backupFolderPath, "Shortcuts");
                if (Directory.Exists(shortcutsFolderPath))
                {
                    Directory.CreateDirectory(backupShortcutsFolderPath);
                    CopyDirectory(shortcutsFolderPath, backupShortcutsFolderPath);
                }

                if (!silent)
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Backup completed successfully.", "Backup");
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Manual backup finished: {backupFolderPath}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"CreateBackup error: {ex.Message}");
                if (!silent)
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"An error occurred during backup: {ex.Message}", "Error");
                }
            }
        }

        public static void RestoreFromBackup(string backupFolder)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Starting restore from backup: {backupFolder}");

                string backupFencesPath = Path.Combine(backupFolder, "fences.json");
                string backupShortcutsPath = Path.Combine(backupFolder, "Shortcuts");

                if (!File.Exists(backupFencesPath) || !Directory.Exists(backupShortcutsPath))
                {
                    string errorMsg = "Invalid backup folder - missing required files.";
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(errorMsg, "Restore Error");
                    return;
                }

                bool restartRequired = false;

                // 2. Handle options.json
                string backupOptionsPath = Path.Combine(backupFolder, "options.json");
                if (File.Exists(backupOptionsPath))
                {
                    bool restoreSettings = MessageBoxesManager.ShowCustomYesNoMessageBox(
                        "This backup contains configuration settings (options.json).\n\n" +
                        "Do you want to restore your global settings as well?",
                        "Restore Settings");

                    if (restoreSettings)
                    {
                        try
                        {
                            string currentOptionsPath = ProfileManager.GetProfileFilePath("options.json");
                            File.Copy(backupOptionsPath, currentOptionsPath, true);
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Restored options.json");
                            restartRequired = true;
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.ImportExport, $"Failed to restore options.json: {ex.Message}");
                        }
                    }
                }

                // 3. Clear & Copy Data
                string currentFencesPath = ProfileManager.GetProfileFilePath("fences.json");
                string currentShortcutsPath = ProfileManager.GetProfileFilePath("Shortcuts");

                var fenceData = FenceManager.GetFenceData();
                fenceData?.Clear();

                File.Copy(backupFencesPath, currentFencesPath, true);

                if (Directory.Exists(currentShortcutsPath))
                {
                    Directory.Delete(currentShortcutsPath, true);
                }
                Directory.CreateDirectory(currentShortcutsPath);
                BackupManager.CopyDirectory(backupShortcutsPath, currentShortcutsPath);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Files restored successfully.");

                if (restartRequired)
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(
                        "Global settings have been restored.\nThe application will now restart to apply changes.",
                        "Restart Required");

                    string appPath = Process.GetCurrentProcess().MainModule.FileName;
                    Process.Start(appPath);
                    Environment.Exit(0);
                }
                else
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Restore completed successfully.", "Restore");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Restore failed: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Restore failed: {ex.Message}", "Error");
            }
        }
        #endregion
    }
}