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
    // <summary>
    // Manages backup operations for Desktop Fences data, including the fences configuration file
    // and associated shortcut files. Backups are stored in timestamped folders for easy recovery.
    // Refactored to centralize all backup/restore functionality.

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
        #endregion

        #region Public Properties

        // Gets whether a restore operation is available for the last deleted fence

        public static bool IsRestoreAvailable => _isRestoreAvailable;


        // Gets the path to the last deleted fence backup folder

        public static string LastDeletedFolderPath => _lastDeletedFolderPath;
        #endregion

        #region Existing Backup Method (Enhanced)

        // Creates a complete backup of all fences and shortcuts with enhanced logging
        //public static void BackupData()
        //{
        //    try
        //    {
        //        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Starting complete data backup");

        //        // Get the directory of the executing assembly
        //        string exePath = Assembly.GetEntryAssembly().Location;
        //        string exeDir = Path.GetDirectoryName(exePath);

        //        // Define source paths for fences.json and Shortcuts folder
        //        string jsonFilePath = Path.Combine(exeDir, "fences.json");
        //        string shortcutsFolderPath = Path.Combine(exeDir, "Shortcuts");

        //        // Define the destination "Backups" folder
        //        string backupsFolderPath = Path.Combine(exeDir, "Backups");
        //        if (!Directory.Exists(backupsFolderPath))
        //        {
        //            Directory.CreateDirectory(backupsFolderPath); // Create Backups folder if it doesn't exist
        //            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Created Backups directory: {backupsFolderPath}");
        //        }

        //        // Create a new backup folder with the current date and time (e.g., "2503181234_backup")
        //        string backupFolderName = DateTime.Now.ToString("yyMMddHHmm") + "_backup";
        //        string backupFolderPath = Path.Combine(backupsFolderPath, backupFolderName);
        //        Directory.CreateDirectory(backupFolderPath);

        //        // Copy the fences.json file to the backup folder
        //        string backupJsonFilePath = Path.Combine(backupFolderPath, "fences.json");
        //        if (File.Exists(jsonFilePath))
        //        {
        //            File.Copy(jsonFilePath, backupJsonFilePath, true); // Overwrite if exists
        //            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Backed up fences.json to: {backupJsonFilePath}");
        //        }
        //        else
        //        {
        //            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.ImportExport, "fences.json not found for backup");
        //        }

        //        // Copy the entire "Shortcuts" folder, if it exists
        //        string backupShortcutsFolderPath = Path.Combine(backupFolderPath, "Shortcuts");
        //        if (Directory.Exists(shortcutsFolderPath))
        //        {
        //            Directory.CreateDirectory(backupShortcutsFolderPath);

        //            // Use our helper method to copy directory recursively
        //            CopyDirectory(shortcutsFolderPath, backupShortcutsFolderPath);

        //            var fileCount = Directory.GetFiles(backupShortcutsFolderPath, "*.*", SearchOption.AllDirectories).Length;
        //            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Backed up {fileCount} files from Shortcuts folder");
        //        }
        //        else
        //        {
        //            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Shortcuts folder not found, skipping shortcut backup");
        //        }

        //        // Notify the user of successful backup
        //        MessageBoxesManager.ShowOKOnlyMessageBoxForm("Backup completed successfully.", "Backup");
        //        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Complete data backup finished successfully: {backupFolderPath}");
        //    }
        //    catch (Exception ex)
        //    {
        //        // Handle any errors during the backup process and inform the user
        //        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Backup failed: {ex.Message}\nStack trace: {ex.StackTrace}");
        //        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"An error occurred during backup: {ex.Message}", "Error");
        //    }
        //}

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

        // Restores fences and shortcuts from a backup folder with comprehensive validation

        // <param name="backupFolder">Path to the backup folder to restore from</param>
    

        // Restores the last deleted fence from backup with validation

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
                    string shortcutsDir = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Shortcuts");

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

        // Exports a single fence to a .fence file with all associated shortcuts

        // <param name="fence">The fence to export</param>
        public static void ExportFence(dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Starting export of fence: {fence.Title}");

                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string fenceTitle = fence.Title.ToString();

                // Sanitize folder name for filesystem compatibility
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    fenceTitle = fenceTitle.Replace(c, '_');
                }

                string exportFolder = Path.Combine(exeDir, "Exports", fenceTitle);
                string fencePath = Path.Combine(exeDir, "Exports", $"{fenceTitle}.fence");

                // Ensure exports directory exists
                string exportsDir = Path.Combine(exeDir, "Exports");
                if (!Directory.Exists(exportsDir))
                {
                    Directory.CreateDirectory(exportsDir);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Created Exports directory: {exportsDir}");
                }

                // Cleanup existing export files
                if (Directory.Exists(exportFolder))
                {
                    Directory.Delete(exportFolder, true);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Cleaned up existing export folder: {exportFolder}");
                }
                if (File.Exists(fencePath))
                {
                    File.Delete(fencePath);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Cleaned up existing fence file: {fencePath}");
                }

                Directory.CreateDirectory(exportFolder);

                // Save fence data as JSON
                string fenceJson = JsonConvert.SerializeObject(fence, Formatting.Indented);
                File.WriteAllText(Path.Combine(exportFolder, "fence.json"), fenceJson);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Exported fence metadata to fence.json");

                // Copy shortcuts for Data fences
                if (fence.ItemsType?.ToString() == "Data")
                {
                    string shortcutsDestDir = Path.Combine(exportFolder, "Shortcuts");
                    Directory.CreateDirectory(shortcutsDestDir);

                    int copiedShortcuts = 0;
                    foreach (var item in fence.Items)
                    {
                        string filename = item.Filename?.ToString();
                        if (!string.IsNullOrEmpty(filename))
                        {
                            string sourcePath = Path.Combine(exeDir, filename);
                            if (File.Exists(sourcePath))
                            {
                                string destName = Path.GetFileName(filename);
                                File.Copy(sourcePath, Path.Combine(shortcutsDestDir, destName));
                                copiedShortcuts++;
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.ImportExport, $"Shortcut file not found for export: {sourcePath}");
                            }
                        }
                    }
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Copied {copiedShortcuts} shortcuts to export");
                }

                // Create zip file and cleanup temporary folder
                ZipFile.CreateFromDirectory(exportFolder, fencePath);
                Directory.Delete(exportFolder, true);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Fence exported successfully to: {fencePath}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Fence exported to:\n{fencePath}", "Export Successful");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Export failed: {ex.Message}\nStack trace: {ex.StackTrace}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Export failed: {ex.Message}", "Error");
            }
        }


        // Imports a fence from a .fence file with comprehensive validation

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

                if (openDialog.ShowDialog() != true)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Import cancelled by user");
                    return;
                }

                string selectedFile = openDialog.FileName;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Selected file for import: {selectedFile}");

                // Create temporary directory for extraction
                string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);

                try
                {
                    // Extract ZIP contents
                    ZipFile.ExtractToDirectory(selectedFile, tempDir);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Extracted import file to: {tempDir}");

                    // Validate export structure
                    string fenceJsonPath = Path.Combine(tempDir, "fence.json");
                    if (!File.Exists(fenceJsonPath))
                    {
                        throw new FileNotFoundException("Missing fence.json in export file - invalid fence export");
                    }

                    // Deserialize fence data with validation
                    string jsonContent = File.ReadAllText(fenceJsonPath);
                    dynamic importedFence = JsonConvert.DeserializeObject<JObject>(jsonContent);

                    if (importedFence == null)
                    {
                        throw new InvalidDataException("Invalid fence data in export file");
                    }

                    // Generate new ID to prevent conflicts with existing fences
                    string newId = Guid.NewGuid().ToString();
                    importedFence["Id"] = newId;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Assigned new fence ID: {newId}");

                    // Handle shortcuts for Data fences
                    if (importedFence.ItemsType?.ToString() == "Data")
                    {
                        string sourceShortcuts = Path.Combine(tempDir, "Shortcuts");
                        string destShortcuts = Path.Combine(exeDir, "Shortcuts");

                        if (Directory.Exists(sourceShortcuts))
                        {
                            // Ensure destination shortcuts directory exists
                            Directory.CreateDirectory(destShortcuts);

                            int importedShortcuts = 0;
                            foreach (string srcPath in Directory.GetFiles(sourceShortcuts))
                            {
                                string fileName = Path.GetFileName(srcPath);
                                string destPath = Path.Combine(destShortcuts, fileName);

                                // Handle duplicate filenames by appending counter
                                int counter = 1;
                                while (File.Exists(destPath))
                                {
                                    string tempName = $"{Path.GetFileNameWithoutExtension(fileName)} ({counter++}){Path.GetExtension(fileName)}";
                                    destPath = Path.Combine(destShortcuts, tempName);
                                }

                                File.Copy(srcPath, destPath);
                                importedShortcuts++;

                                // Update shortcut references in fence items to reflect any name changes
                                var items = importedFence.Items as JArray;
                                if (items != null)
                                {
                                    foreach (var item in items.Where(i => i["Filename"]?.ToString() == fileName))
                                    {
                                        item["Filename"] = Path.Combine("Shortcuts", Path.GetFileName(destPath));
                                    }
                                }
                            }
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Imported {importedShortcuts} shortcuts");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "No shortcuts folder found in import file");
                        }
                    }

                    // Add to fence data and create the fence
                    var fenceData = FenceManager.GetFenceData();
                    fenceData.Add(importedFence);
                    FenceManager.CreateFence(importedFence, new TargetChecker(1000));
                    FenceDataManager.SaveFenceData();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Fence '{importedFence.Title}' imported successfully");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Fence imported successfully!", "Import Complete");
                }
                finally
                {
                    // Cleanup temporary files
                    if (Directory.Exists(tempDir))
                    {
                        Directory.Delete(tempDir, true);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Cleaned up temporary directory: {tempDir}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Fence import failed: {ex.Message}\nStack trace: {ex.StackTrace}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to import fence: {ex.Message}", "Import Error");
            }
        }
        #endregion

        #region Deletion Backup Management

        // Backs up a fence that's being deleted for potential restoration

        // <param name="fence">The fence being deleted</param>
        public static void BackupDeletedFence(dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Creating deletion backup for fence: {fence.Title}");

                // Ensure the backup folder exists
                _lastDeletedFolderPath = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Last Fence Deleted");
                if (!Directory.Exists(_lastDeletedFolderPath))
                {
                    Directory.CreateDirectory(_lastDeletedFolderPath);
                }

                // Clear previous backup files
                foreach (var file in Directory.GetFiles(_lastDeletedFolderPath))
                {
                    File.Delete(file);
                }

                // Store fence reference and mark restore as available
                _lastDeletedFence = fence;
                _isRestoreAvailable = true;

                // Backup shortcuts for Data fences
                if (fence.ItemsType?.ToString() == "Data")
                {
                    var items = fence.Items as JArray;
                    if (items != null)
                    {
                        int backedUpShortcuts = 0;
                        foreach (var item in items)
                        {
                            string itemFilePath = item["Filename"]?.ToString();
                            if (!string.IsNullOrEmpty(itemFilePath) && File.Exists(itemFilePath))
                            {
                                string shortcutPath = Path.Combine(_lastDeletedFolderPath, Path.GetFileName(itemFilePath));
                                File.Copy(itemFilePath, shortcutPath, true);
                                backedUpShortcuts++;
                            }
                            else if (!string.IsNullOrEmpty(itemFilePath))
                            {
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Skipped backing up missing file: {itemFilePath}");
                            }
                        }
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Backed up {backedUpShortcuts} shortcuts for deletion");
                    }
                }

                // Save fence info to JSON for complete restoration
                string fenceJsonPath = Path.Combine(_lastDeletedFolderPath, "fence.json");
                File.WriteAllText(fenceJsonPath, JsonConvert.SerializeObject(fence, Formatting.Indented));

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Deletion backup completed for fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Error creating deletion backup: {ex.Message}");
                // Don't throw - deletion should continue even if backup fails
            }
        }


        // Cleans the last deleted fence backup folder and resets restore availability

        public static void CleanLastDeletedFolder()
        {
            try
            {
                _lastDeletedFolderPath = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Last Fence Deleted");

                if (Directory.Exists(_lastDeletedFolderPath))
                {
                    // Clear all files in the backup folder
                    foreach (var file in Directory.GetFiles(_lastDeletedFolderPath))
                    {
                        File.Delete(file);
                    }
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Cleaned last deleted folder contents");
                }
                else
                {
                    // Create the directory if it doesn't exist
                    Directory.CreateDirectory(_lastDeletedFolderPath);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Created last deleted folder: {_lastDeletedFolderPath}");
                }

                // Reset backup state
                _isRestoreAvailable = false;
                _lastDeletedFence = null;

                // Update heart context menus to reflect that no restore is available
                FenceManager.UpdateAllHeartContextMenus();

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Reset deletion backup state");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Error cleaning last deleted folder: {ex.Message}");
            }
        }
        #endregion

        #region Helper Methods

        // Recursively copies a directory and all its contents

        // <param name="sourceDir">Source directory path</param>
        // <param name="destDir">Destination directory path</param>
        public static void CopyDirectory(string sourceDir, string destDir)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(sourceDir);
                DirectoryInfo[] dirs = dir.GetDirectories();

                // Create destination directory
                Directory.CreateDirectory(destDir);

                // Copy all files in the current directory
                foreach (FileInfo file in dir.GetFiles())
                {
                    string targetFilePath = Path.Combine(destDir, file.Name);
                    file.CopyTo(targetFilePath, false);
                }

                // Recursively copy all subdirectories
                foreach (DirectoryInfo subDir in dirs)
                {
                    string newDestDir = Path.Combine(destDir, subDir.Name);
                    CopyDirectory(subDir.FullName, newDestDir);
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Copied directory: {sourceDir} -> {destDir}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Error copying directory {sourceDir}: {ex.Message}");
                throw; // Re-throw to allow caller to handle
            }
        }
        #endregion



        #region Auto-Backup & Core Logic

        private static System.Windows.Threading.DispatcherTimer _autoBackupTimer;

        // Called at startup to schedule the daily auto-backup
        public static void InitializeAutoBackup()
        {
            if (!SettingsManager.EnableAutoBackup) return;

            // Check if already ran today to prevent spamming backups on every restart
            if (SettingsManager.LastAutoBackupDate.Date == DateTime.Today)
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Auto-backup skipped: Already ran today.");
                return;
            }

            // Schedule for 5 minutes later to avoid slowing down startup
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
                // Double-check setting
                if (!SettingsManager.EnableAutoBackup) return;

                // Format: [datetime]_backup_auto
                string timestamp = DateTime.Now.ToString("yyMMddHHmm");
                string backupFolderName = $"{timestamp}_backup_auto";

                // Run backup silently (no message box)
                CreateBackup(backupFolderName, silent: true);

                // Update last run date
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
        // Includes options.json now
        public static void CreateBackup(string folderName, bool silent = false)
        {
            try
            {
                string exePath = Assembly.GetEntryAssembly().Location;
                string exeDir = Path.GetDirectoryName(exePath);

                // Source Paths
                string jsonFilePath = Path.Combine(exeDir, "fences.json");
                string optionsFilePath = Path.Combine(exeDir, "options.json"); // NEW
                string shortcutsFolderPath = Path.Combine(exeDir, "Shortcuts");

                // Dest Paths
                string backupsFolderPath = Path.Combine(exeDir, "Backups");
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

                // 2. Copy options.json (NEW)
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

        // Restores fences, shortcuts, AND optionally settings from a backup
        public static void RestoreFromBackup(string backupFolder)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Starting restore from backup: {backupFolder}");

                // 1. Validate basic structure
                string backupFencesPath = Path.Combine(backupFolder, "fences.json");
                string backupShortcutsPath = Path.Combine(backupFolder, "Shortcuts");

                if (!File.Exists(backupFencesPath) || !Directory.Exists(backupShortcutsPath))
                {
                    string errorMsg = "Invalid backup folder - missing required files.";
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(errorMsg, "Restore Error");
                    return;
                }

                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
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
                            string currentOptionsPath = Path.Combine(exeDir, "options.json");
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
                string currentFencesPath = Path.Combine(exeDir, "fences.json");
                string currentShortcutsPath = Path.Combine(exeDir, "Shortcuts");

                var fenceData = FenceManager.GetFenceData();
                fenceData?.Clear();

                File.Copy(backupFencesPath, currentFencesPath, true);

                if (Directory.Exists(currentShortcutsPath))
                {
                    Directory.Delete(currentShortcutsPath, true);
                }
                Directory.CreateDirectory(currentShortcutsPath);
                CopyDirectory(backupShortcutsPath, currentShortcutsPath);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Files restored successfully.");

                // 4. DECISION TIME
                if (restartRequired)
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(
                        "Global settings have been restored.\nThe application will now restart to apply changes.",
                        "Restart Required");

                    // FIX: Use Environment.Exit(0) to stop execution IMMEDIATELY.
                    // Application.Current.Shutdown() allows subsequent code (like TrayManager.reloadAllFences)
                    // to run while the app is dying, causing NullReferenceException.
                    string appPath = Process.GetCurrentProcess().MainModule.FileName;
                    Process.Start(appPath);
                    Environment.Exit(0);
                }
                else
                {
                    // Optimization: We DO NOT reload fences here anymore.
                    // The caller (OptionsFormManager) calls TrayManager.reloadAllFences() immediately after we return.
                    // Removing this prevents Double-Loading.
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