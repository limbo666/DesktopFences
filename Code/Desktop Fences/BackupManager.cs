﻿using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Desktop_Fences
{
    // <summary>
    // Manages backup operations for Desktop Fences data, including the fences configuration file
    // and associated shortcut files. Backups are stored in timestamped folders for easy recovery.
    // Refactored to centralize all backup/restore functionality.

    public static class BackupManager
    {
        #region Private Fields
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
        public static void BackupData()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Starting complete data backup");

                // Get the directory of the executing assembly
                string exePath = Assembly.GetEntryAssembly().Location;
                string exeDir = Path.GetDirectoryName(exePath);

                // Define source paths for fences.json and Shortcuts folder
                string jsonFilePath = Path.Combine(exeDir, "fences.json");
                string shortcutsFolderPath = Path.Combine(exeDir, "Shortcuts");

                // Define the destination "Backups" folder
                string backupsFolderPath = Path.Combine(exeDir, "Backups");
                if (!Directory.Exists(backupsFolderPath))
                {
                    Directory.CreateDirectory(backupsFolderPath); // Create Backups folder if it doesn't exist
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Created Backups directory: {backupsFolderPath}");
                }

                // Create a new backup folder with the current date and time (e.g., "2503181234_backup")
                string backupFolderName = DateTime.Now.ToString("yyMMddHHmm") + "_backup";
                string backupFolderPath = Path.Combine(backupsFolderPath, backupFolderName);
                Directory.CreateDirectory(backupFolderPath);

                // Copy the fences.json file to the backup folder
                string backupJsonFilePath = Path.Combine(backupFolderPath, "fences.json");
                if (File.Exists(jsonFilePath))
                {
                    File.Copy(jsonFilePath, backupJsonFilePath, true); // Overwrite if exists
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Backed up fences.json to: {backupJsonFilePath}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.ImportExport, "fences.json not found for backup");
                }

                // Copy the entire "Shortcuts" folder, if it exists
                string backupShortcutsFolderPath = Path.Combine(backupFolderPath, "Shortcuts");
                if (Directory.Exists(shortcutsFolderPath))
                {
                    Directory.CreateDirectory(backupShortcutsFolderPath);

                    // Use our helper method to copy directory recursively
                    CopyDirectory(shortcutsFolderPath, backupShortcutsFolderPath);

                    var fileCount = Directory.GetFiles(backupShortcutsFolderPath, "*.*", SearchOption.AllDirectories).Length;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Backed up {fileCount} files from Shortcuts folder");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Shortcuts folder not found, skipping shortcut backup");
                }

                // Notify the user of successful backup
                TrayManager.Instance.ShowOKOnlyMessageBoxForm("Backup completed successfully.", "Backup");
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Complete data backup finished successfully: {backupFolderPath}");
            }
            catch (Exception ex)
            {
                // Handle any errors during the backup process and inform the user
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Backup failed: {ex.Message}\nStack trace: {ex.StackTrace}");
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"An error occurred during backup: {ex.Message}", "Error");
            }
        }
        #endregion

        #region Restore Operations

        // Restores fences and shortcuts from a backup folder with comprehensive validation

        // <param name="backupFolder">Path to the backup folder to restore from</param>
        public static void RestoreFromBackup(string backupFolder)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Starting restore from backup: {backupFolder}");

                // Validate backup folder structure
                string backupFencesPath = Path.Combine(backupFolder, "fences.json");
                string backupShortcutsPath = Path.Combine(backupFolder, "Shortcuts");

                if (!File.Exists(backupFencesPath) || !Directory.Exists(backupShortcutsPath))
                {
                    string errorMsg = "Invalid backup folder - missing required files (fences.json or Shortcuts folder)";
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, errorMsg);
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm(errorMsg, "Restore Error");
                    return;
                }

                // Get current application directory
                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string currentFencesPath = Path.Combine(exeDir, "fences.json");
                string currentShortcutsPath = Path.Combine(exeDir, "Shortcuts");

                // Clear existing data structures (important for clean restore)
                var fenceData = FenceManager.GetFenceData();
                fenceData?.Clear();

                // Restore fences.json
                File.Copy(backupFencesPath, currentFencesPath, true);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Restored fences.json");

                // Restore shortcuts folder
                if (Directory.Exists(currentShortcutsPath))
                {
                    Directory.Delete(currentShortcutsPath, true);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, "Cleared existing shortcuts");
                }

                Directory.CreateDirectory(currentShortcutsPath);
                CopyDirectory(backupShortcutsPath, currentShortcutsPath);

                var restoredFileCount = Directory.GetFiles(currentShortcutsPath, "*.*", SearchOption.AllDirectories).Length;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Restored {restoredFileCount} shortcut files");

                // Important: Reload fences using FenceManager's method to ensure proper initialization
                FenceManager.LoadAndCreateFences(new TargetChecker(1000));

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Restore from backup completed successfully");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Restore failed: {ex.Message}\nStack trace: {ex.StackTrace}");
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Restore failed: {ex.Message}", "Error");
                throw; // Re-throw to allow calling code to handle
            }
        }


        // Restores the last deleted fence from backup with validation

        public static void RestoreLastDeletedFence()
        {
            try
            {
                if (!_isRestoreAvailable || string.IsNullOrEmpty(_lastDeletedFolderPath) || _lastDeletedFence == null)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "No fence available to restore");
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm("No fence to restore", "Restore");
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
                FenceManager.SaveFenceData();

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
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Error restoring fence: {ex.Message}", "Restore Error");
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
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Fence exported to:\n{fencePath}", "Export Successful");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Export failed: {ex.Message}\nStack trace: {ex.StackTrace}");
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Export failed: {ex.Message}", "Error");
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
                    FenceManager.SaveFenceData();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Fence '{importedFence.Title}' imported successfully");
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm("Fence imported successfully!", "Import Complete");
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
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to import fence: {ex.Message}", "Import Error");
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
    }
}