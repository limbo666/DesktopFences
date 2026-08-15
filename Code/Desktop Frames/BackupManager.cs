using Desktop_Frames.Localization;
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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;

namespace Desktop_Frames
{
	/// <summary>
	/// Manages backup operations for Desktop Frames data, including the frames configuration file
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

        // Manage last deleted frame restoration
        private static string _lastDeletedFolderPath;
        private static dynamic _lastDeletedFrame;
        private static bool _isRestoreAvailable;
        private static System.Windows.Threading.DispatcherTimer _autoBackupTimer;
		#endregion

		#region Public Properties
		// Gets whether a restore operation is available for the last deleted frame
		public static bool IsRestoreAvailable => _isRestoreAvailable;

		// Gets the path to the last deleted frame backup folder
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
            string backupName = DateTime.Now.ToString("yyMMddHHmmss") + "_backup";
            CreateBackup(backupName, silent: false);
        }

        #endregion

        #region Restore Operations

        public static void RestoreLastDeletedFrame()
        {
            try
            {
                if (!_isRestoreAvailable || string.IsNullOrEmpty(_lastDeletedFolderPath) || _lastDeletedFrame == null)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "No frame available to restore");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("No frame to restore", "Restore");
                    return;
                }

				LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Restoring last deleted frame: {_lastDeletedFrame.Title}");

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

				// Restore frame data to the main frame collection
				var FrameData = Framemanager.GetFrameData();
                FrameData.Add(_lastDeletedFrame);
                FrameDataManager.SaveFrameData();

				// Create the frame UI
				Framemanager.CreateFrame(_lastDeletedFrame, new TargetChecker(1000));

                // Clear backup state
                _lastDeletedFrame = null;
                _isRestoreAvailable = false;

                // Update heart context menus to reflect restored state
                Framemanager.UpdateAllHeartContextMenus();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Last deleted frame restored successfully");
            }
			catch (Exception ex)
			{
				LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Error restoring last deleted frame: {ex.Message}");
				MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error restoring frame: {ex.Message}", "Restore Error");
			}
		}
        #endregion

        #region Frame Export/Import Operations

        public static void ExportFrame(dynamic frame)
        {
            Application.Current?.Dispatcher.Invoke(() => Mouse.OverrideCursor = Cursors.Wait);
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Starting export of frame: {frame.Title}");

                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string frameTitle = frame.Title.ToString();

                // Sanitize folder name
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    frameTitle = frameTitle.Replace(c, '_');
                }

                // --- BUG FIX: Windows prohibits trailing spaces and dots in folder names ---
                frameTitle = frameTitle.TrimEnd(' ', '.');

                // Fallback in case the name was completely stripped
                if (string.IsNullOrWhiteSpace(frameTitle))
                {
                    frameTitle = "ExportedFrame_" + Guid.NewGuid().ToString().Substring(0, 8);
                }
                // ---------------------------------------------------------------------------

                // Exports go to GLOBAL "Exports" folder (shared between profiles)
                string exportFolder = Path.Combine(exeDir, "Exports", frameTitle);
                string framePath = Path.Combine(exeDir, "Exports", $"{frameTitle}.frame"); // NEW FORMAT

                // Ensure exports directory exists
                string exportsDir = Path.Combine(exeDir, "Exports");
                if (!Directory.Exists(exportsDir)) Directory.CreateDirectory(exportsDir);

                // Cleanup previous runs
                if (Directory.Exists(exportFolder)) Directory.Delete(exportFolder, true);
                if (File.Exists(framePath)) File.Delete(framePath);

                Directory.CreateDirectory(exportFolder);

                // 1. Save Metadata
                string frameJson = JsonConvert.SerializeObject(frame, Formatting.Indented);
                File.WriteAllText(Path.Combine(exportFolder, "frame.json"), frameJson);

                // 2. Copy Shortcuts (TAB-AWARE FIX)
                if (frame.ItemsType?.ToString() == "Data")
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
                    var mainItems = frame.Items as JArray;
                    if (mainItems != null) CopyItems(mainItems);

                    // B. Copy Tab Items (The Fix)
                    bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";
                    if (tabsEnabled)
                    {
                        var tabs = frame.Tabs as JArray;
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
                ZipFile.CreateFromDirectory(exportFolder, framePath);
                Directory.Delete(exportFolder, true);

				LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Frame exported successfully: {framePath}");
				MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Frame exported to:\n{framePath}", "Export Successful");
			}
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Export failed: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Export failed: {ex.Message}", "Error");
            }
            finally
            {
                Application.Current?.Dispatcher.Invoke(() => Mouse.OverrideCursor = null);
            }
        }

        public static void ImportFrame()
        {
            Application.Current?.Dispatcher.Invoke(() => Mouse.OverrideCursor = Cursors.Wait);
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, "Starting frame import process");

                string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
				string exportsDir = Path.Combine(exeDir, "Exports");

                var openDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Frame Files|*.frame;*.fence", // SUPPORT BOTH EXTENSIONS
                    DefaultExt = ".frame", // DEFAULT TO NEW
                    InitialDirectory = Directory.Exists(exportsDir) ? exportsDir : exeDir,
                    Title = Strings.BackupSelectExportFile
                };

                if (openDialog.ShowDialog() != true) return;

                string selectedFile = openDialog.FileName;
                string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);

                try
                {
                    ZipFile.ExtractToDirectory(selectedFile, tempDir);

                    // Check for new (frame.json) and legacy (fence.json) internal files
                    string frameJsonPath = Path.Combine(tempDir, "frame.json");
                    if (!File.Exists(frameJsonPath))
                    {
                        frameJsonPath = Path.Combine(tempDir, "fence.json");
                    }
                    if (!File.Exists(frameJsonPath)) throw new FileNotFoundException("Invalid export: missing frame data file");

                    string jsonContent = File.ReadAllText(frameJsonPath);

                    // KISS SCRUBBER: Neutralize legacy ghost keys before the app touches it
                    jsonContent = jsonContent.Replace("\"FenceBorderColor\"", "\"FrameBorderColor\"");
                    jsonContent = jsonContent.Replace("\"frameBorderColor\"", "\"FrameBorderColor\"");
                    jsonContent = jsonContent.Replace("\"FenceBorderThickness\"", "\"FrameBorderThickness\"");
                    jsonContent = jsonContent.Replace("\"frameBorderThickness\"", "\"FrameBorderThickness\"");

                    dynamic importedFrame = JsonConvert.DeserializeObject<JObject>(jsonContent);
                    if (importedFrame == null) throw new InvalidDataException("Invalid JSON data");

                    // New ID to avoid conflicts
                    string newId = Guid.NewGuid().ToString();
                    importedFrame["Id"] = newId;

                    // Handle Shortcuts (TAB-AWARE FIX)
                    if (importedFrame.ItemsType?.ToString() == "Data")
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
                                UpdateReferences(importedFrame.Items as JArray);

                                // 2. Update Tab Items (The Fix)
                                var tabs = importedFrame.Tabs as JArray;
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
                    var FrameData = Framemanager.GetFrameData();
                    FrameData.Add(importedFrame);
                    Framemanager.CreateFrame(importedFrame, new TargetChecker(1000));
                    FrameDataManager.SaveFrameData();

                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Frame imported successfully!", "Import Complete");
                }
                finally
                {
                    if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Import failed: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to import frame: {ex.Message}", "Import Error");
            }
            finally
            {
                Application.Current?.Dispatcher.Invoke(() => Mouse.OverrideCursor = null);
            }
        }

        #endregion

        #region Deletion Backup Management

        public static void BackupDeletedFrame(dynamic frame)
		{
			try
			{
				LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Creating deletion backup for frame: {frame.Title}");

				// FIX: Use Profile Path
				_lastDeletedFolderPath = ProfileManager.GetProfileFilePath("Last Frame Deleted");

                if (!Directory.Exists(_lastDeletedFolderPath))
                {
                    Directory.CreateDirectory(_lastDeletedFolderPath);
                }

                // Clear previous backup files
                foreach (var file in Directory.GetFiles(_lastDeletedFolderPath))
                {
                    File.Delete(file);
                }

                _lastDeletedFrame = frame;
                _isRestoreAvailable = true;

                // Backup shortcuts for Data frames
                if (frame.ItemsType?.ToString() == "Data")
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
                    var mainItems = frame.Items as JArray;
                    if (mainItems != null) BackupItems(mainItems);

                    // 2. Backup Tab Items
                    bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";
                    if (tabsEnabled)
                    {
                        var tabs = frame.Tabs as JArray;
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

                string frameJsonPath = Path.Combine(_lastDeletedFolderPath, "frame.json");
                File.WriteAllText(frameJsonPath, JsonConvert.SerializeObject(frame, Formatting.Indented));

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Deletion backup completed for frame '{frame.Title}'");
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
                _lastDeletedFolderPath = ProfileManager.GetProfileFilePath("Last Frame Deleted");

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
                _lastDeletedFrame = null;
                Framemanager.UpdateAllHeartContextMenus();
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
                    // FIX: Allow overwrite to prevent crashes on rapid consecutive backups
                    file.CopyTo(targetFilePath, true);
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

                string timestamp = DateTime.Now.ToString("yyMMddHHmmss");
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
            // KISS: Instantly trigger the Windows "Wait/Processing" cursor
            Application.Current?.Dispatcher.Invoke(() => Mouse.OverrideCursor = Cursors.Wait);
            try
            {
                // SOURCE: Profile Directory (Dynamic)
                string jsonFilePath = ProfileManager.GetProfileFilePath("frames.json");
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

                // 1. Copy frames.json
                string backupJsonFilePath = Path.Combine(backupFolderPath, "frames.json");
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
            finally
            {
                // KISS: Guarantee the cursor returns to normal, even on failure
                Application.Current?.Dispatcher.Invoke(() => Mouse.OverrideCursor = null);
            }
        }
        public static void RestoreFromBackup(string backupFolder)
        {
            Application.Current?.Dispatcher.Invoke(() => Mouse.OverrideCursor = Cursors.Wait);
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Starting restore from backup: {backupFolder}");

                // Check for new frames.json, fallback to legacy fences.json if it's an old backup
                string backupFramesPath = Path.Combine(backupFolder, "frames.json");
                if (!File.Exists(backupFramesPath))
                {
                    backupFramesPath = Path.Combine(backupFolder, "fences.json");
                }
                string backupShortcutsPath = Path.Combine(backupFolder, "Shortcuts");

                if (!File.Exists(backupFramesPath) || !Directory.Exists(backupShortcutsPath))
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
                string currentFramesPath = ProfileManager.GetProfileFilePath("frames.json");
                string currentShortcutsPath = ProfileManager.GetProfileFilePath("Shortcuts");

                var FrameData = Framemanager.GetFrameData();
                FrameData?.Clear();

                // KISS SCRUBBER: Read backup, sanitize legacy keys, and write to new format
                string backupContent = File.ReadAllText(backupFramesPath);
                backupContent = backupContent.Replace("\"FenceBorderColor\"", "\"FrameBorderColor\"");
                backupContent = backupContent.Replace("\"frameBorderColor\"", "\"FrameBorderColor\"");
                backupContent = backupContent.Replace("\"FenceBorderThickness\"", "\"FrameBorderThickness\"");
                backupContent = backupContent.Replace("\"frameBorderThickness\"", "\"FrameBorderThickness\"");

                File.WriteAllText(currentFramesPath, backupContent);

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

                    // Spawns a hidden command prompt that waits ~2 seconds before launching the app.
                    // UseShellExecute = false explicitly prevents the child cmd.exe from inheriting the Single-Instance Mutex.
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = $"/c ping 127.0.0.1 -n 3 > nul & start \"\" \"{appPath}\"",
                        WindowStyle = ProcessWindowStyle.Hidden,
                        CreateNoWindow = true,
                        UseShellExecute = false
                    });

                    // OS-LEVEL TERMINATION: Bypasses .NET finalizers and WPF dispatchers entirely. 
                    // Guarantees the old instance is instantly destroyed and cannot deadlock.
                    Process.GetCurrentProcess().Kill();
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
            finally
            {
                Application.Current?.Dispatcher.Invoke(() => Mouse.OverrideCursor = null);
            }
        }
        #endregion
    }
}