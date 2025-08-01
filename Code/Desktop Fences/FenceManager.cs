﻿using IWshRuntimeLibrary;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Animation;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.IO;
using System.Windows.Shapes;
using System.Windows.Media.Effects;

using Microsoft.Win32;
using System.IO.Compression;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Net.NetworkInformation;



namespace Desktop_Fences
{
    public static class FenceManager
    {

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern uint ExtractIconEx(string szFileName, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

        [DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr hIcon);

        private static List<dynamic> _fenceData;
        private static string _jsonFilePath;
        private static dynamic _options;
        private static readonly Dictionary<string, ImageSource> iconCache = new Dictionary<string, ImageSource>();
        private static Dictionary<dynamic, PortalFenceManager> _portalFences = new Dictionary<dynamic, PortalFenceManager>();
        private static string _lastDeletedFolderPath;
        private static dynamic _lastDeletedFence;
        private static bool _isRestoreAvailable;
        // Stores heart TextBlock references for each fence to enable efficient ContextMenu updates
        private static readonly Dictionary<dynamic, TextBlock> _heartTextBlocks = new Dictionary<dynamic, TextBlock>();

        //resize feedback
        private static Window _sizeFeedbackWindow;
        private static System.Windows.Threading.DispatcherTimer _hideTimer;

        // Add near other static fields
        private static TargetChecker _currentTargetChecker;

        // Track fences currently in rollup/rolldown transition to prevent event conflicts
        private static readonly HashSet<string> _fencesInTransition = new HashSet<string>();
        // Emergency cleanup timer to prevent permanently stuck transition states
        private static System.Windows.Threading.DispatcherTimer _transitionCleanupTimer;


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

        // Add this new method
        public static void ReloadFences()
        {
            try
            {
                // Clear existing data
                _fenceData?.Clear();

                // Stop previous target checker
                _currentTargetChecker?.Stop();

                // Create new target checker
                _currentTargetChecker = new TargetChecker(1000);

                // Reload settings and fences
                LoadAndCreateFences(_currentTargetChecker);
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.Error, $"Error reloading fences: {ex.Message}");
                throw;
            }
        }


        public static void RestoreFromBackup(string backupFolder)
        {
            try
            {
                // Validate backup folder
                string backupFencesPath = System.IO.Path.Combine(backupFolder, "fences.json");
                string backupShortcutsPath = System.IO.Path.Combine(backupFolder, "Shortcuts");

                if (!System.IO.File.Exists(backupFencesPath) || !Directory.Exists(backupShortcutsPath))
                {
                    // MessageBox.Show("Invalid backup folder - missing required files", "Restore Error",

                    //              MessageBoxButton.OK, MessageBoxImage.Error);
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm("Invalid backup folder - missing required files", "Restore Error");
                    return;
                }

                // Clear existing fences
                _fenceData.Clear();
                _heartTextBlocks.Clear();
                _portalFences.Clear();

                // Restore fences.json
                string currentFencesPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "fences.json");
                System.IO.File.Copy(backupFencesPath, currentFencesPath, true);

                // Restore shortcuts
                string currentShortcutsPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Shortcuts");
                if (Directory.Exists(currentShortcutsPath))
                {
                    Directory.Delete(currentShortcutsPath, true);
                }
                Directory.CreateDirectory(currentShortcutsPath);
                CopyDirectory(backupShortcutsPath, currentShortcutsPath);

                // Reload fences
                LoadAndCreateFences(new TargetChecker(1000));
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.Error, $"Restore failed: {ex.Message}");
                //    MessageBox.Show($"Restore failed: {ex.Message}", "Error",
                //               MessageBoxButton.OK, MessageBoxImage.Error);

                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Restore failed: {ex.Message}", "Error");
            }
        }

        // Add the ExportFence method
        public static void ExportFence(dynamic fence)
        {
            try
            {
                string exeDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string fenceTitle = fence.Title.ToString();

                // Sanitize folder name
                foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                {
                    fenceTitle = fenceTitle.Replace(c, '_');
                }

                string exportFolder = System.IO.Path.Combine(exeDir, "Exports", fenceTitle);
                string fencePath = System.IO.Path.Combine(exeDir, "Exports", $"{fenceTitle}.fence");

                // Cleanup existing
                if (Directory.Exists(exportFolder))
                    Directory.Delete(exportFolder, true);
                if (System.IO.File.Exists(fencePath))
                    System.IO.File.Delete(fencePath);

                Directory.CreateDirectory(exportFolder);

                // Save fence data
                string fenceJson = JsonConvert.SerializeObject(fence, Formatting.Indented);
                System.IO.File.WriteAllText(System.IO.Path.Combine(exportFolder, "fence.json"), fenceJson);

                // Copy shortcuts for Data fences
                if (fence.ItemsType?.ToString() == "Data")
                {
                    string shortcutsDestDir = System.IO.Path.Combine(exportFolder, "Shortcuts");
                    Directory.CreateDirectory(shortcutsDestDir);

                    foreach (var item in fence.Items)
                    {
                        string filename = item.Filename?.ToString();
                        if (!string.IsNullOrEmpty(filename))
                        {
                            string sourcePath = System.IO.Path.Combine(exeDir, filename);
                            if (System.IO.File.Exists(sourcePath))
                            {
                                string destName = System.IO.Path.GetFileName(filename);
                                System.IO.File.Copy(sourcePath, System.IO.Path.Combine(shortcutsDestDir, destName));
                            }
                        }
                    }
                }

                // Create zip and cleanup
                ZipFile.CreateFromDirectory(exportFolder, fencePath);
                Directory.Delete(exportFolder, true);

                //   MessageBox.Show($"Fence exported to:\n{fencePath}", "Export Successful",
                //               MessageBoxButton.OK, MessageBoxImage.Information);

                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Fence exported to:\n{fencePath}", "Export Successful");
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.Error, $"Export failed: {ex.Message}");
                //   MessageBox.Show($"Export failed: {ex.Message}", "Error",
                //                  MessageBoxButton.OK, MessageBoxImage.Error);
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Export failed: {ex.Message}", "Error");
            }
        }


        private static double GetDpiScaleFactor(Window window)
        {
            // Get the screen where the window is located based on its position
            var screen = System.Windows.Forms.Screen.FromPoint(
                new System.Drawing.Point((int)window.Left, (int)window.Top));

            // Use Graphics to get the screen's DPI
            using (var graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiX = graphics.DpiX; // Horizontal DPI
                return dpiX / 96.0; // Standard DPI is 96, so scale factor = dpiX / 96
            }
        }
        private static void AdjustFencePositionToScreen(NonActivatingWindow win)
        {
            // Get the DPI scale factor for the window
            double dpiScale = GetDpiScaleFactor(win);

            // Determine the screen based on the current position in pixels
            var screen = System.Windows.Forms.Screen.FromPoint(
                new System.Drawing.Point((int)(win.Left * dpiScale), (int)(win.Top * dpiScale)));

            // Get the screen's working area in pixels (excludes taskbars)
            var workingArea = screen.WorkingArea;

            // Convert window position and size from DIUs to pixels
            double winLeftPx = win.Left * dpiScale;
            double winTopPx = win.Top * dpiScale;
            double winWidthPx = win.Width * dpiScale;
            double winHeightPx = win.Height * dpiScale;

            // Calculate new position in pixels
            double newLeftPx = winLeftPx;
            double newTopPx = winTopPx;

            // Ensure the right edge doesn't exceed the working area's right boundary
            if (newLeftPx + winWidthPx > workingArea.Right)
            {
                newLeftPx = workingArea.Right - winWidthPx;
            }
            // Ensure the left edge isn't off-screen to the left
            if (newLeftPx < workingArea.Left)
            {
                newLeftPx = workingArea.Left;
            }

            // Ensure the bottom edge doesn't exceed the working area's bottom boundary
            if (newTopPx + winHeightPx > workingArea.Bottom)
            {
                newTopPx = workingArea.Bottom - winHeightPx;
            }
            // Ensure the top edge isn't off-screen to the top
            if (newTopPx < workingArea.Top)
            {
                newTopPx = workingArea.Top;
            }

            // Convert the adjusted position back to DIUs
            double newLeft = newLeftPx / dpiScale;
            double newTop = newTopPx / dpiScale;

            // Apply the new position if it has changed
            if (newLeft != win.Left || newTop != win.Top)
            {
                win.Left = newLeft;
                win.Top = newTop;
                Log(LogLevel.Debug, LogCategory.ImportExport, $"Adjusted fence '{win.Title}' position to ({newLeft}, {newTop}) to fit within screen bounds.");
            }
            SaveFenceData();
        }

        // Import function implementation
        public static void ImportFence()
        {
            try
            {
                string exeDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string exportsDir = System.IO.Path.Combine(exeDir, "Exports");

                var openDialog = new OpenFileDialog
                {
                    Filter = "Fence Files|*.fence", // Changed filter
                    DefaultExt = ".fence",          // Added default extension
                    InitialDirectory = Directory.Exists(exportsDir) ? exportsDir : exeDir,
                    Title = "Select Fence Export File"
                };

                if (openDialog.ShowDialog() != true) return;

                string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);

                try
                {
                    // Extract ZIP contents
                    ZipFile.ExtractToDirectory(openDialog.FileName, tempDir);

                    // Validate export structure
                    string fenceJsonPath = System.IO.Path.Combine(tempDir, "fence.json");
                    if (!System.IO.File.Exists(fenceJsonPath))
                    {
                        throw new FileNotFoundException("Missing fence.json in export file");
                    }

                    // Deserialize fence data
                    string jsonContent = System.IO.File.ReadAllText(fenceJsonPath);
                    dynamic importedFence = JsonConvert.DeserializeObject<JObject>(jsonContent);

                    // Generate new ID to prevent conflicts
                    importedFence["Id"] = Guid.NewGuid().ToString();

                    // Handle shortcuts for Data fences
                    if (importedFence.ItemsType?.ToString() == "Data")
                    {
                        string sourceShortcuts = System.IO.Path.Combine(tempDir, "Shortcuts");
                        string destShortcuts = System.IO.Path.Combine(exeDir, "Shortcuts");

                        if (Directory.Exists(sourceShortcuts))
                        {
                            Directory.CreateDirectory(destShortcuts);
                            foreach (string srcPath in Directory.GetFiles(sourceShortcuts))
                            {
                                string fileName = System.IO.Path.GetFileName(srcPath);
                                string destPath = System.IO.Path.Combine(destShortcuts, fileName);

                                // Handle duplicate filenames
                                int counter = 1;
                                while (System.IO.File.Exists(destPath))
                                {
                                    string tempName = $"{System.IO.Path.GetFileNameWithoutExtension(fileName)} ({counter++}){System.IO.Path.GetExtension(fileName)}";
                                    destPath = System.IO.Path.Combine(destShortcuts, tempName);
                                }

                                System.IO.File.Copy(srcPath, destPath);

                                // Update shortcut references in fence items
                                var items = importedFence.Items as JArray;
                                foreach (var item in items.Where(i => i["Filename"]?.ToString() == fileName))
                                {
                                    item["Filename"] = System.IO.Path.Combine("Shortcuts", System.IO.Path.GetFileName(destPath));
                                }
                            }
                        }
                    }


                    _fenceData.Add(importedFence);
                    CreateFence(importedFence, new TargetChecker(1000));
                    SaveFenceData();

                    //MessageBox.Show("Fence imported successfully!", "Import Complete",
                    //              MessageBoxButton.OK, MessageBoxImage.Information);


                    TrayManager.Instance.ShowOKOnlyMessageBoxForm("Fence imported successfully!", "Import Complete");
                }
                finally
                {
                    // Cleanup temporary files
                    if (Directory.Exists(tempDir))
                    {
                        Directory.Delete(tempDir, true);
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.Error, $"Fence import failed: {ex.Message}");
                // MessageBox.Show($"Failed to import fence: {ex.Message}", "Import Error",
                //              MessageBoxButton.OK, MessageBoxImage.Error);

                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to import fence: {ex.Message}", "Import Error");
            }
        }



        public static void UpdateHeartContextMenus()
        {
            UpdateAllHeartContextMenus();
        }


        // Builds the heart ContextMenu for a fence with consistent items and dynamic state
        private static ContextMenu BuildHeartContextMenu(dynamic fence)
        {
            var menu = new ContextMenu();

            // About item
            var aboutItem = new MenuItem { Header = "About..." };
            aboutItem.Click += (s, e) => TrayManager.Instance.ShowAboutForm();
            menu.Items.Add(aboutItem);

            // Options item
            var optionsItem = new MenuItem { Header = "Options..." };
            optionsItem.Click += (s, e) => TrayManager.Instance.ShowOptionsForm();
            menu.Items.Add(optionsItem);

            // Separator
            menu.Items.Add(new Separator());

            // New Fence item
            var newFenceItem = new MenuItem { Header = "New Fence" };
            newFenceItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                Log(LogLevel.Info, LogCategory.FenceUpdate, $"Creating new fence at mouse position: X={mousePosition.X}, Y={mousePosition.Y}");
                //CreateNewFence("New Fence", "Data", mousePosition.X, mousePosition.Y);
                CreateNewFence("", "Data", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newFenceItem);

            // New Portal Fence item
            var newPortalFenceItem = new MenuItem { Header = "New Portal Fence" };
            newPortalFenceItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                Log(LogLevel.Info, LogCategory.FenceUpdate, $"Creating new portal fence at mouse position: X={mousePosition.X}, Y={mousePosition.Y}");
                CreateNewFence("New Portal Fence", "Portal", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newPortalFenceItem);

            menu.Items.Add(new Separator());


            // Delete this fence
            var deleteThisFence = new MenuItem { Header = "Delete this Fence" };
            deleteThisFence.Click += (s, e) =>
            {
                bool result = TrayManager.Instance.ShowCustomMessageBoxForm(); // Call the method and store the result  

                if (result == true)
                {
                    // Ensure the backup folder exists
                    _lastDeletedFolderPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Last Fence Deleted");
                    if (!Directory.Exists(_lastDeletedFolderPath))
                    {
                        Directory.CreateDirectory(_lastDeletedFolderPath);
                    }

                    // Clear previous backup
                    foreach (var file in Directory.GetFiles(_lastDeletedFolderPath))
                    {
                        System.IO.File.Delete(file);
                    }

                    // Backup the fence and its shortcuts
                    _lastDeletedFence = fence;
                    _isRestoreAvailable = true;
                    UpdateAllHeartContextMenus(); // Update all heart menus to show restore option

                    // Backup shortcuts for Data fences
                    if (fence.ItemsType?.ToString() == "Data")
                    {
                        var items = fence.Items as JArray;
                        if (items != null)
                        {
                            foreach (var item in items)
                            {
                                string itemFilePath = item["Filename"]?.ToString();
                                if (!string.IsNullOrEmpty(itemFilePath) && System.IO.File.Exists(itemFilePath))
                                {
                                    string shortcutPath = System.IO.Path.Combine(_lastDeletedFolderPath, System.IO.Path.GetFileName(itemFilePath));
                                    System.IO.File.Copy(itemFilePath, shortcutPath, true);
                                }
                                else
                                {
                                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Skipped backing up missing file: {itemFilePath}");
                                }
                            }
                        }
                    }

                    // Save fence info to JSON
                    string fenceJsonPath = System.IO.Path.Combine(_lastDeletedFolderPath, "fence.json");
                    System.IO.File.WriteAllText(fenceJsonPath, JsonConvert.SerializeObject(fence, Formatting.Indented));

                    // Proceed with deletion
                    _fenceData.Remove(fence);
                    _heartTextBlocks.Remove(fence); // Remove from heart TextBlocks dictionary
                    if (_portalFences.ContainsKey(fence))
                    {
                        _portalFences[fence].Dispose();
                        _portalFences.Remove(fence);
                    }
                    SaveFenceData();

                    // Find and close the window - need to find the window associated with this fence
                    var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                    var win = windows.FirstOrDefault(w => w.Tag?.ToString() == fence.Id?.ToString());
                    if (win != null)
                    {
                        win.Close();
                    }

                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Fence {fence.Title} removed successfully from heart context menu");
                }
            };
            menu.Items.Add(deleteThisFence);



            // Restore Fence item
            var restoreItem = new MenuItem
            {
                Header = "Restore Last Deleted Fence",
                Visibility = _isRestoreAvailable ? Visibility.Visible : Visibility.Collapsed
            };
            restoreItem.Click += (s, e) => RestoreLastDeletedFence();
            menu.Items.Add(restoreItem);

            menu.Items.Add(new Separator());
            // New Export Fence menu item
            var exportItem = new MenuItem { Header = "Export this Fence" };
            exportItem.Click += (s, e) => ExportFence(fence);
            menu.Items.Add(exportItem);

            var importItem = new MenuItem { Header = "Import a Fence..." };
            importItem.Click += (s, e) => ImportFence();
            menu.Items.Add(importItem);


            // Separator
            menu.Items.Add(new Separator());

            // Exit item
            var exitItem = new MenuItem { Header = "Exit" };
            exitItem.Click += (s, e) => Application.Current.Shutdown();
            menu.Items.Add(exitItem);

            return menu;
        }


        // Updates all heart ContextMenus across all fences using stored TextBlock references
        public static void UpdateAllHeartContextMenus()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                foreach (var entry in _heartTextBlocks)
                {
                    var fence = entry.Key;
                    var heart = entry.Value;
                    if (heart != null)
                    {
                        heart.ContextMenu = BuildHeartContextMenu(fence);
                        Log(LogLevel.Info, LogCategory.FenceUpdate, $"Updated heart ContextMenu for fence '{fence.Title}'");
                    }
                    else
                    {
                        Log(LogLevel.Info, LogCategory.FenceUpdate, $"Skipped update for fence '{fence.Title}': heart TextBlock is null");
                    }
                }
            });
        }



        // Helper methods for web link detection and handling
        private static bool IsWebLinkShortcut(string filePath)
        {
            try
            {
                string extension = System.IO.Path.GetExtension(filePath).ToLower();

                if (extension == ".url")
                {
                    // Check if .url file contains web URL
                    if (System.IO.File.Exists(filePath))
                    {
                        string content = System.IO.File.ReadAllText(filePath);
                        return content.Contains("URL=http://") || content.Contains("URL=https://");
                    }
                }
                else if (extension == ".lnk")
                {
                    // Check if .lnk file targets web URL
                    if (System.IO.File.Exists(filePath))
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                        string target = shortcut.TargetPath ?? "";
                        return target.StartsWith("http://") || target.StartsWith("https://");
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.IconHandling, $"Error checking if {filePath} is web link: {ex.Message}");
            }
            return false;
        }

        private static string ExtractWebUrlFromFile(string filePath)
        {
            try
            {
                string extension = System.IO.Path.GetExtension(filePath).ToLower();

                if (extension == ".url")
                {
                    // Extract URL from .url file
                    if (System.IO.File.Exists(filePath))
                    {
                        string[] lines = System.IO.File.ReadAllLines(filePath);
                        foreach (string line in lines)
                        {
                            if (line.StartsWith("URL="))
                            {
                                return line.Substring(4); // Remove "URL=" prefix
                            }
                        }
                    }
                }
                else if (extension == ".lnk")
                {
                    // Extract URL from .lnk file
                    if (System.IO.File.Exists(filePath))
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                        return shortcut.TargetPath ?? "";
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.IconHandling, $"Error extracting web URL from {filePath}: {ex.Message}");
            }
            return null;
        }

        private static void CreateWebLinkShortcut(string targetUrl, string shortcutPath, string displayName)
        {
            try
            {
                // For web links, create a .url file instead of .lnk file
                // Change the extension to .url if it's not already
                string urlFilePath = shortcutPath;
                if (System.IO.Path.GetExtension(shortcutPath).ToLower() == ".lnk")
                {
                    urlFilePath = System.IO.Path.ChangeExtension(shortcutPath, ".url");
                }

                // Create a .url file content
                string urlFileContent = $"[InternetShortcut]\r\nURL={targetUrl}\r\nIconIndex=0\r\n";

                // Write the .url file
                System.IO.File.WriteAllText(urlFilePath, urlFileContent, System.Text.Encoding.ASCII);

                Log(LogLevel.Debug, LogCategory.IconHandling, $"Created web link URL file: {urlFilePath} -> {targetUrl}");
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, LogCategory.IconHandling, $"Error creating web link shortcut {shortcutPath}: {ex.Message}");
                throw;
            }
        }


        // Helper method for network path detection
        private static bool IsNetworkPath(string filePath)
        {
            try
            {
                bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";

                if (isShortcut)
                {
                    // For shortcuts, check the target path
                    if (System.IO.File.Exists(filePath))
                    {
                        string targetPath = Utility.GetShortcutTarget(filePath);
                        if (!string.IsNullOrEmpty(targetPath))
                        {
                            // Check if target is UNC path
                            bool isUncPath = targetPath.StartsWith("\\\\");
                            Log(LogLevel.Debug, LogCategory.IconHandling, $"Shortcut {filePath} targets {targetPath}, IsUNC: {isUncPath}");
                            return isUncPath;
                        }
                    }
                }
                else
                {
                    // For direct paths, check if it's UNC
                    bool isUncPath = filePath.StartsWith("\\\\");
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"Direct path {filePath}, IsUNC: {isUncPath}");
                    return isUncPath;
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.IconHandling, $"Error checking if {filePath} is network path: {ex.Message}");
            }
            return false;
        }



        public enum LaunchEffect
        {
            Zoom,        // First effect
            Bounce,      // Scale up and down a few times
            FadeOut,     // Fade out
            SlideUp,     // Slide up and return
            Rotate,      // Spin 360 degrees
            Agitate,     // Shake back and forth
            GrowAndFly,  // Grow and fly away
            Pulse,       // Pulisng
            Elastic,     // Elastic *****
            Flip3D,     //Flip 3D
            Spiral      // Spiral

        }

        public static List<dynamic> GetFenceData()
        {
            return _fenceData;
        }

        /// <summary>
        /// Logs a message with the specified level and category.
        /// 
        /// LogLevel:
        /// Debug: Verbose logs (e.g., UI updates, icon caching, position changes).
        /// Info: General operations (e.g., fence creation, property updates).
        /// Warn: Non-critical issues (e.g., missing files, invalid settings).
        /// Error: Critical failures (e.g., exceptions).
        /// 
        /// LogCategory:
        /// General: Miscellaneous or uncategorized logs.
        /// FenceCreation: Logs related to creating or initializing fences.
        /// FenceUpdate: Logs for updating fence properties or content.
        /// UI: Logs for UI interactions (e.g., mouse events, context menus).
        /// IconHandling: Logs for icon loading, updating, or removal.
        /// Error: Logs for errors or exceptions.
        /// ImportExport: Logs for importing/exporting fences.
        /// Settings: Logs for settings changes.
        /// </summary>
        /// <param name="level"></param>
        /// <param name="category"></param>
        /// <param name="message"></param>


        public static void Log(LogLevel level, LogCategory category, string message)
        {
            // Check if logging is enabled and if the level/category is allowed
            if (!(_options?.IsLogEnabled ?? true))
                return;

            LogLevel minLevel = SettingsManager.MinLogLevel; // New setting
            List<LogCategory> enabledCategories = SettingsManager.EnabledLogCategories; // New setting

            if (level < minLevel || !enabledCategories.Contains(category))
                return;

            string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");

            // Implement log rotation
            const long maxFileSize = 5 * 1024 * 1024; // 5MB
            if (System.IO.File.Exists(logPath))
            {
                FileInfo fileInfo = new FileInfo(logPath);
                if (fileInfo.Length > maxFileSize)
                {
                    string archivePath = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(logPath),
                        $"Desktop_Fences_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                    System.IO.File.Move(logPath, archivePath);

                    // Clean up old logs (keep last 5)
                    var logFiles = Directory.GetFiles(System.IO.Path.GetDirectoryName(logPath), "Desktop_Fences_*.log")
                        .OrderByDescending(f => System.IO.File.GetCreationTime(f))
                        .Skip(5);
                    foreach (var oldLog in logFiles)
                    {
                        try
                        {
                            System.IO.File.Delete(oldLog);
                        }
                        catch (Exception ex)
                        {
                            // Log deletion failure to a minimal fallback log
                            System.IO.File.AppendAllText(
                                System.IO.Path.Combine(System.IO.Path.GetDirectoryName(logPath), "Desktop_Fences_Fallback.log"),
                                $"{DateTime.Now}: Failed to delete old log {oldLog}: {ex.Message}\n");
                        }
                    }
                }
            }

            // Write log with level and category
            string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}][{category}] {message}\n";
            System.IO.File.AppendAllText(logPath, logMessage);
        }





        //// Advanced debugging system for fence resize issues // keep for future use and alternative debugging
        //private static void DebugLog(string category, string fenceId, string message, double? height = null, double? unrolledHeight = null, bool? isRolled = null)
        //{
        //    try
        //    {
        //        string debugPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "debug.log");
        //        string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        //        string debugMessage = $"[{timestamp}] [{category}] FenceID:{fenceId} | {message}";

        //        if (height.HasValue) debugMessage += $" | Height:{height:F1}";
        //        if (unrolledHeight.HasValue) debugMessage += $" | UnrolledHeight:{unrolledHeight:F1}";
        //        if (isRolled.HasValue) debugMessage += $" | IsRolled:{isRolled}";

        //        debugMessage += "\n";
        //        System.IO.File.AppendAllText(debugPath, debugMessage);
        //    }
        //    catch { } // Silent fail to avoid breaking main functionality
        //}

        //private static void DebugLogFenceState(string context, string fenceId, NonActivatingWindow win = null)
        //{
        //    try
        //    {
        //        var fence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
        //        if (fence != null)
        //        {
        //            double winHeight = win?.Height ?? -1;
        //            double fenceHeight = Convert.ToDouble(fence.Height?.ToString() ?? "0");
        //            double unrolledHeight = Convert.ToDouble(fence.UnrolledHeight?.ToString() ?? "0");
        //            bool isRolled = fence.IsRolled?.ToString().ToLower() == "true";

        //           // DebugLog("STATE", fenceId, $"{context} | WinHeight:{winHeight:F1} | FenceHeight:{fenceHeight:F1}",
        //            //    winHeight, unrolledHeight, isRolled);
        //        }
        //    }
        //    catch { }
        //}



        // Update fence property, save to JSON, and apply runtime changes

        public static void UpdateFenceProperty(dynamic fence, string propertyName, string value, string logMessage)
        {
            try
            {


                string fenceId = fence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    Log(LogLevel.Info, LogCategory.FenceUpdate, $"Fence '{fence.Title}' has no Id");
                    return;
                }

                // Skip updates if fence is in transition (except for IsRolled and UnrolledHeight which are rollup-specific)
                if (_fencesInTransition.Contains(fenceId) && propertyName != "IsRolled" && propertyName != "UnrolledHeight")
                {
                    Log(LogLevel.Debug, LogCategory.FenceUpdate, $"Skipping {propertyName} update for fence '{fenceId}' - in transition");
                    return;
                }

                // Find by GUID instead of title or reference
                int index = _fenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (index >= 0)
                {
                    // Get the fence from _fenceData
                    dynamic actualFence = _fenceData[index];





                    // Convert to dictionary safely
                    IDictionary<string, object> fenceDict = actualFence as IDictionary<string, object> ?? ((JObject)actualFence).ToObject<IDictionary<string, object>>();

                    // Handle IsHidden specifically to store as string to match JSON format
                    if (propertyName == "IsHidden")
                    {
                        // Convert boolean-like string input to string "true" or "false"
                        bool parsedValue = value?.ToLower() == "true";
                        fenceDict[propertyName] = parsedValue.ToString().ToLower(); // Store as "true" or "false"
                    }
                    else
                    {
                        // Update other properties as provided
                        fenceDict[propertyName] = value;
                    }

                    // Update the fence in _fenceData
                    _fenceData[index] = JObject.FromObject(fenceDict);
                    SaveFenceData();
                    Log(LogLevel.Info, LogCategory.FenceUpdate, $"{logMessage} for fence '{fence.Title}'");

                    // Find the window to apply runtime changes
                    var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                    //var win = windows.FirstOrDefault(w => w.Title == fence.Title.ToString());
                    var win = windows.FirstOrDefault(w => w.Tag?.ToString() == fenceId);
                    if (win != null)
                    {
                        // Apply runtime changes
                        if (propertyName == "CustomColor")
                        {
                            Utility.ApplyTintAndColorToFence(win, string.IsNullOrEmpty(value) ? _options.SelectedColor : value);
                            Log(LogLevel.Info, LogCategory.FenceUpdate, $"Applied color '{value ?? "Default"}' to fence '{fence.Title}' at runtime");
                        }
                        else if (propertyName == "IsHidden")
                        {
                            // Update visibility based on IsHidden
                            bool isHidden = value?.ToLower() == "true";
                            win.Visibility = isHidden ? Visibility.Hidden : Visibility.Visible;
                            Log(LogLevel.Info, LogCategory.FenceUpdate, $"Set visibility to {(isHidden ? "Hidden" : "Visible")} for fence '{fence.Title}'");
                        }
                        //step 7
                        else if (propertyName == "IsRolled")
                        {
                            bool isRolled = value?.ToLower() == "true";
                            double targetHeight = isRolled ? 26 : Convert.ToDouble(actualFence.UnrolledHeight?.ToString() ?? "130");
                            var heightAnimation = new DoubleAnimation(win.Height, targetHeight, TimeSpan.FromSeconds(0.3))
                            {
                                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                            };
                            win.BeginAnimation(Window.HeightProperty, heightAnimation);
                            // Update WrapPanel visibility
                            var border = win.Content as Border;
                            if (border != null)
                            {
                                var dockPanel = border.Child as DockPanel;
                                if (dockPanel != null)
                                {
                                    var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                                    if (scrollViewer != null)
                                    {
                                        var wpcont = scrollViewer.Content as WrapPanel;
                                        if (wpcont != null)
                                        {
                                            wpcont.Visibility = isRolled ? Visibility.Collapsed : Visibility.Visible;
                                            Log(LogLevel.Info, LogCategory.FenceUpdate, $"Set WrapPanel visibility to {(isRolled ? "Collapsed" : "Visible")} for fence '{actualFence.Title}'");
                                        }
                                    }
                                }
                            }
                            Log(LogLevel.Info, LogCategory.FenceUpdate, $"Set height to {targetHeight} for fence '{actualFence.Title}' (IsRolled={isRolled})");
                        }
                        else if (propertyName == "UnrolledHeight")
                        {
                            Log(LogLevel.Info, LogCategory.FenceUpdate, $"Set UnrolledHeight to {value} for fence '{actualFence.Title}'");
                        }







                        // Update context menu checkmarks
                        if (win.ContextMenu != null)
                        {
                            var customizeMenu = win.ContextMenu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Header.ToString() == "Customize");
                            if (customizeMenu != null)
                            {
                                var submenu = propertyName == "CustomColor"
                                    ? customizeMenu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Header.ToString() == "Color")
                                    : customizeMenu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Header.ToString() == "Launch Effect");
                                if (submenu != null)
                                {
                                    foreach (MenuItem item in submenu.Items)
                                    {
                                        item.IsChecked = item.Tag?.ToString() == value || (value == null && item.Header.ToString() == "Default");
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Log(LogLevel.Warn, LogCategory.FenceUpdate, $"Failed to find window for fence '{fence.Title}' to apply {propertyName}");
                    }
                }
                else
                {
                    Log(LogLevel.Warn, LogCategory.FenceUpdate, $"Failed to find fence '{fence.Title}' in _fenceData for {propertyName} update");
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, LogCategory.FenceUpdate, $"Error updating {propertyName} for fence '{fence.Title}': {ex.Message}");
            }
        }

        public static void RestoreLastDeletedFence()
        {
            if (!_isRestoreAvailable || string.IsNullOrEmpty(_lastDeletedFolderPath) || _lastDeletedFence == null)
            {
                // MessageBox.Show("No fence to restore.", "Restore", MessageBoxButton.OK, MessageBoxImage.Information);
                TrayManager.Instance.ShowOKOnlyMessageBoxForm("No fence to restore", "Restore");
                return;
            }

            // Restore shortcuts
            var shortcutFiles = Directory.GetFiles(_lastDeletedFolderPath, "*.lnk");
            foreach (var shortcutFile in shortcutFiles)
            {
                string destinationPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Shortcuts", System.IO.Path.GetFileName(shortcutFile));
                System.IO.File.Copy(shortcutFile, destinationPath, true);
            }

            // Restore fence data
            _fenceData.Add(_lastDeletedFence);
            SaveFenceData();
            CreateFence(_lastDeletedFence, new TargetChecker(1000));

            // Clear backup state
            _lastDeletedFence = null;
            _isRestoreAvailable = false;

            UpdateAllHeartContextMenus();


            Log(LogLevel.Info, LogCategory.ImportExport, "Fence restored successfully.");
        }

        public static void CleanLastDeletedFolder()
        {
            _lastDeletedFolderPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Last Fence Deleted");
            if (Directory.Exists(_lastDeletedFolderPath))
            {
                foreach (var file in Directory.GetFiles(_lastDeletedFolderPath))
                {
                    System.IO.File.Delete(file);
                }
            }
            else
            {
                Directory.CreateDirectory(_lastDeletedFolderPath);
            }
            _isRestoreAvailable = false;
            _lastDeletedFence = null;


            UpdateAllHeartContextMenus();
        }
        public static void LoadAndCreateFences(TargetChecker targetChecker)
        {
            string exePath = Assembly.GetEntryAssembly().Location;
            string exeDir = System.IO.Path.GetDirectoryName(exePath);
            _jsonFilePath = System.IO.Path.Combine(exeDir, "fences.json");
            // Below added for reload function
            _currentTargetChecker = targetChecker;


            SettingsManager.LoadSettings();
            CleanLastDeletedFolder();
            _options = new
            {
                IsSnapEnabled = SettingsManager.IsSnapEnabled,
                ShowBackgroundImageOnPortalFences = SettingsManager.ShowBackgroundImageOnPortalFences,
                Showintray = SettingsManager.ShowInTray,
                TintValue = SettingsManager.TintValue,
                SelectedColor = SettingsManager.SelectedColor,
                IsLogEnabled = SettingsManager.IsLogEnabled,
                singleClickToLaunch = SettingsManager.SingleClickToLaunch,
                LaunchEffect = SettingsManager.LaunchEffect,
                CheckNetworkPaths = false // TODO: Add to SettingsManager later when UI is ready
            };

            if (System.IO.File.Exists(_jsonFilePath))
            {
                string jsonContent = System.IO.File.ReadAllText(_jsonFilePath);
                try
                {
                    _fenceData = JsonConvert.DeserializeObject<List<dynamic>>(jsonContent);
                }
                catch (JsonSerializationException)
                {
                    var singleFence = JsonConvert.DeserializeObject<dynamic>(jsonContent);
                    _fenceData = new List<dynamic> { singleFence };
                }

                if (_fenceData == null || _fenceData.Count == 0)
                {
                    InitializeDefaultFence();
                }
                MigrateLegacyJson();
            }
            else
            {
                InitializeDefaultFence();
            }
            // Sanitize Portal Fences with missing target folders
            var invalidFences = new List<dynamic>();
            foreach (dynamic fence in _fenceData.ToList()) // Use ToList to avoid collection modification issues
            {
                if (fence.ItemsType?.ToString() == "Portal")
                {
                    string targetPath = fence.Path?.ToString();
                    if (string.IsNullOrEmpty(targetPath) || !System.IO.Directory.Exists(targetPath))
                    {
                        invalidFences.Add(fence);
                        Log(LogLevel.Warn, LogCategory.FenceCreation, $"Marked Portal Fence '{fence.Title}' for removal due to missing target folder: {targetPath ?? "null"}");
                    }
                }
            }

            // Remove invalid fences and save
            if (invalidFences.Any())
            {
                foreach (var fence in invalidFences)
                {
                    _fenceData.Remove(fence);
                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Removed Portal Fence '{fence.Title}' from _fenceData");
                }
                SaveFenceData();
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Saved updated fences.json after removing {invalidFences.Count} invalid Portal Fences");
            }

            // Clear any stuck transition states from previous session
            ClearAllTransitionStates();
            // Start emergency cleanup timer if not already running
            if (_transitionCleanupTimer == null)
            {
                _transitionCleanupTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(10) // Check every 10 seconds
                };
                _transitionCleanupTimer.Tick += (s, e) =>
                {
                    if (_fencesInTransition.Count > 0)
                    {
                        Log(LogLevel.Warn, LogCategory.FenceUpdate, $"Emergency cleanup: Found {_fencesInTransition.Count} fences stuck in transition state");
                        ClearAllTransitionStates();
                    }
                };
                _transitionCleanupTimer.Start();
            }


            foreach (dynamic fence in _fenceData.ToList())
            {
                CreateFence(fence, targetChecker);
            }
        }


        private static void InitializeDefaultFence()
        {
            string defaultJson = "[{\"Id\":\"" + Guid.NewGuid().ToString() + "\",\"Title\":\"New Fence\",\"X\":20,\"Y\":20,\"Width\":230,\"Height\":130,\"ItemsType\":\"Data\",\"Items\":[],\"IsLocked\":\"false\",\"IsHidden\":\"false\",\"CustomColor\":null,\"CustomLaunchEffect\":null,\"IsRolled\":\"false\",\"UnrolledHeight\":130}]";
            System.IO.File.WriteAllText(_jsonFilePath, defaultJson);
            _fenceData = JsonConvert.DeserializeObject<List<dynamic>>(defaultJson);
        }


        private static void MigrateLegacyJson()
        {
            try
            {
                bool jsonModified = false;
                var validColors = new HashSet<string> { "Red", "Green", "Teal", "Blue", "Bismark", "White", "Beige", "Gray", "Black", "Purple", "Fuchsia", "Yellow", "Orange" };
                var validEffects = Enum.GetNames(typeof(LaunchEffect)).ToHashSet();

                for (int i = 0; i < _fenceData.Count; i++)
                {
                    dynamic fence = _fenceData[i];
                    IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();

                    // Add GUID if missing
                    if (!fenceDict.ContainsKey("Id"))
                    {
                        fenceDict["Id"] = Guid.NewGuid().ToString();
                        jsonModified = true;
                        Log(LogLevel.Debug, LogCategory.General, $"Added Id to {fence.Title}");
                    }

                    // Existing migration: Handle Portal Fence ItemsType
                    if (fence.ItemsType?.ToString() == "Portal")
                    {
                        string portalPath = fence.Items?.ToString();
                        if (!string.IsNullOrEmpty(portalPath) && !System.IO.Directory.Exists(portalPath))
                        {
                            fenceDict["IsFolder"] = true;
                            jsonModified = true;
                            Log(LogLevel.Debug, LogCategory.General, $"Migrated legacy portal fence {fence.Title}: Added IsFolder=true");
                        }
                    }
                    else
                    {
                        var items = fence.Items as JArray ?? new JArray();
                        foreach (var item in items)
                        {
                            if (item["IsFolder"] == null)
                            {
                                string path = item["Filename"]?.ToString();
                                item["IsFolder"] = System.IO.Directory.Exists(path);
                                jsonModified = true;
                                Log(LogLevel.Debug, LogCategory.General, $"Migrated item in {fence.Title}: Added IsFolder for {path}");
                            }
                        }
                        fenceDict["Items"] = items; // Ensure Items updates are captured
                    }
                    // Migration: Add or validate IsLocked
                    if (!fenceDict.ContainsKey("IsLocked"))
                    {
                        fenceDict["IsLocked"] = "false"; // Default to string "false"
                        jsonModified = true;
                        Log(LogLevel.Debug, LogCategory.General, $"Added IsLocked=\"false\" to {fence.Title}");
                    }



                    // Migration: Add or validate IsRolled
                    if (!fenceDict.ContainsKey("IsRolled"))
                    {
                        fenceDict["IsRolled"] = "false";
                        jsonModified = true;
                        Log(LogLevel.Debug, LogCategory.General, $"Added IsRolled=\"false\" to {fence.Title}");
                    }
                    else
                    {
                        bool isRolled = false;
                        if (fenceDict["IsRolled"] is bool boolValue)
                        {
                            isRolled = boolValue;
                        }
                        else if (fenceDict["IsRolled"] is string stringValue)
                        {
                            isRolled = stringValue.ToLower() == "true";
                        }
                        else
                        {
                            Log(LogLevel.Debug, LogCategory.General, $"Invalid IsRolled value '{fenceDict["IsRolled"]}' in {fence.Title}, resetting to \"false\"");
                            isRolled = false;
                            jsonModified = true;
                        }
                        fenceDict["IsRolled"] = isRolled.ToString().ToLower();
                        Log(LogLevel.Debug, LogCategory.General, $"Preserved IsRolled: \"{isRolled.ToString().ToLower()}\" for {fence.Title}");
                    }

                    // Migration: Add or validate UnrolledHeight
                    if (!fenceDict.ContainsKey("UnrolledHeight"))
                    {
                        double height = fenceDict.ContainsKey("Height") ? Convert.ToDouble(fenceDict["Height"]) : 130;
                        fenceDict["UnrolledHeight"] = height;
                        jsonModified = true;
                        Log(LogLevel.Debug, LogCategory.General, $"Added UnrolledHeight={height} to {fence.Title}");
                    }
                    else
                    {
                        double unrolledHeight;
                        if (!double.TryParse(fenceDict["UnrolledHeight"]?.ToString(), out unrolledHeight) || unrolledHeight <= 0)
                        {
                            unrolledHeight = fenceDict.ContainsKey("Height") ? Convert.ToDouble(fenceDict["Height"]) : 130;
                            fenceDict["UnrolledHeight"] = unrolledHeight;
                            jsonModified = true;
                            Log(LogLevel.Debug, LogCategory.General, $"Invalid UnrolledHeight '{fenceDict["UnrolledHeight"]}' in {fence.Title}, reset to {unrolledHeight}");
                        }
                    }





                    // Migration: Add or validate CustomColor
                    if (!fenceDict.ContainsKey("CustomColor"))
                    {
                        fenceDict["CustomColor"] = null;
                        jsonModified = true;
                        Log(LogLevel.Debug, LogCategory.General, $"Added CustomColor=null to {fence.Title}");
                    }
                    else
                    {
                        string customColor = fenceDict["CustomColor"]?.ToString();
                        if (!string.IsNullOrEmpty(customColor) && !validColors.Contains(customColor))
                        {
                            Log(LogLevel.Debug, LogCategory.General, $"Invalid CustomColor '{customColor}' in {fence.Title}, resetting to null");
                            fenceDict["CustomColor"] = null;
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate CustomLaunchEffect
                    if (!fenceDict.ContainsKey("CustomLaunchEffect"))
                    {
                        fenceDict["CustomLaunchEffect"] = null;
                        jsonModified = true;
                        Log(LogLevel.Debug, LogCategory.General, $"Added CustomLaunchEffect=null to {fence.Title}");
                    }
                    else
                    {
                        string customEffect = fenceDict["CustomLaunchEffect"]?.ToString();
                        if (!string.IsNullOrEmpty(customEffect) && !validEffects.Contains(customEffect))
                        {
                            Log(LogLevel.Debug, LogCategory.General, $"Invalid CustomLaunchEffect '{customEffect}' in {fence.Title}, resetting to null");
                            fenceDict["CustomLaunchEffect"] = null;
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate IsHidden
                    if (!fenceDict.ContainsKey("IsHidden"))
                    {
                        fenceDict["IsHidden"] = "false"; // Default to string "false"
                        jsonModified = true;
                        Log(LogLevel.Debug, LogCategory.General, $"Added IsHidden=\"false\" to {fence.Title}");
                    }
                    else
                    {
                        // Handle both boolean and string IsHidden values
                        bool isHidden = false;
                        if (fenceDict["IsHidden"] is bool boolValue)
                        {
                            isHidden = boolValue;
                        }
                        else if (fenceDict["IsHidden"] is string stringValue)
                        {
                            isHidden = stringValue.ToLower() == "true";
                        }
                        else
                        {
                            Log(LogLevel.Debug, LogCategory.General, $"Invalid IsHidden value '{fenceDict["IsHidden"]}' in {fence.Title}, resetting to \"false\"");
                            isHidden = false;
                            jsonModified = true;
                        }
                        fenceDict["IsHidden"] = isHidden.ToString().ToLower(); // Store as "true" or "false"
                        Log(LogLevel.Debug, LogCategory.General, $"Preserved IsHidden: \"{isHidden.ToString().ToLower()}\" for {fence.Title}");
                    }

                    // Migration: Add or validate IsLink for items
                    if (fence.ItemsType?.ToString() == "Data")
                    {
                        var items = fence.Items as JArray ?? new JArray();
                        bool itemsModified = false;

                        foreach (var item in items)
                        {
                            var itemDict = item as JObject;
                            if (itemDict != null)
                            {
                                if (!itemDict.ContainsKey("IsLink"))
                                {
                                    // For existing items, default to false (assume they're not web links)
                                    itemDict["IsLink"] = false;
                                    itemsModified = true;
                                    string filename = itemDict["Filename"]?.ToString() ?? "Unknown";
                                    Log(LogLevel.Debug, LogCategory.General, $"Added IsLink=false to existing item {filename} in {fence.Title}");
                                }
                                else
                                {
                                    // Validate existing IsLink values - handle JToken properly
                                    bool isLink = false;
                                    var isLinkToken = itemDict["IsLink"];

                                    if (isLinkToken?.Type == JTokenType.Boolean)
                                    {
                                        isLink = isLinkToken.Value<bool>();
                                    }
                                    else if (isLinkToken?.Type == JTokenType.String)
                                    {
                                        string stringValue = isLinkToken.Value<string>();
                                        isLink = stringValue?.ToLower() == "true";
                                    }
                                    else
                                    {
                                        Log(LogLevel.Debug, LogCategory.General, $"Invalid IsLink value '{isLinkToken}' in {fence.Title}, resetting to false");
                                        isLink = false;
                                        itemsModified = true;
                                    }
                                    itemDict["IsLink"] = isLink; // Store as boolean
                                    Log(LogLevel.Debug, LogCategory.General, $"Preserved IsLink: {isLink} for item in {fence.Title}");
                                }
                            }
                        }




                        if (itemsModified)
                        {
                            fenceDict["Items"] = items; // Ensure Items updates are captured
                            jsonModified = true;
                        }
                    }


                    // Migration: Add or validate IsNetwork for items
                    if (fence.ItemsType?.ToString() == "Data")
                    {
                        var items = fence.Items as JArray ?? new JArray();
                        bool itemsModified = false;

                        foreach (var item in items)
                        {
                            var itemDict = item as JObject;
                            if (itemDict != null)
                            {
                                if (!itemDict.ContainsKey("IsNetwork"))
                                {
                                    // For existing items, check if they're network paths
                                    string filename = itemDict["Filename"]?.ToString() ?? "";
                                    bool isNetwork = IsNetworkPath(filename);
                                    itemDict["IsNetwork"] = isNetwork;
                                    itemsModified = true;
                                    Log(LogLevel.Debug, LogCategory.General, $"Added IsNetwork={isNetwork} to existing item {filename} in {fence.Title}");
                                }
                                else
                                {
                                    // Validate existing IsNetwork values
                                    bool isNetwork = false;
                                    var isNetworkToken = itemDict["IsNetwork"];

                                    if (isNetworkToken?.Type == JTokenType.Boolean)
                                    {
                                        isNetwork = isNetworkToken.Value<bool>();
                                    }
                                    else if (isNetworkToken?.Type == JTokenType.String)
                                    {
                                        string stringValue = isNetworkToken.Value<string>();
                                        isNetwork = stringValue?.ToLower() == "true";
                                    }
                                    else
                                    {
                                        Log(LogLevel.Debug, LogCategory.General, $"Invalid IsNetwork value '{isNetworkToken}' in {fence.Title}, resetting to false");
                                        isNetwork = false;
                                        itemsModified = true;
                                    }
                                    itemDict["IsNetwork"] = isNetwork; // Store as boolean
                                    Log(LogLevel.Debug, LogCategory.General, $"Preserved IsNetwork: {isNetwork} for item in {fence.Title}");
                                }
                            }
                        }

                        if (itemsModified)
                        {
                            fenceDict["Items"] = items; // Ensure Items updates are captured
                            jsonModified = true;
                        }
                    }

                    // Update the original fence in _fenceData with the modified dictionary
                    _fenceData[i] = JObject.FromObject(fenceDict);
                }

                if (jsonModified)
                {
                    SaveFenceData();
                    Log(LogLevel.Info, LogCategory.General, "Migrated fences.json with updated fields");
                }
                else
                {
                    Log(LogLevel.Debug, LogCategory.General, "No migration needed for fences.json");
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.Error, $"Error migrating fences.json: {ex.Message}\nStackTrace: {ex.StackTrace}");
            }
        }





        private static void UpdateLockState(TextBlock lockIcon, dynamic fence, bool? forceState = null, bool saveToJson = true)
        {
            // Get the actual fence from _fenceData using Id to ensure correct reference
            string fenceId = fence.Id?.ToString();
            if (string.IsNullOrEmpty(fenceId))
            {
                Log(LogLevel.Info, LogCategory.FenceUpdate, $"Fence '{fence.Title}' has no Id, cannot update lock state");
                return;
            }

            int index = _fenceData.FindIndex(f => f.Id?.ToString() == fenceId);
            if (index < 0)
            {
                Log(LogLevel.Info, LogCategory.FenceUpdate, $"Fence '{fence.Title}' not found in _fenceData, cannot update lock state");
                return;
            }

            dynamic actualFence = _fenceData[index];
            bool isLocked = forceState ?? (actualFence.IsLocked?.ToString().ToLower() == "true");

            // Only update JSON if explicitly requested (e.g., during toggle, not initialization)
            if (saveToJson)
            {
                UpdateFenceProperty(actualFence, "IsLocked", isLocked.ToString().ToLower(), $"Fence {(isLocked ? "locked" : "unlocked")}");
            }

            // Update UI on the main thread
            Application.Current.Dispatcher.Invoke(() =>



            {
                // Update lock icon
                lockIcon.Foreground = isLocked ? System.Windows.Media.Brushes.DeepPink : System.Windows.Media.Brushes.White;
                lockIcon.ToolTip = isLocked ? "Fence is locked (click to unlock)" : "Fence is unlocked (click to lock)";

                // Find the NonActivatingWindow
                NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(lockIcon);
                if (win == null)
                {
                    Log(LogLevel.Info, LogCategory.UI, $"Could not find NonActivatingWindow for fence '{actualFence.Title}'");
                    return;
                }

                // Update ResizeMode
                win.ResizeMode = isLocked ? ResizeMode.NoResize : ResizeMode.CanResizeWithGrip;
                Log(LogLevel.Info, LogCategory.FenceUpdate, $"Set ResizeMode to {win.ResizeMode} for fence '{actualFence.Title}'");
            });

            Log(LogLevel.Debug, LogCategory.UI, $"Updated lock state for fence '{actualFence.Title}': IsLocked={isLocked}");
        }


        private static void CreateFence(dynamic fence, TargetChecker targetChecker)
        {
            // Check for valid Portal Fence target folder
            if (fence.ItemsType?.ToString() == "Portal")
            {
                string targetPath = fence.Path?.ToString();
                if (string.IsNullOrEmpty(targetPath) || !System.IO.Directory.Exists(targetPath))
                {
                    Log(LogLevel.Info, LogCategory.FenceCreation, $"Skipping creation of Portal Fence '{fence.Title}' due to missing target folder: {targetPath ?? "null"}");
                    _fenceData.Remove(fence);
                    SaveFenceData();
                    return;
                }

            }
            DockPanel dp = new DockPanel();
            Border cborder = new Border
            {
                Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 0, 0, 0)),
                CornerRadius = new CornerRadius(6),
                Child = dp
            };



            // NEW: Add heart symbol in top-left corner
            TextBlock heart = new TextBlock
            {
                Text = "♥",
                FontSize = 22,
                Foreground = System.Windows.Media.Brushes.White, // Match title and icon text color
                Margin = new Thickness(5, -3, 0, 0), // Position top-left, aligned with title
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                Cursor = Cursors.Hand
            };
            dp.Children.Add(heart);
            // Store heart TextBlock reference for this fence
            _heartTextBlocks[fence] = heart;


            // Create and assign heart ContextMenu using centralized builder
            heart.ContextMenu = BuildHeartContextMenu(fence);


            // Handle left-click to open heart context menu
            heart.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ChangedButton == MouseButton.Left && heart.ContextMenu != null)
                {
                    heart.ContextMenu.IsOpen = true;
                    e.Handled = true;
                }
            };

            // Add a protection symbol in top-right corner
            TextBlock lockIcon = new TextBlock
            {
                Text = "🛡️",
                FontSize = 14,
                Foreground = fence.IsLocked?.ToString().ToLower() == "true" ? System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.White,
                Margin = new Thickness(0, 3, 2, 0), // Adjusted for top-right positioning
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Cursor = Cursors.Hand,
                ToolTip = fence.IsLocked?.ToString().ToLower() == "true" ? "Fence is locked (click to unlock)" : "Fence is unlocked (click to lock)"
            };



            // Set initial state without saving to JSON
            UpdateLockState(lockIcon, fence, null, saveToJson: false);

            // Lock icon click handler
            lockIcon.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ChangedButton == MouseButton.Left)
                {
                    // Get the NonActivatingWindow to find the fence by Id
                    NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(lockIcon);
                    if (win == null)
                    {
                        Log(LogLevel.Warn, LogCategory.Error, $"Could not find NonActivatingWindow for lock icon click in fence '{fence.Title}'");
                        return;
                    }

                    // Find the fence in _fenceData using the window's Tag (Id)
                    string fenceId = win.Tag?.ToString();
                    if (string.IsNullOrEmpty(fenceId))
                    {
                        Log(LogLevel.Warn, LogCategory.Error, $"Fence Id is missing for window '{win.Title}'");
                        return;
                    }

                    dynamic currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    if (currentFence == null)
                    {
                        Log(LogLevel.Info, LogCategory.FenceUpdate, $"Fence with Id '{fenceId}' not found in _fenceData");
                        return;
                    }

                    // Toggle the lock state
                    bool currentState = currentFence.IsLocked?.ToString().ToLower() == "true";
                    bool newState = !currentState;

                    // Update UI and JSON on the main thread
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        UpdateLockState(lockIcon, currentFence, newState, saveToJson: true);
                    });
                }
            };



            // Create a Grid for the titlebar - move here to ensure it is created before mouse handler
            Grid titleGrid = new Grid
            {
                Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(20, 0, 0, 0))
            };
            titleGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0, GridUnitType.Pixel) }); // No left spacer
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Title takes most space
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30, GridUnitType.Pixel) }); // Lock icon fixed width


            // Handle Ctrl+Click for roll-up/roll-down
            titleGrid.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ChangedButton == MouseButton.Left && Keyboard.IsKeyDown(Key.LeftCtrl))
                {
                    NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(titleGrid);
                    string fenceId = win?.Tag?.ToString();

                    //  DebugLog("IMMEDIATE", fenceId ?? "UNKNOWN", $"Ctrl+Click FIRST LINE - win.Height:{win?.Height ?? -1:F1}");

                    if (string.IsNullOrEmpty(fenceId) || win == null)
                    {
                        //     DebugLog("ERROR", fenceId ?? "UNKNOWN", "Missing window or fenceId in Ctrl+Click");
                        return;
                    }

                    // DebugLog("EVENT", fenceId, "Ctrl+Click TRIGGERED");
                    //  DebugLogFenceState("CTRL_CLICK_START", fenceId, win);

                    dynamic currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    if (currentFence == null)
                    {
                        //     DebugLog("ERROR", fenceId, "Fence not found in _fenceData for Ctrl+Click");
                        return;
                    }

                    bool isRolled = currentFence.IsRolled?.ToString().ToLower() == "true";

                    // Get fence data height (always accurate from SizeChanged handler)
                    double fenceHeight = Convert.ToDouble(currentFence.Height?.ToString() ?? "130");
                    double windowHeight = win.Height;

                    //  DebugLog("SYNC_CHECK", fenceId, $"Before sync - FenceHeight:{fenceHeight:F1} | WindowHeight:{windowHeight:F1} | IsRolled:{isRolled}");

                    if (!isRolled)
                    {
                        // ROLLUP: Use fence data height (always current)
                        //   DebugLog("ACTION", fenceId, "Starting ROLLUP");

                        _fencesInTransition.Add(fenceId);
                        //   DebugLog("TRANSITION", fenceId, "Added to transition state");

                        // FINAL FIX: Use fence data height which is always accurate from SizeChanged handler
                        double currentHeight = fenceHeight;
                        //   DebugLog("ROLLUP_HEIGHT_SOURCE", fenceId, $"Using fence.Height:{fenceHeight:F1} (win.Height was stale:{win.Height:F1})");

                        // Save current fence height as UnrolledHeight
                        IDictionary<string, object> fenceDict = currentFence as IDictionary<string, object> ?? ((JObject)currentFence).ToObject<IDictionary<string, object>>();
                        fenceDict["UnrolledHeight"] = currentHeight;
                        fenceDict["IsRolled"] = "true";

                        int fenceIndex = _fenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                        if (fenceIndex >= 0)
                        {
                            _fenceData[fenceIndex] = JObject.FromObject(fenceDict);
                        }
                        SaveFenceData();

                        //   DebugLog("SAVE", fenceId, $"Saved ROLLUP state | UnrolledHeight:{currentHeight:F1} | IsRolled:true");

                        // Roll up animation - starts from current height
                        double targetHeight = 26;
                        var heightAnimation = new DoubleAnimation(currentHeight, targetHeight, TimeSpan.FromSeconds(0.3))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };

                        heightAnimation.Completed += (animSender, animArgs) =>
                        {
                            //      DebugLog("ANIMATION", fenceId, "ROLLUP animation completed");

                            // Update WrapPanel visibility
                            var border = win.Content as Border;
                            if (border != null)
                            {
                                var dockPanel = border.Child as DockPanel;
                                if (dockPanel != null)
                                {
                                    var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                                    if (scrollViewer != null)
                                    {
                                        var wpcont = scrollViewer.Content as WrapPanel;
                                        if (wpcont != null)
                                        {
                                            wpcont.Visibility = Visibility.Collapsed;
                                            //   DebugLog("UI", fenceId, "Set WrapPanel visibility to Collapsed");
                                        }
                                    }
                                }
                            }

                            _fencesInTransition.Remove(fenceId);
                            //     DebugLog("TRANSITION", fenceId, "Removed from transition state");
                            //  DebugLogFenceState("ROLLUP_COMPLETE", fenceId, win);
                        };

                        win.BeginAnimation(Window.HeightProperty, heightAnimation);
                        //     DebugLog("ANIMATION", fenceId, $"Started ROLLUP animation from {currentHeight:F1} to height {targetHeight:F1}");
                    }
                    else
                    {
                        // ROLLDOWN: Roll down to UnrolledHeight
                        double unrolledHeight = Convert.ToDouble(currentFence.UnrolledHeight?.ToString() ?? "130");
                        // //   DebugLog("ACTION", fenceId, $"Starting ROLLDOWN to {unrolledHeight:F1}");

                        _fencesInTransition.Add(fenceId);
                        //   DebugLog("TRANSITION", fenceId, "Added to transition state");

                        IDictionary<string, object> fenceDict = currentFence as IDictionary<string, object> ?? ((JObject)currentFence).ToObject<IDictionary<string, object>>();
                        fenceDict["IsRolled"] = "false";

                        int fenceIndex = _fenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                        if (fenceIndex >= 0)
                        {
                            _fenceData[fenceIndex] = JObject.FromObject(fenceDict);
                        }
                        SaveFenceData();

                        //   DebugLog("SAVE", fenceId, $"Saved ROLLDOWN state | IsRolled:false | TargetHeight:{unrolledHeight:F1}");

                        var heightAnimation = new DoubleAnimation(win.Height, unrolledHeight, TimeSpan.FromSeconds(0.3))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };

                        heightAnimation.Completed += (animSender, animArgs) =>
                        {
                            //  DebugLog("ANIMATION", fenceId, "ROLLDOWN animation completed");

                            // Update WrapPanel visibility
                            var border = win.Content as Border;
                            if (border != null)
                            {
                                var dockPanel = border.Child as DockPanel;
                                if (dockPanel != null)
                                {
                                    var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                                    if (scrollViewer != null)
                                    {
                                        var wpcont = scrollViewer.Content as WrapPanel;
                                        if (wpcont != null)
                                        {
                                            wpcont.Visibility = Visibility.Visible;
                                            //   DebugLog("UI", fenceId, "Set WrapPanel visibility to Visible");
                                        }
                                    }
                                }
                            }

                            _fencesInTransition.Remove(fenceId);
                            //  DebugLog("TRANSITION", fenceId, "Removed from transition state");
                            //  DebugLogFenceState("ROLLDOWN_COMPLETE", fenceId, win);
                        };

                        win.BeginAnimation(Window.HeightProperty, heightAnimation);
                        //  DebugLog("ANIMATION", fenceId, $"Started ROLLDOWN animation to height {unrolledHeight:F1}");
                    }
                    e.Handled = true;
                }
            };


            ContextMenu CnMnFenceManager = new ContextMenu();
            MenuItem miNewFence = new MenuItem { Header = "New Fence" };
            MenuItem miNewPortalFence = new MenuItem { Header = "New Portal Fence" };
            //  MenuItem miRemoveFence = new MenuItem { Header = "Delete Fence" };
            // MenuItem miXT = new MenuItem { Header = "Exit" };
            MenuItem miHide = new MenuItem { Header = "Hide Fence" }; // New Hide Fence item

            // Add Customize submenu
            MenuItem miCustomize = new MenuItem { Header = "Customize Fence" };
            MenuItem miColors = new MenuItem { Header = "Color" };
            MenuItem miEffects = new MenuItem { Header = "Launch Effect" };

            // Valid options from MigrateLegacyJson
            var validColors = new HashSet<string> { "Red", "Green", "Teal", "Blue", "Bismark", "White", "Beige", "Gray", "Black", "Purple", "Fuchsia", "Yellow", "Orange" };
            var validEffects = Enum.GetNames(typeof(LaunchEffect)).ToHashSet();
            string currentCustomColor = fence.CustomColor?.ToString();
            string currentCustomEffect = fence.CustomLaunchEffect?.ToString();


            // Add color options
            MenuItem miColorDefault = new MenuItem { Header = "Default", Tag = null };
            miColorDefault.Click += (s, e) =>
       {
           // NEW: Uncheck all color items first
           foreach (MenuItem item in miColors.Items)
           {
               item.IsChecked = false;
           }
           // Now check Default
           miColorDefault.IsChecked = true;
           UpdateFenceProperty(fence, "CustomColor", null, "Color set to Default");
       };

            miColorDefault.IsCheckable = true;
            miColorDefault.IsChecked = string.IsNullOrEmpty(currentCustomColor); // Check if null
            miColors.Items.Add(miColorDefault);
            foreach (var color in validColors)
            {
                MenuItem miColor = new MenuItem { Header = color, Tag = color };
                //  miColor.Click += (s, e) => UpdateFenceProperty(fence, "CustomColor", color, $"Color set to {color}");
                miColor.Click += (s, e) =>
                {
                    // Uncheck all color items
                    foreach (MenuItem item in miColors.Items)
                    {
                        item.IsChecked = false;
                    }
                    miColor.IsChecked = true;
                    UpdateFenceProperty(fence, "CustomColor", color, $"Color set to {color}");
                };
                miColor.IsCheckable = true;
                miColor.IsChecked = color.Equals(currentCustomColor, StringComparison.OrdinalIgnoreCase); // Case-insensitive match
                miColors.Items.Add(miColor);
            }

            // Add effect options
            MenuItem miEffectDefault = new MenuItem { Header = "Default", Tag = null };
            // miEffectDefault.Click += (s, e) => UpdateFenceProperty(fence, "CustomLaunchEffect", null, "Launch Effect set to Default");
            miEffectDefault.Click += (s, e) =>
            {
                // Uncheck all effect items
                foreach (MenuItem item in miEffects.Items)
                {
                    item.IsChecked = false;
                }
                miEffectDefault.IsChecked = true;
                UpdateFenceProperty(fence, "CustomLaunchEffect", null, "Launch Effect set to Default");
            };
            miEffectDefault.IsCheckable = true;
            miEffectDefault.IsChecked = string.IsNullOrEmpty(currentCustomEffect); // Check if null
            miEffects.Items.Add(miEffectDefault);
            foreach (var effect in validEffects)
            {
                MenuItem miEffect = new MenuItem { Header = effect, Tag = effect };

                miEffect.Click += (s, e) =>
              {
                  // Uncheck all effect items
                  foreach (MenuItem item in miEffects.Items)
                  {
                      item.IsChecked = false;
                  }
                  miEffect.IsChecked = true;
                  UpdateFenceProperty(fence, "CustomLaunchEffect", effect, $"Launch Effect set to {effect}");
              };
                miEffect.IsCheckable = true;
                miEffect.IsChecked = effect.Equals(currentCustomEffect, StringComparison.OrdinalIgnoreCase); // Case-insensitive match
                miEffects.Items.Add(miEffect);
            }


            miCustomize.Items.Add(miColors);
            miCustomize.Items.Add(miEffects);


            //   CnMnFenceManager.Items.Add(miRemoveFence);
            CnMnFenceManager.Items.Add(miHide); // Add Hide Fence


            NonActivatingWindow win = new NonActivatingWindow
            {
                ContextMenu = CnMnFenceManager,
                AllowDrop = true,
                AllowsTransparency = true,
                Background = System.Windows.Media.Brushes.Transparent,
                Title = fence.Title?.ToString() ?? "New Fence", // Handle null title
                ShowInTaskbar = false,
                WindowStyle = WindowStyle.None,
                Content = cborder,
                ResizeMode = fence.IsLocked?.ToString().ToLower() == "true" ? ResizeMode.NoResize : ResizeMode.CanResizeWithGrip,
                // ResizeMode = ResizeMode.CanResizeWithGrip,
                Width = (double)fence.Width,
                Height = (double)fence.Height,
                Top = (double)fence.Y,
                Left = (double)fence.X,
                Tag = fence.Id?.ToString() ?? Guid.NewGuid().ToString() // Ensure ID exists
            };




            //Peek behind fence

            MenuItem miPeekBehind = new MenuItem { Header = "Peek Behind" };
            CnMnFenceManager.Items.Add(miPeekBehind);

            miPeekBehind.Click += (s, e) =>
            {
                Log(LogLevel.Info, LogCategory.UI, $"Initiating Peek Behind for fence '{fence.Title}'");

                // Create a separate transparent window for the countdown
                var countdownWindow = new Window
                {
                    Width = 60,
                    Height = 40,
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = System.Windows.Media.Brushes.Transparent,
                    ShowInTaskbar = false,
                    Topmost = true,
                    Left = win.Left + (win.Width - 60) / 2, // Center horizontally
                    Top = win.Top + (win.Height - 40) / 2   // Center vertically
                };

                // Create countdown label
                Label countdownLabel = new Label
                {
                    Content = "10",
                    Foreground = System.Windows.Media.Brushes.White,
                    Background = System.Windows.Media.Brushes.Black,
                    Opacity = 0.7,
                    FontSize = 12,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Padding = new Thickness(4)
                };
                countdownWindow.Content = countdownLabel;

                // Show the countdown window
                countdownWindow.Show();
                Log(LogLevel.Debug, LogCategory.UI, $"Created countdown window for fence '{fence.Title}' at ({countdownWindow.Left}, {countdownWindow.Top})");

                // Fade out animation for the fence
                var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.3))
                {
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
                };

                // Countdown timer
                int countdownSeconds = 10;
                var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
                timer.Tick += (timerSender, timerArgs) =>
                {
                    countdownSeconds--;
                    countdownLabel.Content = countdownSeconds.ToString();
                    if (countdownSeconds <= 0)
                    {
                        timer.Stop();
                        // Fade in animation for the fence
                        var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromSeconds(0.8))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                        };
                        win.BeginAnimation(UIElement.OpacityProperty, fadeIn);
                        // Close the countdown window
                        countdownWindow.Close();
                        Log(LogLevel.Info, LogCategory.UI, $"Peek Behind completed for fence '{fence.Title}'");
                    }
                };

                // Start fade out and timer
                win.BeginAnimation(UIElement.OpacityProperty, fadeOut);
                timer.Start();
                Log(LogLevel.Debug, LogCategory.UI, $"Started Peek Behind fade-out and countdown for fence '{fence.Title}'");
            };


            // Add "Open Folder in Explorer" only for portal fences
            if (fence.ItemsType?.ToString() == "Portal")
            {
                MenuItem miOpenFolder = new MenuItem { Header = "Open fence folder" };
                miOpenFolder.Click += (s, e) =>
                {
                    try
                    {
                        string folderPath = fence.Path?.ToString();
                        if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                        {
                            Process.Start("explorer.exe", folderPath);
                            Log(LogLevel.Debug, LogCategory.General, $"Opened folder in Explorer: {folderPath}");
                        }
                        else
                        {
                            Log(LogLevel.Info, LogCategory.General, $"Folder path is invalid or does not exist: {folderPath}");
                            MessageBox.Show("The folder path is invalid or does not exist.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log(LogLevel.Warn, LogCategory.Error, $"Error opening folder in Explorer: {ex.Message}");
                        MessageBox.Show("An error occurred while trying to open the folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                };
                CnMnFenceManager.Items.Add(miOpenFolder);
            }






            CnMnFenceManager.Items.Add(new Separator());
            CnMnFenceManager.Items.Add(miCustomize); // Add Customize submenu
                                                     //  CnMnFenceManager.Items.Add(new Separator());

            //   CnMnFenceManager.Items.Add(new Separator());
            //   CnMnFenceManager.Items.Add(miXT);

            // Handle both JObject and ExpandoObject access
            string isHiddenString = "false"; // Default value
            try
            {
                if (fence is JObject jFence)
                {
                    // Existing fence from JSON
                    isHiddenString = jFence["IsHidden"]?.ToString().ToLower() ?? "false";
                }
                else
                {
                    // New ExpandoObject fence
                    isHiddenString = (fence.IsHidden?.ToString() ?? "false").ToLower();
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warn, LogCategory.Error, $"Error reading IsHidden: {ex.Message}");
            }

            bool isHidden = isHiddenString == "true";
            Log(LogLevel.Debug, LogCategory.Settings, $"Fence '{fence.Title}' IsHidden state: {isHidden}");



            // Adjust the fence position to ensure it fits within screen bounds
            AdjustFencePositionToScreen(win);

            win.Loaded += (s, e) =>
            {
                UpdateLockState(lockIcon, fence, null, saveToJson: false);
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Applied lock state for fence '{fence.Title}' on load: IsLocked={fence.IsLocked?.ToString().ToLower()}");

                // Apply IsRolled state
                bool isRolled = fence.IsRolled?.ToString().ToLower() == "true";
                double targetHeight = 26; // Default for rolled-up state
                if (!isRolled)
                {
                    double unrolledHeight = (double)fence.Height; // Default to fence.Height
                    if (fence.UnrolledHeight != null)
                    {
                        if (double.TryParse(fence.UnrolledHeight.ToString(), out double parsedHeight))
                        {
                            if (parsedHeight > 0)
                            {
                                unrolledHeight = parsedHeight;
                            }
                            else
                            {
                                Log(LogLevel.Debug, LogCategory.FenceCreation, $"UnrolledHeight {parsedHeight} is invalid (non-positive) for fence '{fence.Title}', using Height={unrolledHeight}");
                            }
                        }
                        else
                        {
                            Log(LogLevel.Debug, LogCategory.FenceCreation, $"Failed to parse UnrolledHeight '{fence.UnrolledHeight}' for fence '{fence.Title}', using Height={unrolledHeight}");
                        }
                    }
                    else
                    {
                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"UnrolledHeight is null for fence '{fence.Title}', using Height={unrolledHeight}");
                    }
                    targetHeight = unrolledHeight;
                    Log(LogLevel.Info, LogCategory.FenceCreation, $"Applied rolled-down state for fence '{fence.Title}' on load: Height={targetHeight}");
                }
                else
                {
                    Log(LogLevel.Info, LogCategory.FenceCreation, $"Applied rolled-up state for fence '{fence.Title}' on load: Height={targetHeight}");
                }
                win.Height = targetHeight;

                // Apply WrapPanel visibility
                var border = win.Content as Border;
                if (border != null)
                {
                    var dockPanel = border.Child as DockPanel;
                    if (dockPanel != null)
                    {
                        var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                        if (scrollViewer != null)
                        {
                            var wpcont = scrollViewer.Content as WrapPanel;
                            if (wpcont != null)
                            {
                                wpcont.Visibility = isRolled ? Visibility.Collapsed : Visibility.Visible;
                                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Set initial WrapPanel visibility to {(isRolled ? "Collapsed" : "Visible")} for fence '{fence.Title}'");
                            }
                        }
                    }
                }

                if (isHidden)
                {
                    win.Visibility = Visibility.Hidden;
                    TrayManager.AddHiddenFence(win);
                    Log(LogLevel.Info, LogCategory.FenceCreation, $"Hid fence '{fence.Title}' after loading at startup");
                }
            };
            // Step 5



            if (isHidden)
            {
                TrayManager.AddHiddenFence(win);
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Added fence '{fence.Title}' to hidden list at startup");
            }

            // Hide click
            miHide.Click += (s, e) =>
            {
                UpdateFenceProperty(fence, "IsHidden", "true", $"Hid fence '{fence.Title}'");
                TrayManager.AddHiddenFence(win);
                Log(LogLevel.Debug, LogCategory.General, $"Triggered Hide Fence for '{fence.Title}'");
            };


            // Handle manual resize to update both Height and UnrolledHeight
            win.SizeChanged += (s, e) =>
            {
                // Get current fence reference by ID to avoid stale references
                string fenceId = win.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    //DebugLog("ERROR", fenceId ?? "UNKNOWN", "Fence Id missing during resize");
                    return;
                }

                // DebugLog("EVENT", fenceId, $"SizeChanged TRIGGERED | OldSize:{e.PreviousSize.Width:F1}x{e.PreviousSize.Height:F1} | NewSize:{e.NewSize.Width:F1}x{e.NewSize.Height:F1}");

                // Skip updates if this fence is currently in a rollup/rolldown transition
                if (_fencesInTransition.Contains(fenceId))
                {
                    //    DebugLog("SKIP", fenceId, "Skipping size update - in transition");
                    return;
                }

                // Find the current fence in _fenceData using ID
                var currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence != null)
                {
                    // CRITICAL FIX: Use e.NewSize instead of win.Height/win.Width which can be stale
                    double newHeight = e.NewSize.Height;
                    double newWidth = e.NewSize.Width;
                    double oldHeight = Convert.ToDouble(currentFence.Height?.ToString() ?? "0");
                    double oldUnrolledHeight = Convert.ToDouble(currentFence.UnrolledHeight?.ToString() ?? "0");

                    bool isRolled = currentFence.IsRolled?.ToString().ToLower() == "true";

                    //    DebugLog("VALUES", fenceId, $"FIXED: Using e.NewSize | OldH:{oldHeight:F1} | OldUH:{oldUnrolledHeight:F1} | NewH:{newHeight:F1} | NewW:{newWidth:F1} | IsRolled:{isRolled}");

                    // Update Width and Height with the actual new values
                    currentFence.Width = newWidth;
                    currentFence.Height = newHeight;
                    //   DebugLog("HEIGHT_UPDATE", fenceId, $"Set fence.Height to {newHeight:F1}");
                    //  DebugLog("UPDATE", fenceId, $"Updated Width and Height in fence object to {newWidth:F1}x{newHeight:F1}");

                    // Handle UnrolledHeight update
                    if (!isRolled)
                    {
                        //     DebugLog("LOGIC", fenceId, $"Fence is UNROLLED - checking UnrolledHeight update");

                        if (Math.Abs(newHeight - 26) > 5) // Only if height is significantly different from rolled-up height
                        {
                            double heightDifference = Math.Abs(newHeight - oldUnrolledHeight);
                            //       DebugLog("LOGIC", fenceId, $"Height difference from old UnrolledHeight: {heightDifference:F1}");

                            currentFence.UnrolledHeight = newHeight;
                            //      DebugLog("UPDATE", fenceId, $"UPDATED UnrolledHeight from {oldUnrolledHeight:F1} to {newHeight:F1}");
                        }
                        else
                        {
                            //     DebugLog("SKIP", fenceId, $"Height {newHeight:F1} too close to rolled height (26), not updating UnrolledHeight");
                        }
                    }
                    else
                    {
                        //   DebugLog("SKIP", fenceId, $"Fence is ROLLED UP, not updating UnrolledHeight");
                    }

                    // Save to JSON
                    // DebugLog("SAVE", fenceId, "Calling SaveFenceData()");
                    SaveFenceData();
                    var verifyFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    double verifyHeight = Convert.ToDouble(verifyFence.Height?.ToString() ?? "0");
                    double verifyUnrolled = Convert.ToDouble(verifyFence.UnrolledHeight?.ToString() ?? "0");
                    //   DebugLog("VERIFY_AFTER_SAVE", fenceId, $"After save - fence.Height:{verifyHeight:F1} | fence.UnrolledHeight:{verifyUnrolled:F1}");
                    // Log AFTER state with the correct values
                    //  DebugLog("STATE", fenceId, $"AFTER_UPDATE | ActualHeight:{newHeight:F1} | FenceHeight:{newHeight:F1}", newHeight, Convert.ToDouble(currentFence.UnrolledHeight?.ToString() ?? "0"), isRolled);

                    //  DebugLog("EVENT", fenceId, "SizeChanged COMPLETED");
                }
                else
                {
                    //   DebugLog("ERROR", fenceId, "Fence not found in _fenceData during resize");
                }
            };

            win.Show();



            //miRemoveFence.Click += (s, e) =>
            //{
            //    bool result = TrayManager.Instance.ShowCustomMessageBoxForm(); // Call the method and store the result  
            //                                                                   //  System.Windows.Forms.MessageBox.Show(result.ToString()); // Display the result in a MessageBox  

            //    //   var result = MessageBox.Show("Are you sure you want to remove this fence?", "Confirm", MessageBoxButton.YesNo);
            //    //   if (result == MessageBoxResult.Yes)
            //    if (result == true)
            //    {
            //        // Ensure the backup folder exists
            //        _lastDeletedFolderPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Last Fence Deleted");
            //        if (!Directory.Exists(_lastDeletedFolderPath))
            //        {
            //            Directory.CreateDirectory(_lastDeletedFolderPath);
            //        }

            //        // Clear previous backup
            //        foreach (var file in Directory.GetFiles(_lastDeletedFolderPath))
            //        {
            //            System.IO.File.Delete(file);
            //        }

            //        // Backup the fence and its shortcuts
            //        _lastDeletedFence = fence;
            //        _isRestoreAvailable = true;
            //        UpdateAllHeartContextMenus();

            //        if (fence.ItemsType?.ToString() == "Data")
            //        {
            //            var items = fence.Items as JArray;
            //            if (items != null)
            //            {
            //                foreach (var item in items)
            //                {
            //                    string itemFilePath = item["Filename"]?.ToString();
            //                    if (!string.IsNullOrEmpty(itemFilePath) && System.IO.File.Exists(itemFilePath))
            //                    {
            //                        string shortcutPath = System.IO.Path.Combine(_lastDeletedFolderPath, System.IO.Path.GetFileName(itemFilePath));
            //                        System.IO.File.Copy(itemFilePath, shortcutPath, true);
            //                    }
            //                    else
            //                    {
            //                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Skipped backing up missing file: {itemFilePath}");
            //                    }
            //                }
            //            }
            //        }

            //        // Save fence info to JSON
            //        string fenceJsonPath = System.IO.Path.Combine(_lastDeletedFolderPath, "fence.json");
            //        System.IO.File.WriteAllText(fenceJsonPath, JsonConvert.SerializeObject(fence, Formatting.Indented));

            //        // Proceed with deletion
            //        _fenceData.Remove(fence);
            //        _heartTextBlocks.Remove(fence);
            //        if (_portalFences.ContainsKey(fence))
            //        {
            //            _portalFences[fence].Dispose();
            //            _portalFences.Remove(fence);
            //        }
            //        SaveFenceData();
            //        win.Close();
            //        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Fence {fence.Title} removed successfully");
            //    }
            //};


            miNewFence.Click += (s, e) =>
            {


                System.Windows.Point mousePosition = win.PointToScreen(new System.Windows.Point(0, 0));
                mousePosition = Mouse.GetPosition(win);
                System.Windows.Point absolutePosition = win.PointToScreen(mousePosition);
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Creating new fence at position: X={absolutePosition.X}, Y={absolutePosition.Y}");
                CreateNewFence("New Fence", "Data", absolutePosition.X, absolutePosition.Y);
            };

            miNewPortalFence.Click += (s, e) =>
            {


                System.Windows.Point mousePosition = win.PointToScreen(new System.Windows.Point(0, 0));
                mousePosition = Mouse.GetPosition(win);
                System.Windows.Point absolutePosition = win.PointToScreen(mousePosition);
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Creating new portal fence at position: X={absolutePosition.X}, Y={absolutePosition.Y}");
                CreateNewFence("New Portal Fence", "Portal", absolutePosition.X, absolutePosition.Y);
            };


            Label titlelabel = new Label
            {
                Content = fence.Title.ToString(),
                Foreground = System.Windows.Media.Brushes.White,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Stretch, // Stretch to fill column
                Cursor = Cursors.SizeAll
            };
            Grid.SetColumn(titlelabel, 1);
            titleGrid.Children.Add(titlelabel);

            TextBox titletb = new TextBox
            {
                HorizontalContentAlignment = HorizontalAlignment.Center,
                Visibility = Visibility.Collapsed
            };
            Grid.SetColumn(titletb, 1);
            titleGrid.Children.Add(titletb);

            // Move lockIcon to the Grid
            Grid.SetColumn(lockIcon, 2);
            Grid.SetRow(lockIcon, 0);
            titleGrid.Children.Add(lockIcon);

            // Add the titleGrid to the DockPanel
            DockPanel.SetDock(titleGrid, Dock.Top);
            dp.Children.Add(titleGrid);



            string fenceId = win.Tag?.ToString();
            if (string.IsNullOrEmpty(fenceId))
            {
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Fence Id is missing for window '{win.Title}'");
                return;
            }
            dynamic currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
            if (currentFence == null)
            {
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Fence with Id '{fenceId}' not found in _fenceData");
                return;
            }
            bool isLocked = currentFence.IsLocked?.ToString().ToLower() == "true";

            titlelabel.MouseDown += (sender, e) =>
            {
                if (e.ClickCount == 2)
                {


                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Entering edit mode for fence: {fence.Title}");

                    titletb.Text = titlelabel.Content.ToString();
                    titlelabel.Visibility = Visibility.Collapsed;
                    titletb.Visibility = Visibility.Visible;
                    win.ShowActivated = true;
                    win.Activate();
                    Keyboard.Focus(titletb);
                    titletb.SelectAll();
                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Focus set to title textbox for fence: {fence.Title}");
                }
                //  else if (e.LeftButton == MouseButtonState.Pressed)
                else if (e.LeftButton == MouseButtonState.Pressed)
                {
                    string fenceId = win.Tag?.ToString();
                    if (string.IsNullOrEmpty(fenceId))
                    {
                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Fence Id is missing for window '{win.Title}' during MouseDown");
                        return;
                    }
                    dynamic currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    if (currentFence == null)
                    {
                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Fence with Id '{fenceId}' not found in _fenceData during MouseDown");
                        return;
                    }
                    bool isLocked = currentFence.IsLocked?.ToString().ToLower() == "true";
                    if (!isLocked)
                    {
                        win.DragMove();
                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Dragging fence '{currentFence.Title}'");
                    }
                    else
                    {
                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"DragMove blocked for locked fence '{currentFence.Title}'");
                    }
                }
            };

            titletb.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Enter)
                {


                    fence.Title = titletb.Text;
                    titlelabel.Content = titletb.Text;
                    win.Title = titletb.Text;
                    titletb.Visibility = Visibility.Collapsed;
                    titlelabel.Visibility = Visibility.Visible;
                    SaveFenceData();
                    win.ShowActivated = false;
                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Exited edit mode via Enter, new title for fence: {fence.Title}");
                    win.Focus();
                }
            };

            titletb.LostFocus += (sender, e) =>
            {


                fence.Title = titletb.Text;
                titlelabel.Content = titletb.Text;
                win.Title = titletb.Text;
                titletb.Visibility = Visibility.Collapsed;
                titlelabel.Visibility = Visibility.Visible;
                SaveFenceData();
                win.ShowActivated = false;
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Exited edit mode via click, new title for fence: {fence.Title}");
            };

            WrapPanel wpcont = new WrapPanel();
            ScrollViewer wpcontscr = new ScrollViewer
            {
                Content = wpcont,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            // Προσθήκη watermark για Portal Fences
            if (fence.ItemsType?.ToString() == "Portal")
            {
                if (_options.ShowBackgroundImageOnPortalFences ?? true)
                {
                    wpcontscr.Background = new ImageBrush
                    {
                        ImageSource = new BitmapImage(new Uri("pack://application:,,,/Resources/portal.png")),
                        Opacity = 0.2,
                        Stretch = Stretch.UniformToFill
                    };
                }
            }



            dp.Children.Add(wpcontscr);

            void InitContent()
            {
                // Apply initial visibility based on IsRolled
                bool isRolled = fence.IsRolled?.ToString().ToLower() == "true";
                wpcont.Visibility = isRolled ? Visibility.Collapsed : Visibility.Visible;
                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Set initial WrapPanel visibility to {(isRolled ? "Collapsed" : "Visible")} for fence '{fence.Title}' in InitContent");

                wpcont.Children.Clear();
                if (fence.ItemsType?.ToString() == "Data")
                {
                    var items = fence.Items as JArray;
                    if (items != null)
                    {
                        foreach (dynamic icon in items)
                        {
                            AddIcon(icon, wpcont);
                            StackPanel sp = wpcont.Children[wpcont.Children.Count - 1] as StackPanel;
                            if (sp != null)
                            {
                                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
                                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
                                string arguments = null;
                                if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk")
                                {
                                    WshShell shell = new WshShell();
                                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                                    arguments = shortcut.Arguments;
                                }
                                ClickEventAdder(sp, filePath, isFolder, arguments);

                                // targetChecker.AddCheckAction(filePath, () => UpdateIcon(sp, filePath, isFolder), isFolder);
                                // Only add to TargetChecker if it's not a web link
                                //bool itemIsLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];
                                //if (!itemIsLink)
                                //{
                                //    targetChecker.AddCheckAction(filePath, () => UpdateIcon(sp, filePath, isFolder), isFolder);
                                //}
                                //else
                                //{
                                //    Log(LogLevel.Debug, LogCategory.IconHandling, $"Excluded web link {filePath} from target checking");
                                //}
                                // Only add to TargetChecker based on type and settings
                                bool itemIsLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];
                                bool itemIsNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"];
                                bool allowNetworkChecking = _options.CheckNetworkPaths ?? false;

                                if (!itemIsLink && (!itemIsNetwork || allowNetworkChecking))
                                {
                                    targetChecker.AddCheckAction(filePath, () => UpdateIcon(sp, filePath, isFolder), isFolder);
                                    if (itemIsNetwork && allowNetworkChecking)
                                    {
                                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Added network path {filePath} to target checking (user enabled)");
                                    }
                                }
                                else
                                {
                                    if (itemIsLink)
                                    {
                                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Excluded web link {filePath} from target checking");
                                    }
                                    if (itemIsNetwork && !allowNetworkChecking)
                                    {
                                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Excluded network path {filePath} from target checking (user setting)");
                                    }
                                }

                                ContextMenu mn = new ContextMenu();
                                MenuItem miRunAsAdmin = new MenuItem { Header = "Run as administrator" };
                                MenuItem miEdit = new MenuItem { Header = "Edit..." };
                                MenuItem miMove = new MenuItem { Header = "Move..." };
                                MenuItem miRemove = new MenuItem { Header = "Remove" };
                                MenuItem miFindTarget = new MenuItem { Header = "Open target folder..." };
                                MenuItem miCopyPath = new MenuItem { Header = "Copy path" };
                                MenuItem miCopyFolder = new MenuItem { Header = "Folder path" };
                                MenuItem miCopyFullPath = new MenuItem { Header = "Full path" };
                                miCopyPath.Items.Add(miCopyFolder);
                                miCopyPath.Items.Add(miCopyFullPath);

                                mn.Items.Add(miEdit);
                                mn.Items.Add(miMove);
                                mn.Items.Add(miRemove);
                                mn.Items.Add(new Separator());
                                mn.Items.Add(miRunAsAdmin);
                                mn.Items.Add(new Separator());
                                mn.Items.Add(miCopyPath);
                                mn.Items.Add(miFindTarget);
                                sp.ContextMenu = mn;

                                miRunAsAdmin.IsEnabled = Utility.IsExecutableFile(filePath);
                                miMove.Click += (sender, e) => MoveItem(icon, fence, win.Dispatcher);
                                miEdit.Click += (sender, e) => EditItem(icon, fence, win);
                                miRemove.Click += (sender, e) =>
                                {
                                    var items = fence.Items as JArray;
                                    if (items != null)
                                    {
                                        var itemToRemove = items.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                                        if (itemToRemove != null)
                                        {

                                            Log(LogLevel.Debug, LogCategory.FenceCreation, $"Removing icon for {filePath} from fence");
                                            var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.3));
                                            fade.Completed += (s, a) =>
                                            {
                                                items.Remove(itemToRemove);
                                                wpcont.Children.Remove(sp);
                                                targetChecker.RemoveCheckAction(filePath);
                                                SaveFenceData();

                                                string shortcutPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Shortcuts", System.IO.Path.GetFileName(filePath));
                                                if (System.IO.File.Exists(shortcutPath))
                                                {
                                                    try
                                                    {
                                                        System.IO.File.Delete(shortcutPath);
                                                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Deleted shortcut: {shortcutPath}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Log(LogLevel.Error, LogCategory.General, $"Failed to delete shortcut {shortcutPath}: {ex.Message}");
                                                    }
                                                }
                                                // Delete backup shortcut if it exists
                                                string tempShortcutsDir = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Temp Shortcuts");
                                                string backupPath = System.IO.Path.Combine(tempShortcutsDir, System.IO.Path.GetFileName(filePath));
                                                if (System.IO.File.Exists(backupPath))
                                                {
                                                    try
                                                    {
                                                        System.IO.File.Delete(backupPath);
                                                        Log(LogLevel.Debug, LogCategory.General, $"Deleted backup shortcut: {backupPath}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Log(LogLevel.Error, LogCategory.General, $"Failed to delete backup shortcut {backupPath}: {ex.Message}");
                                                    }
                                                }





                                            };
                                            sp.BeginAnimation(UIElement.OpacityProperty, fade);
                                        }
                                    }
                                };

                                miFindTarget.Click += (sender, e) =>
                                {
                                    string target = Utility.GetShortcutTarget(filePath);
                                    if (!string.IsNullOrEmpty(target) && (System.IO.File.Exists(target) || System.IO.Directory.Exists(target)))
                                    {
                                        Process.Start("explorer.exe", $"/select,\"{target}\"");
                                    }
                                };

                                miCopyFolder.Click += (sender, e) =>
                                {
                                    string folderPath = System.IO.Path.GetDirectoryName(Utility.GetShortcutTarget(filePath));
                                    Clipboard.SetText(folderPath);
                                    bool isLogEnabled = _options.IsLogEnabled ?? true;
                                    if (isLogEnabled)
                                    {
                                        string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                                        System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: Copied target folder path to clipboard: {folderPath}\n");
                                    }
                                };

                                miCopyFullPath.Click += (sender, e) =>
                                {
                                    string targetPath = Utility.GetShortcutTarget(filePath);
                                    Clipboard.SetText(targetPath);
                                    bool isLogEnabled = _options.IsLogEnabled ?? true;
                                    if (isLogEnabled)
                                    {
                                        string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                                        System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: Copied target full path to clipboard: {targetPath}\n");
                                    }
                                };

                                miRunAsAdmin.Click += (sender, e) =>
                                {
                                    string targetPath = Utility.GetShortcutTarget(filePath);
                                    string runArguments = null;
                                    if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk")
                                    {
                                        WshShell shell = new WshShell();
                                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                                        runArguments = shortcut.Arguments;
                                    }
                                    ProcessStartInfo psi = new ProcessStartInfo
                                    {
                                        FileName = targetPath,
                                        UseShellExecute = true,
                                        Verb = "runas"
                                    };
                                    if (!string.IsNullOrEmpty(runArguments))
                                    {
                                        psi.Arguments = runArguments;
                                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Run as admin with arguments: {runArguments} for {filePath}");
                                    }
                                    try
                                    {
                                        Process.Start(psi);
                                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Successfully launched {targetPath} as admin");
                                    }
                                    catch (Exception ex)
                                    {
                                        Log(LogLevel.Error, LogCategory.General, $"Failed to launch {targetPath} as admin: {ex.Message}");
                                        // MessageBox.Show($"Error running as admin: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                        TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Error running as admin: {ex.Message}", "Error");
                                    }
                                };
                            }
                        }
                    }
                }
                else if (fence.ItemsType?.ToString() == "Portal")
                {
                    try
                    {
                        _portalFences[fence] = new PortalFenceManager(fence, wpcont);
                    }
                    catch (Exception ex)
                    {
                        // MessageBox.Show($"Failed to initialize Portal Fence: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to initialize Portal Fence: {ex.Message}", "Error");
                        _fenceData.Remove(fence);
                        SaveFenceData();
                        win.Close();
                    }
                }
            }
            win.Drop += (sender, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Handled = true;
                    string[] droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
                    foreach (string droppedFile in droppedFiles)
                    {
                        try
                        {
                            //  Debug.WriteLine($"Dropped file: {droppedFile}");
                            if (!System.IO.File.Exists(droppedFile) && !System.IO.Directory.Exists(droppedFile))
                            {
                                //  MessageBox.Show($"Invalid file or directory: {droppedFile}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Invalid file or directory: {droppedFile}", "Error");
                                continue;
                            }



                            if (fence.ItemsType?.ToString() == "Data")
                            {
                                // Logic for Data Fences (shortcuts)
                                if (!System.IO.Directory.Exists("Shortcuts")) System.IO.Directory.CreateDirectory("Shortcuts");
                                string baseShortcutName = System.IO.Path.Combine("Shortcuts", System.IO.Path.GetFileName(droppedFile));
                                string shortcutName = baseShortcutName;
                                int counter = 1;

                                bool isDroppedShortcut = System.IO.Path.GetExtension(droppedFile).ToLower() == ".lnk";
                                bool isDroppedUrlFile = System.IO.Path.GetExtension(droppedFile).ToLower() == ".url";
                                bool isWebLink = IsWebLinkShortcut(droppedFile); // Use our new helper method

                                string targetPath;
                                bool isFolder = false;
                                string webUrl = null;

                                if (isWebLink)
                                {
                                    // Handle web link shortcuts
                                    webUrl = ExtractWebUrlFromFile(droppedFile);
                                    targetPath = webUrl; // For web links, target is the URL
                                    isFolder = false; // Web links are never folders
                                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Detected web link: {droppedFile} -> {webUrl}");
                                }
                                else
                                {
                                    // Handle regular file/folder shortcuts
                                    targetPath = isDroppedShortcut ? Utility.GetShortcutTarget(droppedFile) : droppedFile;
                                    isFolder = System.IO.Directory.Exists(targetPath) || (isDroppedShortcut && string.IsNullOrEmpty(System.IO.Path.GetExtension(targetPath)));
                                }

                                if (!isDroppedShortcut && !isDroppedUrlFile)
                                {
                                    // Create shortcut for dropped file/folder
                                    shortcutName = baseShortcutName + ".lnk";
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{System.IO.Path.GetFileNameWithoutExtension(droppedFile)} ({counter++}).lnk");
                                    }
                                    WshShell shell = new WshShell();
                                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutName);
                                    shortcut.TargetPath = droppedFile;
                                    shortcut.Save();
                                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Created unique shortcut: {shortcutName}");
                                }
                                else
                                {
                                    // Handle dropped shortcut files
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{System.IO.Path.GetFileNameWithoutExtension(droppedFile)} ({counter++}).lnk");
                                    }

                                    if (isWebLink)
                                    {
                                        // Create new shortcut targeting the web URL directly
                                        CreateWebLinkShortcut(webUrl, shortcutName, System.IO.Path.GetFileNameWithoutExtension(droppedFile));
                                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Created web link shortcut: {shortcutName} -> {webUrl}");
                                    }
                                    else
                                    {
                                        // Copy regular shortcut
                                        System.IO.File.Copy(droppedFile, shortcutName, false);
                                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Copied unique shortcut: {shortcutName}");
                                    }
                                }

                                dynamic newItem = new System.Dynamic.ExpandoObject();
                                IDictionary<string, object> newItemDict = newItem;
                                newItemDict["Filename"] = shortcutName;
                                newItemDict["IsFolder"] = isFolder;
                                newItemDict["IsLink"] = isWebLink; // Set IsLink property
                                newItemDict["IsNetwork"] = IsNetworkPath(shortcutName); // Set IsNetwork property
                                newItemDict["DisplayName"] = System.IO.Path.GetFileNameWithoutExtension(droppedFile);

                                // Log the detection results
                                if ((bool)newItemDict["IsNetwork"])
                                {
                                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Detected network path for new item: {shortcutName}");
                                }

                                var items = fence.Items as JArray ?? new JArray();
                                items.Add(JObject.FromObject(newItem));
                                fence.Items = items;
                                AddIcon(newItem, wpcont);
                                StackPanel sp = wpcont.Children[wpcont.Children.Count - 1] as StackPanel;
                                if (sp != null)
                                {
                                    ClickEventAdder(sp, shortcutName, isFolder);
                                    targetChecker.AddCheckAction(shortcutName, () => UpdateIcon(sp, shortcutName, isFolder), isFolder);

                                    // Προσθήκη Context Menu
                                    ContextMenu mn = new ContextMenu();
                                    MenuItem miRunAsAdmin = new MenuItem { Header = "Run as administrator" };
                                    MenuItem miEdit = new MenuItem { Header = "Edit" };
                                    MenuItem miMove = new MenuItem { Header = "Move.." };
                                    MenuItem miRemove = new MenuItem { Header = "Remove" };
                                    MenuItem miFindTarget = new MenuItem { Header = "Find target..." };
                                    MenuItem miCopyPath = new MenuItem { Header = "Copy path" };
                                    MenuItem miCopyFolder = new MenuItem { Header = "Folder" };
                                    MenuItem miCopyFullPath = new MenuItem { Header = "Full path" };
                                    miCopyPath.Items.Add(miCopyFolder);
                                    miCopyPath.Items.Add(miCopyFullPath);

                                    mn.Items.Add(miEdit);
                                    mn.Items.Add(miMove);
                                    mn.Items.Add(miRemove);
                                    mn.Items.Add(new Separator());
                                    mn.Items.Add(miRunAsAdmin);
                                    mn.Items.Add(new Separator());
                                    mn.Items.Add(miCopyPath);
                                    mn.Items.Add(miFindTarget);
                                    sp.ContextMenu = mn;

                                    miRunAsAdmin.IsEnabled = Utility.IsExecutableFile(shortcutName);
                                    miMove.Click += (sender, e) => MoveItem(newItem, fence, win.Dispatcher);
                                    miEdit.Click += (sender, e) => EditItem(newItem, fence, win);
                                    miRemove.Click += (sender, e) =>
                                    {
                                        var items = fence.Items as JArray;
                                        if (items != null)
                                        {
                                            var itemToRemove = items.FirstOrDefault(i => i["Filename"]?.ToString() == shortcutName);
                                            if (itemToRemove != null)
                                            {
                                                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Removing icon for {shortcutName} from fence");
                                                var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.3));
                                                fade.Completed += (s, a) =>
                                                {
                                                    items.Remove(itemToRemove);
                                                    wpcont.Children.Remove(sp);
                                                    targetChecker.RemoveCheckAction(shortcutName);
                                                    SaveFenceData();

                                                    string shortcutPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Shortcuts", System.IO.Path.GetFileName(shortcutName));
                                                    if (System.IO.File.Exists(shortcutPath))
                                                    {
                                                        try
                                                        {
                                                            System.IO.File.Delete(shortcutPath);
                                                            Log(LogLevel.Debug, LogCategory.FenceCreation, $"Deleted shortcut: {shortcutPath}");
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Log(LogLevel.Error, LogCategory.General, $"Failed to delete shortcut {shortcutPath}: {ex.Message}");
                                                        }
                                                    }
                                                };
                                                sp.BeginAnimation(UIElement.OpacityProperty, fade);
                                            }
                                        }
                                    };

                                    miFindTarget.Click += (sender, e) =>
                                    {
                                        string target = Utility.GetShortcutTarget(shortcutName);
                                        if (!string.IsNullOrEmpty(target) && (System.IO.File.Exists(target) || System.IO.Directory.Exists(target)))
                                        {
                                            Process.Start("explorer.exe", $"/select,\"{target}\"");
                                        }
                                    };

                                    miCopyFolder.Click += (sender, e) =>
                                    {
                                        string folderPath = System.IO.Path.GetDirectoryName(Utility.GetShortcutTarget(shortcutName));
                                        Clipboard.SetText(folderPath);
                                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Copied target folder path to clipboard: {folderPath}");
                                    };

                                    miCopyFullPath.Click += (sender, e) =>
                                    {
                                        string targetPath = Utility.GetShortcutTarget(shortcutName);
                                        Clipboard.SetText(targetPath);
                                        Log(LogLevel.Debug, LogCategory.FenceCreation, $"Copied target full path to clipboard: {targetPath}");
                                    };

                                    miRunAsAdmin.Click += (sender, e) =>
                                    {
                                        Process.Start(new ProcessStartInfo
                                        {
                                            FileName = Utility.GetShortcutTarget(shortcutName),
                                            UseShellExecute = true,
                                            Verb = "runas"
                                        });
                                    };
                                }
                                Log(LogLevel.Debug, LogCategory.FenceCreation, $"Added shortcut to Data Fence: {shortcutName}");
                            }
                            else if (fence.ItemsType?.ToString() == "Portal")
                            {
                                // Λογική για Portal Fences (copy) - παραμένει ίδια
                                IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
                                string destinationFolder = fenceDict.ContainsKey("Path") ? fenceDict["Path"]?.ToString() : null;

                                if (string.IsNullOrEmpty(destinationFolder))
                                {
                                    //MessageBox.Show($"No destination folder defined for this Portal Fence. Please recreate the fence.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    TrayManager.Instance.ShowOKOnlyMessageBoxForm($"No destination folder defined for this Portal Fence. Please recreate the fence.", "Error");
                                    Log(LogLevel.Warn, LogCategory.FenceCreation, $"No Path defined for Portal Fence: {fence.Title}");
                                    continue;
                                }

                                if (!System.IO.Directory.Exists(destinationFolder))
                                {
                                    // MessageBox.Show($"The destination folder '{destinationFolder}' no longer exists. Please update the Portal Fence settings.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    TrayManager.Instance.ShowOKOnlyMessageBoxForm($"The destination folder '{destinationFolder}' no longer exists. Please update the Portal Fence settings.", "Error");
                                    Log(LogLevel.Warn, LogCategory.FenceCreation, $"Destination folder missing for Portal Fence: {destinationFolder}");
                                    continue;
                                }

                                string destinationPath = System.IO.Path.Combine(destinationFolder, System.IO.Path.GetFileName(droppedFile));
                                int counter = 1;
                                string baseName = System.IO.Path.GetFileNameWithoutExtension(droppedFile);
                                string extension = System.IO.Path.GetExtension(droppedFile);

                                while (System.IO.File.Exists(destinationPath) || System.IO.Directory.Exists(destinationPath))
                                {
                                    destinationPath = System.IO.Path.Combine(destinationFolder, $"{baseName} ({counter}){extension}");
                                    counter++;
                                }

                                if (System.IO.File.Exists(droppedFile))
                                {
                                    System.IO.File.Copy(droppedFile, destinationPath, false);
                                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Copied file to Portal Fence: {destinationPath}");
                                }
                                else if (System.IO.Directory.Exists(droppedFile))
                                {
                                    CopyDirectory(droppedFile, destinationPath);
                                    Log(LogLevel.Debug, LogCategory.FenceCreation, $"Copied directory to Portal Fence: {destinationPath}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                              //  MessageBox.Show($"Failed to add {droppedFile}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to add {droppedFile}: {ex.Message}", "Error");
                            Log(LogLevel.Error, LogCategory.FenceCreation, $"Failed to add {droppedFile}: {ex.Message}");
                        }
                    }
                    SaveFenceData(); // Αποθηκεύουμε τις αλλαγές στο JSON
                }
            };


            if (SettingsManager.EnableDimensionSnap)
            {
                win.SizeChanged += UpdateSizeFeedback;
            }




            win.LocationChanged += (s, e) =>
            {
                // Get current fence reference by ID to avoid stale references
                string fenceId = win.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    Log(LogLevel.Warn, LogCategory.FenceUpdate, $"Fence Id missing during position change for window '{win.Title}'");
                    return;
                }

                // Find the current fence in _fenceData using ID
                var currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence != null)
                {
                    // Update position and save immediately
                    currentFence.X = win.Left;
                    currentFence.Y = win.Top;
                    SaveFenceData();
                    Log(LogLevel.Debug, LogCategory.FenceUpdate, $"Position updated for fence '{currentFence.Title}' to X={win.Left}, Y={win.Top}");
                }
                else
                {
                    Log(LogLevel.Warn, LogCategory.FenceUpdate, $"Fence with Id '{fenceId}' not found during position change");
                }
            };

            InitContent();
            win.Show();


            IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
            SnapManager.AddSnapping(win, fenceDict);
            // Apply custom color if present, otherwise use global
            string customColor = fence.CustomColor?.ToString();
            Utility.ApplyTintAndColorToFence(win, string.IsNullOrEmpty(customColor) ? _options.SelectedColor : customColor);
            targetChecker.Start();
        }

        // Βοηθητική μέθοδος για αντιγραφή φακέλων
        private static void CopyDirectory(string sourceDir, string destDir)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDir);
            DirectoryInfo[] dirs = dir.GetDirectories();

            Directory.CreateDirectory(destDir);

            foreach (FileInfo file in dir.GetFiles())
            {
                string targetFilePath = System.IO.Path.Combine(destDir, file.Name);
                file.CopyTo(targetFilePath, false);
            }

            foreach (DirectoryInfo subDir in dirs)
            {
                string newDestDir = System.IO.Path.Combine(destDir, subDir.Name);
                CopyDirectory(subDir.FullName, newDestDir);
            }
        }
        // Return here
        private static void ShowSizeFeedback(double width, double height)
        {
            if (_sizeFeedbackWindow == null)
            {
                _sizeFeedbackWindow = new Window
                {
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = System.Windows.Media.Brushes.Transparent,
                    Width = 100,
                    Height = 30,
                    ShowInTaskbar = false,
                    Topmost = true
                };

                var label = new Label
                {
                    Content = "",
                    Foreground = System.Windows.Media.Brushes.White,
                    Background = System.Windows.Media.Brushes.Black,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                _sizeFeedbackWindow.Content = label;
            }

            var labelContent = (Label)_sizeFeedbackWindow.Content;
            labelContent.Content = $"{Math.Round(width)} x {Math.Round(height)}";

            var mousePos = System.Windows.Forms.Cursor.Position;
            _sizeFeedbackWindow.Left = mousePos.X + 10;
            _sizeFeedbackWindow.Top = mousePos.Y + 10;

            _sizeFeedbackWindow.Show();
        }

        private static void HideSizeFeedback()
        {
            if (_sizeFeedbackWindow != null)
            {
                _sizeFeedbackWindow.Hide();
            }
        }


        public static void OnResizingStarted(NonActivatingWindow fence)
        {
            if (SettingsManager.EnableDimensionSnap)
            {
                fence.SizeChanged += UpdateSizeFeedback;
                ShowSizeFeedback(fence.Width, fence.Height);
            }
        }

        public static void OnResizingEnded(NonActivatingWindow fence)
        {
            if (SettingsManager.EnableDimensionSnap)
            {
                fence.SizeChanged -= UpdateSizeFeedback;
                HideSizeFeedback();

                double snappedWidth = Math.Round(fence.Width / 10.0) * 10;
                double snappedHeight = Math.Round(fence.Height / 10.0) * 10;

                fence.Width = snappedWidth;
                fence.Height = snappedHeight;

                dynamic fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == fence.Title);
                if (fenceData != null)
                {
                    fenceData.Width = snappedWidth;
                    fenceData.Height = snappedHeight;
                    FenceManager.SaveFenceData();
                }

                ShowSizeFeedback(snappedWidth, snappedHeight);
                _hideTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(2)
                };
                _hideTimer.Tick += (s, e) =>
                {
                    HideSizeFeedback();
                    _hideTimer.Stop();
                };
                _hideTimer.Start();
            }
        }

        private static void UpdateSizeFeedback(object sender, SizeChangedEventArgs e)
        {
            var fence = sender as NonActivatingWindow;
            if (fence != null)
            {
                ShowSizeFeedback(fence.Width, fence.Height);
            }
        }





        public static void AddIcon(dynamic icon, WrapPanel wpcont)
        {
            StackPanel sp = new StackPanel { Margin = new Thickness(5), Width = 60 };
            System.Windows.Controls.Image ico = new System.Windows.Controls.Image { Width = 40, Height = 40, Margin = new Thickness(5) };
            //IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            //string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            //bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];

            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
            bool isLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"]; // Add IsLink detection
            bool isNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"]; // Add IsNetwork detection

            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
            string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;
            string arguments = iconDict.ContainsKey("Arguments") ? (string)iconDict["Arguments"] : null;

            ImageSource shortcutIcon = null;
            // Handle web link icons first (before other shortcut processing)
            if (isLink)
            {
                shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));
                Log(LogLevel.Debug, LogCategory.IconHandling, $"Using link-White.png for web link {filePath}");
            }
            else if (isShortcut)

            {
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                targetPath = shortcut.TargetPath;
                arguments = shortcut.Arguments;

                // Handle custom IconLocation with index
                if (!string.IsNullOrEmpty(shortcut.IconLocation))
                {
                    string[] iconParts = shortcut.IconLocation.Split(',');
                    string iconPath = iconParts[0];
                    int iconIndex = 0;
                    if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex))
                    {
                        iconIndex = parsedIndex;
                    }

                    if (System.IO.File.Exists(iconPath))
                    {
                        try
                        {
                            IntPtr[] hIcon = new IntPtr[1];
                            uint result = ExtractIconEx(iconPath, iconIndex, hIcon, null, 1);
                            if (result > 0 && hIcon[0] != IntPtr.Zero)
                            {
                                try
                                {
                                    shortcutIcon = Imaging.CreateBitmapSourceFromHIcon(
                                        hIcon[0],
                                        Int32Rect.Empty,
                                        BitmapSizeOptions.FromEmptyOptions()
                                    );
                                    Log(LogLevel.Info, LogCategory.IconHandling, $"Extracted icon at index {iconIndex} from {iconPath} for {filePath}");
                                }
                                finally
                                {
                                    DestroyIcon(hIcon[0]);
                                }
                            }
                            else
                            {
                                Log(LogLevel.Warn, LogCategory.IconHandling, $"Failed to extract icon at index {iconIndex} from {iconPath} for {filePath}. Result: {result}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log(LogLevel.Error, LogCategory.IconHandling, $"Error extracting icon from {iconPath} at index {iconIndex} for {filePath}: {ex.Message}");
                        }
                    }
                    else
                    {
                        Log(LogLevel.Info, LogCategory.IconHandling, $"Icon file not found: {iconPath} for {filePath}");
                    }
                }

                // Fallback only if no custom icon was successfully extracted
                if (shortcutIcon == null)
                {
                    if (System.IO.Directory.Exists(targetPath))
                    {
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Using folder-White.png for shortcut {filePath} targeting folder {targetPath}");
                    }
                    else if (System.IO.File.Exists(targetPath))
                    {
                        try
                        {
                            shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                            // do not remove the below commented log line, it may needed for future debugging
                            // Log(LogLevel.Debug, LogCategory.IconHandling, $"Using target file icon for {filePath}: {targetPath}");
                        }
                        catch (Exception ex)
                        {
                            shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                            Log(LogLevel.Warn, LogCategory.IconHandling, $"Failed to extract target file icon for {filePath}: {ex.Message}");
                        }
                    }
                    else
                    {
                        shortcutIcon = isFolder
                            ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                            : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")}");
                    }
                }
            }
            else
            {
                // Non-shortcut handling remains unchanged
                if (System.IO.Directory.Exists(filePath))
                {
                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"Using folder-White.png for {filePath}");
                }
                else if (System.IO.File.Exists(filePath))
                {
                    try
                    {
                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource();
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Using file icon for {filePath}");
                    }
                    catch (Exception ex)
                    {
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        Log(LogLevel.Warn, LogCategory.IconHandling, $"Failed to extract icon for {filePath}: {ex.Message}");
                    }
                }
                else
                {
                    shortcutIcon = isFolder
                        ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                        : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")}");
                }
            }

            ico.Source = shortcutIcon;
        //    sp.Children.Add(ico);


            // Add network indicator overlay if this is a network path
            if (isNetwork)
            {
                // Create a Grid to overlay the network indicator on the icon
                Grid iconGrid = new Grid
                {
                    Width = 48,
                    Height = 48,
                    Margin = new Thickness(2)
                };

                // Add the icon to the grid (don't add to StackPanel first)
                iconGrid.Children.Add(ico);

                // Create network indicator (🔗 symbol in royal blue)
                TextBlock networkIndicator = new TextBlock
                {
                    Text = "🔗",
                    FontSize = 14,
                   
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(65, 135, 225)), // Royal Blue #4169E1
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(2, 2, 0, 0), // Slight offset to position in top-left corner

                    Effect = new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 315, // Top-left to bottom-right shadow
                        ShadowDepth = 1.5,
                        BlurRadius = 2,
                        Opacity = 0.7
                    }
                };

                iconGrid.Children.Add(networkIndicator);

                // Add the grid to StackPanel instead of the icon
                sp.Children.Add(iconGrid);

                Log(LogLevel.Debug, LogCategory.IconHandling, $"Added network indicator to {filePath}");
            }
            else
            {
                // For non-network items, add icon directly as before
                sp.Children.Add(ico);
            }

            string displayText = (!iconDict.ContainsKey("DisplayName") || iconDict["DisplayName"] == null || string.IsNullOrEmpty((string)iconDict["DisplayName"]))
                ? System.IO.Path.GetFileNameWithoutExtension(filePath)
                : (string)iconDict["DisplayName"];
            if (displayText.Length > 20)
            {
                displayText = displayText.Substring(0, 20) + "...";
            }

            TextBlock lbl = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                TextTrimming = TextTrimming.None,
                HorizontalAlignment = HorizontalAlignment.Center,
                Foreground = System.Windows.Media.Brushes.White,
                MaxWidth = double.MaxValue,
                Width = double.NaN,
                TextAlignment = TextAlignment.Center,
                Text = displayText
            };
            lbl.Effect = new DropShadowEffect
            {
                Color = Colors.Black,
                Direction = 315,
                ShadowDepth = 2,
                BlurRadius = 3,
                Opacity = 0.8
            };

            sp.Children.Add(lbl);

            sp.Tag = filePath;
            string toolTipText = $"{System.IO.Path.GetFileName(filePath)}\nTarget: {targetPath ?? "N/A"}";
            if (!string.IsNullOrEmpty(arguments))
            {
                toolTipText += $"\nParameters: {arguments}";
            }
            sp.ToolTip = new ToolTip { Content = toolTipText };

            wpcont.Children.Add(sp);
        }


        public static void SaveFenceData()
        {
            var serializedData = new List<JObject>();
            foreach (dynamic fence in _fenceData)
            {
                IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
                // Convert IsHidden to string format
                if (fenceDict.ContainsKey("IsHidden"))
                {
                    bool isHidden = false;
                    if (fenceDict["IsHidden"] is bool boolValue)
                    {
                        isHidden = boolValue;
                    }
                    else if (fenceDict["IsHidden"] is string stringValue)
                    {
                        isHidden = stringValue.ToLower() == "true";
                    }
                    fenceDict["IsHidden"] = isHidden.ToString().ToLower();
                }
                // Convert IsRolled to string format
                if (fenceDict.ContainsKey("IsRolled"))
                {
                    bool isRolled = false;
                    if (fenceDict["IsRolled"] is bool boolValue)
                    {
                        isRolled = boolValue;
                    }
                    else if (fenceDict["IsRolled"] is string stringValue)
                    {
                        isRolled = stringValue.ToLower() == "true";
                    }
                    fenceDict["IsRolled"] = isRolled.ToString().ToLower();
                }
                // Ensure UnrolledHeight is a valid number
                if (fenceDict.ContainsKey("UnrolledHeight"))
                {
                    double unrolledHeight;
                    if (!double.TryParse(fenceDict["UnrolledHeight"]?.ToString(), out unrolledHeight) || unrolledHeight <= 0)
                    {
                        unrolledHeight = fenceDict.ContainsKey("Height") ? Convert.ToDouble(fenceDict["Height"]) : 130;
                        fenceDict["UnrolledHeight"] = unrolledHeight;
                    }
                    else
                    {
                        fenceDict["UnrolledHeight"] = unrolledHeight;
                    }
                }
                serializedData.Add(JObject.FromObject(fenceDict));
            }

            string formattedJson = JsonConvert.SerializeObject(serializedData, Formatting.Indented);
            System.IO.File.WriteAllText(_jsonFilePath, formattedJson);
            Log(LogLevel.Debug, LogCategory.Settings, $"Saved fences.json with consistent IsHidden, IsRolled, and UnrolledHeight string format");
        }


        private static void CreateNewFence(string title, string itemsType, double x = 20, double y = 20, string customColor = null, string customLaunchEffect = null)
        {
            // Generate random name instead of using the passed title
            string fenceName = GenerateRandomName();


            dynamic newFence = new System.Dynamic.ExpandoObject();
            newFence.Id = Guid.NewGuid().ToString();
            IDictionary<string, object> newFenceDict = newFence;
            // newFenceDict["Title"] = title;
            newFenceDict["Title"] = fenceName; // Use random name
            newFenceDict["X"] = x;
            newFenceDict["Y"] = y;
            newFenceDict["Width"] = 230;
            newFenceDict["Height"] = 130;
            newFenceDict["ItemsType"] = itemsType;
            newFenceDict["Items"] = itemsType == "Portal" ? "" : new JArray();
            newFenceDict["CustomColor"] = customColor; // Use passed value
            newFenceDict["CustomLaunchEffect"] = customLaunchEffect; // Use passed value
            newFenceDict["IsHidden"] = false; // Use passed value
            newFenceDict["IsLocked"] = false; // Init ISLocked
                                              // Step 2
            newFenceDict["IsRolled"] = "false";
            newFenceDict["UnrolledHeight"] = 130;
            //Step 2


            if (itemsType == "Portal")
                if (itemsType == "Portal")
                {
                    using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                    {
                        dialog.Description = "Select the folder to monitor for this Portal Fence";
                        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            newFenceDict["Path"] = dialog.SelectedPath;
                        }
                        else
                        {
                            return;
                        }
                    }
                }
            _fenceData.Add(newFence);
            SaveFenceData();
            CreateFence(newFence, new TargetChecker(1000));
        }
        private static void MoveItem(dynamic item, dynamic sourceFence, Dispatcher dispatcher)
        {

            var moveWindow = new Window
            {
                Title = "Move To",
                Width = 250,
                Height = 300,
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };
            StackPanel sp = new StackPanel();
            moveWindow.Content = sp;

            foreach (var fence in _fenceData)
            {
                if (fence.ItemsType?.ToString() != "Portal")
                {
                    Button btn = new Button { Content = fence.Title.ToString(), Margin = new Thickness(5) };



                    btn.Click += (s, e) =>
                    {
                        var sourceItems = sourceFence.Items as JArray;
                        var destItems = fence.Items as JArray ?? new JArray();
                        if (sourceItems != null)
                        {
                            IDictionary<string, object> itemDict = item is IDictionary<string, object> dict ? dict : ((JObject)item).ToObject<IDictionary<string, object>>();
                            string filename = itemDict.ContainsKey("Filename") ? itemDict["Filename"].ToString() : "Unknown";

                            Log(LogLevel.Info, LogCategory.IconHandling, $"Moving item {filename} from {sourceFence.Title} to {fence.Title}");

                            // Find the JToken in sourceItems that matches the Filename
                            var itemToRemove = sourceItems.FirstOrDefault(i =>
                                i["Filename"]?.ToString() == filename
                            );

                            if (itemToRemove != null)
                            {
                                sourceItems.Remove(itemToRemove); // Remove the JToken from the JArray
                                destItems.Add(itemToRemove); // Add the JToken to the destination
                                fence.Items = destItems;
                                SaveFenceData();
                            }
                            else
                            {
                                Log(LogLevel.Warn, LogCategory.IconHandling, $"Item {filename} not found in source fence '{sourceFence.Title}'");
                            }

                            moveWindow.Close();
                            // ... rest of the code ...
                            var waitWindow = new Window
                            {
                                Title = "Desktop Fences +",
                                Width = 200,
                                Height = 100,
                                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                                WindowStyle = WindowStyle.None,
                                Background = System.Windows.Media.Brushes.LightGray,
                                Topmost = true
                            };
                            var waitStack = new StackPanel
                            {
                                HorizontalAlignment = HorizontalAlignment.Center,
                                VerticalAlignment = VerticalAlignment.Center
                            };
                            waitWindow.Content = waitStack;

                            string exePath = Assembly.GetEntryAssembly().Location;
                            var iconImage = new System.Windows.Controls.Image
                            {
                                Source = System.Drawing.Icon.ExtractAssociatedIcon(exePath).ToImageSource(),
                                Width = 32,
                                Height = 32,
                                Margin = new Thickness(0, 0, 0, 5)
                            };
                            waitStack.Children.Add(iconImage);

                            var waitLabel = new Label
                            {
                                Content = "Please wait...",
                                HorizontalAlignment = HorizontalAlignment.Center
                            };
                            waitStack.Children.Add(waitLabel);

                            waitWindow.Show();

                            dispatcher.InvokeAsync(() =>
                            {
                                if (Application.Current != null && !Application.Current.Dispatcher.HasShutdownStarted)
                                {
                                    Application.Current.Windows.OfType<NonActivatingWindow>().ToList().ForEach(w => w.Close());
                                    LoadAndCreateFences(new TargetChecker(1000));
                                    waitWindow.Close();
                                    Log(LogLevel.Info, LogCategory.IconHandling, $"Item moved successfully to {fence.Title}");
                                }
                                else
                                {
                                    waitWindow.Close();
                                    Log(LogLevel.Warn, LogCategory.General, $"Skipped fence reload due to application shutdown");
                                }
                            }, DispatcherPriority.Background);
                        }

                    };
                    sp.Children.Add(btn);
                }
            }
            moveWindow.ShowDialog();
        }
        private static WrapPanel FindWrapPanel(DependencyObject parent, int depth = 0, int maxDepth = 10)
        {
            // Prevent infinite recursion
            if (parent == null || depth > maxDepth)
            {
                Log(LogLevel.Debug, LogCategory.IconHandling, $"FindWrapPanel: Reached max depth {maxDepth} or null parent at depth {depth}");
                return null;
            }

            // Check if current element is a WrapPanel
            if (parent is WrapPanel wrapPanel)
            {
                Log(LogLevel.Debug, LogCategory.General, $"FindWrapPanel: Found WrapPanel at depth {depth}");
                return wrapPanel;
            }

            // Recurse through visual tree
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                Log(LogLevel.Debug, LogCategory.General, $"FindWrapPanel: Checking child {i} at depth {depth}, type: {child?.GetType()?.Name ?? "null"}");
                var result = FindWrapPanel(child, depth + 1, maxDepth);
                if (result != null)
                {
                    return result;
                }
            }

            Log(LogLevel.Debug, LogCategory.General, $"FindWrapPanel: No WrapPanel found under parent {parent?.GetType()?.Name ?? "null"} at depth {depth}");
            return null;
        }

        private static void EditItem(dynamic icon, dynamic fence, NonActivatingWindow win)
        {
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            string displayName = iconDict.ContainsKey("DisplayName") ? (string)iconDict["DisplayName"] : System.IO.Path.GetFileNameWithoutExtension(filePath);
            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";

            if (!isShortcut)
            {
                //   MessageBox.Show("Edit is only available for shortcuts.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);

                TrayManager.Instance.ShowOKOnlyMessageBoxForm("Edit is only available for shortcuts.", "Info");
                return;
            }

            var editWindow = new EditShortcutWindow(filePath, displayName);
            if (editWindow.ShowDialog() == true)
            {
                string newDisplayName = editWindow.NewDisplayName;
                iconDict["DisplayName"] = newDisplayName;
                Log(LogLevel.Debug, LogCategory.IconHandling, $"Updated DisplayName for {filePath} to {newDisplayName}");

                // Update fence data
                var items = fence.Items as JArray;
                if (items != null)
                {
                    var itemToUpdate = items.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                    if (itemToUpdate != null)
                    {
                        itemToUpdate["DisplayName"] = newDisplayName;
                        SaveFenceData();
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Fence data updated for {filePath}");
                    }
                    else
                    {
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Failed to find item {filePath} in fence items for update");
                    }
                }

                // Update UI
                if (win == null)
                {
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"Window is null for fence when updating {filePath}");
                    return;
                }

                // Attempt to find WrapPanel directly
                Log(LogLevel.Debug, LogCategory.IconHandling, $"Starting UI update for {filePath}. Window content type: {win.Content?.GetType()?.Name ?? "null"}");
                WrapPanel wpcont = null;
                var border = win.Content as Border;
                if (border != null)
                {
                    Log(LogLevel.Debug, LogCategory.General, $"Found Border. Child type: {border.Child?.GetType()?.Name ?? "null"}");
                    var dockPanel = border.Child as DockPanel;
                    if (dockPanel != null)
                    {
                        Log(LogLevel.Warn, LogCategory.General, $"Found DockPanel. Checking for ScrollViewer...");
                        var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                        if (scrollViewer != null)
                        {
                            Log(LogLevel.Debug, LogCategory.General, $"Found ScrollViewer. Content type: {scrollViewer.Content?.GetType()?.Name ?? "null"}");
                            wpcont = scrollViewer.Content as WrapPanel;
                        }
                        else
                        {
                            Log(LogLevel.Debug, LogCategory.General, $"ScrollViewer not found in DockPanel for {filePath}");
                        }
                    }
                    else
                    {
                        Log(LogLevel.Debug, LogCategory.General, $"DockPanel not found in Border for {filePath}");
                    }
                }
                else
                {
                    Log(LogLevel.Debug, LogCategory.General, $"Border not found in window content for {filePath}");
                }

                // Fallback to visual tree traversal if direct lookup fails
                if (wpcont == null)
                {
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"Direct WrapPanel lookup failed, attempting visual tree traversal for {filePath}");
                    wpcont = FindWrapPanel(win);
                }

                if (wpcont != null)
                {
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"Found WrapPanel. Checking for StackPanel with Tag.FilePath: {filePath}");
                    // Access Tag.FilePath to match ClickEventAdder's anonymous object
                    var sp = wpcont.Children.OfType<StackPanel>()
                        .FirstOrDefault(s => s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == filePath);
                    if (sp != null)
                    {
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Found StackPanel for {filePath}. Tag: {sp.Tag?.ToString() ?? "null"}");
                        var lbl = sp.Children.OfType<TextBlock>().FirstOrDefault();
                        if (lbl != null)
                        {
                            // Apply truncation logic (same as in AddIcon)
                            string displayText = string.IsNullOrEmpty(newDisplayName)
                                ? System.IO.Path.GetFileNameWithoutExtension(filePath)
                                : newDisplayName;
                            if (displayText.Length > 20)
                            {
                                displayText = displayText.Substring(0, 20) + "...";
                            }
                            lbl.Text = displayText;
                            lbl.InvalidateVisual(); // Force UI refresh
                            Log(LogLevel.Debug, LogCategory.IconHandling, $"Updated TextBlock for {filePath} to '{displayText}'");
                        }
                        else
                        {
                            Log(LogLevel.Debug, LogCategory.IconHandling, $"TextBlock not found in StackPanel for {filePath}. Children: {string.Join(", ", sp.Children.OfType<FrameworkElement>().Select(c => c.GetType().Name))}");
                        }
                        var ico = sp.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                        if (ico != null)
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                            ImageSource shortcutIcon = null;

                            if (!string.IsNullOrEmpty(shortcut.IconLocation))
                            {
                                string[] iconParts = shortcut.IconLocation.Split(',');
                                string iconPath = iconParts[0];
                                int iconIndex = 0;
                                if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex))
                                {
                                    iconIndex = parsedIndex; // Use the specified index
                                }

                                if (System.IO.File.Exists(iconPath))
                                {
                                    try
                                    {
                                        IntPtr[] hIcon = new IntPtr[1];
                                        uint result = ExtractIconEx(iconPath, iconIndex, hIcon, null, 1);
                                        if (result > 0 && hIcon[0] != IntPtr.Zero)
                                        {
                                            try
                                            {
                                                shortcutIcon = Imaging.CreateBitmapSourceFromHIcon(
                                                    hIcon[0],
                                                    Int32Rect.Empty,
                                                    BitmapSizeOptions.FromEmptyOptions()
                                                );
                                                Log(LogLevel.Debug, LogCategory.IconHandling, $"Extracted icon at index {iconIndex} from {iconPath} for {filePath}");
                                            }
                                            finally
                                            {
                                                DestroyIcon(hIcon[0]); // Clean up icon handle
                                            }
                                        }
                                        else
                                        {
                                            Log(LogLevel.Debug, LogCategory.IconHandling, $"Failed to extract icon at index {iconIndex} from {iconPath} for {filePath}. Result: {result}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log(LogLevel.Error, LogCategory.IconHandling, $"Error extracting icon from {iconPath} at index {iconIndex} for {filePath}: {ex.Message}");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, LogCategory.IconHandling, $"Icon file not found: {iconPath} for {filePath}");
                                }
                            }

                            // Fallback logic if no custom icon is extracted
                            if (shortcutIcon == null)
                            {
                                string targetPath = shortcut.TargetPath;
                                if (System.IO.File.Exists(targetPath))
                                {
                                    try
                                    {
                                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                                        Log(LogLevel.Debug, LogCategory.IconHandling, $"No custom icon, using target icon for {filePath}: {targetPath}");
                                    }
                                    catch (Exception ex)
                                    {
                                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Failed to extract target icon for {filePath}: {ex.Message}");
                                    }
                                }
                                else if (System.IO.Directory.Exists(targetPath))
                                {
                                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                                    Log(LogLevel.Debug, LogCategory.IconHandling, $"No custom icon, using folder-White.png for {filePath}");
                                }
                                else
                                {
                                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                                    Log(LogLevel.Debug, LogCategory.IconHandling, $"No custom icon or valid target, using file-WhiteX.png for {filePath}");
                                }
                            }

                            ico.Source = shortcutIcon;
                            iconCache[filePath] = shortcutIcon;
                            Log(LogLevel.Debug, LogCategory.IconHandling, $"Updated icon for {filePath}");
                        }


                        else
                        {
                            Log(LogLevel.Error, LogCategory.IconHandling, $"Image not found in StackPanel for {filePath}");
                        }
                    }
                    else
                    {
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"StackPanel not found for {filePath} in WrapPanel. Available Tag.FilePaths: {string.Join(", ", wpcont.Children.OfType<StackPanel>().Select(s => s.Tag != null ? (s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() ?? "null") : "null"))}");

                        // Fallback: Rebuild single icon
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Attempting to rebuild icon for {filePath}");
                        // Remove existing StackPanel if present (in case Tag lookup failed due to corruption)
                        var oldSp = wpcont.Children.OfType<StackPanel>()
                            .FirstOrDefault(s => s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == filePath);
                        if (oldSp != null)
                        {
                            wpcont.Children.Remove(oldSp);
                            Log(LogLevel.Debug, LogCategory.IconHandling, $"Removed old StackPanel for {filePath}");
                        }

                        // Re-add icon
                        AddIcon(icon, wpcont);
                        var newSp = wpcont.Children.OfType<StackPanel>().LastOrDefault();
                        if (newSp != null)
                        {
                            bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
                            string arguments = null;
                            if (isShortcut)
                            {
                                WshShell shell = new WshShell();
                                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                                arguments = shortcut.Arguments;
                            }
                            ClickEventAdder(newSp, filePath, isFolder, arguments);
                            Log(LogLevel.Debug, LogCategory.IconHandling, $"Rebuilt icon for {filePath} with new StackPanel");
                        }
                        else
                        {
                            Log(LogLevel.Debug, LogCategory.IconHandling, $"Failed to rebuild icon for {filePath}: No new StackPanel created");
                        }
                    }
                }
                else
                {
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"WrapPanel not found for {filePath}. UI update skipped.");
                }
            }
        }
        private static void BackupOrRestoreShortcut(string filePath, bool targetExists, bool isFolder)
        {


            string exeDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            string tempShortcutsDir = System.IO.Path.Combine(exeDir, "Temp Shortcuts");
            string backupFileName = System.IO.Path.GetFileName(filePath);
            string backupPath = System.IO.Path.Combine(tempShortcutsDir, backupFileName);

            try
            {
                // Ensure TempShortcuts directory exists
                if (!Directory.Exists(tempShortcutsDir))
                {
                    Directory.CreateDirectory(tempShortcutsDir);
                    Log(LogLevel.Debug, LogCategory.General, $"Created TempShortcuts directory: {tempShortcutsDir}");
                }

                if (!targetExists)
                {
                    // Backup shortcut if target is missing and not already backed up
                    if (System.IO.File.Exists(filePath) && !System.IO.File.Exists(backupPath))
                    {
                        System.IO.File.Copy(filePath, backupPath, true);
                        Log(LogLevel.Debug, LogCategory.General, $"Backed up shortcut {filePath} to {backupPath}");
                    }
                }
                else
                {
                    // Restore shortcut from backup if target exists
                    if (System.IO.File.Exists(backupPath))
                    {
                        // Verify the backup has a custom icon before restoring
                        WshShell shell = new WshShell();
                        IWshShortcut backupShortcut = (IWshShortcut)shell.CreateShortcut(backupPath);
                        if (!string.IsNullOrEmpty(backupShortcut.IconLocation))
                        {
                            System.IO.File.Copy(backupPath, filePath, true);
                            Log(LogLevel.Debug, LogCategory.General, $"Restored shortcut {filePath} from {backupPath} with custom icon");
                        }
                        // Delete backup
                        System.IO.File.Delete(backupPath);
                        Log(LogLevel.Debug, LogCategory.General, $"Deleted backup {backupPath} after restoration");
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, LogCategory.General, $"Error in BackupOrRestoreShortcut for {filePath}: {ex.Message}");
            }
        }

        private static void UpdateIcon(StackPanel sp, string filePath, bool isFolder)
        {


            if (Application.Current == null)
            {
                Log(LogLevel.Error, LogCategory.IconHandling, "Application.Current is null, cannot update icon.");
                return;
            }

            Application.Current.Dispatcher.Invoke(() =>

            {
                // Early return for web links - they should never have their icons updated by target checking
                // Find the fence item to check IsLink property
                bool isWebLink = false;
                try
                {
                    // Search through all fences to find this item and check IsLink
                    foreach (var fence in _fenceData)
                    {
                        if (fence.ItemsType?.ToString() == "Data")
                        {
                            var items = fence.Items as JArray;
                            if (items != null)
                            {
                                foreach (var item in items)
                                {
                                    string itemPath = item["Filename"]?.ToString();
                                    bool itemIsLink = item["IsLink"]?.ToObject<bool>() ?? false;
                                    if (itemPath == filePath && itemIsLink)
                                    {
                                        isWebLink = true;
                                        break;
                                    }
                                }
                            }
                        }
                        if (isWebLink) break;
                    }

                    if (isWebLink)
                    {
                        // do not remove the below commented log line, it may needed for future debugging
                     //   Log(LogLevel.Debug, LogCategory.IconHandling, $"Skipping icon update for web link: {filePath}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log(LogLevel.Warn, LogCategory.IconHandling, $"Error checking IsLink for {filePath}: {ex.Message}");
                }
                System.Windows.Controls.Image ico = sp.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                if (ico == null)
                {
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"No Image found in StackPanel for {filePath}");
                    return;
                }

                bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
                string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;
                bool targetExists = System.IO.File.Exists(targetPath) || System.IO.Directory.Exists(targetPath);
                bool isTargetFolder = System.IO.Directory.Exists(targetPath);

                // Correct isFolder for shortcut to folder
                if (isShortcut && isTargetFolder)
                {
                    isFolder = true;
                    // do not remove the below commented log line, it may needed for future debugging
                   // Log(LogLevel.Debug, LogCategory.IconHandling, $"Corrected isFolder to true for shortcut {filePath} targeting folder {targetPath}");
                }

                ImageSource newIcon = null;

                // Handle backup/restore for shortcuts
                if (isShortcut)
                {
                    BackupOrRestoreShortcut(filePath, targetExists, isFolder);
                }

                if (!targetExists)
                {
                    // Use missing icon for both shortcuts and non-shortcuts when target is missing
                    newIcon = isFolder
                        ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                        : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")}");
                }
                else if (isShortcut)
                {
                    WshShell shell = new WshShell();

                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);

                    // Check for custom IconLocation
                    if (!string.IsNullOrEmpty(shortcut.IconLocation))
                    {
                        string[] iconParts = shortcut.IconLocation.Split(',');
                        string iconPath = iconParts[0];
                        int iconIndex = 0;
                        if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex))
                        {
                            iconIndex = parsedIndex;
                        }

                        if (System.IO.File.Exists(iconPath))
                        {
                            try
                            {
                                IntPtr[] hIcon = new IntPtr[1];
                                uint result = ExtractIconEx(iconPath, iconIndex, hIcon, null, 1);
                                if (result > 0 && hIcon[0] != IntPtr.Zero)
                                {
                                    try
                                    {
                                        newIcon = Imaging.CreateBitmapSourceFromHIcon(
                                            hIcon[0],
                                            Int32Rect.Empty,
                                            BitmapSizeOptions.FromEmptyOptions()
                                        );
                                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Extracted custom icon at index {iconIndex} from {iconPath} for {filePath}");
                                    }
                                    finally
                                    {
                                        DestroyIcon(hIcon[0]);
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Debug, LogCategory.IconHandling, $"Failed to extract custom icon at index {iconIndex} from {iconPath} for {filePath}. Result: {result}");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Debug, LogCategory.IconHandling, $"Error extracting custom icon from {iconPath} at index {iconIndex} for {filePath}: {ex.Message}");
                            }
                        }
                        else
                        {
                            // do not remove the below commented log line, it may needed for future debugging
                            //Log(LogLevel.Debug, LogCategory.IconHandling, $"Custom icon file not found: {iconPath} for {filePath}");
                        }


                    }




                    // Fallback if no custom icon
                    if (newIcon == null)
                    {
                        if (isTargetFolder)
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                            // do not remove the below commented log line, it may needed for future debugging
                            //  Log(LogLevel.Debug, LogCategory.IconHandling, $"Using folder-White.png for shortcut {filePath} targeting folder {targetPath}");
                        }
                        else
                        {
                            try
                            {
                                newIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                                // do not remove the below commented log line, it may needed for future debugging
                                //  Log(LogLevel.Debug, LogCategory.IconHandling, $"Using target file icon for shortcut {filePath}: {targetPath}");
                            }
                            catch (Exception ex)
                            {
                                newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                                
                                Log(LogLevel.Error, LogCategory.IconHandling, $"Failed to extract target file icon for shortcut {filePath}: {ex.Message}");
                            }
                        }
                    }
                }
                else
                {
                    // Non-shortcut handling
                    if (isTargetFolder)
                    {
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Using folder-White.png for {filePath}");
                    }
                    else
                    {
                        try
                        {
                            newIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource();
                            // do not remove the below commented log line, it may needed for future debugging
                            // Log(LogLevel.Debug, LogCategory.IconHandling, $"Using file icon for {filePath}");
                        }
                        catch (Exception ex)
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                            Log(LogLevel.Error, LogCategory.IconHandling, $"Failed to extract icon for {filePath}: {ex.Message}");
                        }
                    }
                }

                // Update icon only if different
                if (ico.Source != newIcon)
                {
                    ico.Source = newIcon;
                    // Update cache if used
                    if (iconCache.ContainsKey(filePath))
                    {
                        iconCache[filePath] = newIcon;
                        Log(LogLevel.Debug, LogCategory.IconHandling, $"Updated icon cache for {filePath}");
                    }

                    // do not remove the below commented log line, it may needed for future debugging
                    // Log(LogLevel.Debug, LogCategory.IconHandling, $"Icon updated for {filePath}");
                }
                else
                {
                    Log(LogLevel.Debug, LogCategory.IconHandling, $"No icon update needed for {filePath}: same icon");
                }
            });
        }

        // Safety method to ensure no fences are stuck in transition state
        public static void ClearAllTransitionStates()
        {
            if (_fencesInTransition.Count > 0)
            {
                Log(LogLevel.Debug, LogCategory.FenceUpdate, $"Clearing {_fencesInTransition.Count} stuck transition states");
                _fencesInTransition.Clear();
            }
        }

        public static void UpdateOptionsAndClickEvents()
        {


            Log(LogLevel.Debug, LogCategory.Settings, $"Updating options, new singleClickToLaunch={SettingsManager.SingleClickToLaunch}");

            // Update _options
            _options = new
            {
                IsSnapEnabled = SettingsManager.IsSnapEnabled,
                TintValue = SettingsManager.TintValue,
                SelectedColor = SettingsManager.SelectedColor,
                IsLogEnabled = SettingsManager.IsLogEnabled,
                singleClickToLaunch = SettingsManager.SingleClickToLaunch,
                LaunchEffect = SettingsManager.LaunchEffect // Ensure this is here
            };


            if (Application.Current != null)
            {

                // Force UI update on the main thread
                Application.Current.Dispatcher.Invoke(() =>
            {
                int updatedItems = 0;
                foreach (var win in Application.Current.Windows.OfType<NonActivatingWindow>())
                {
                    var wpcont = ((win.Content as Border)?.Child as DockPanel)?.Children
                        .OfType<ScrollViewer>().FirstOrDefault()?.Content as WrapPanel;
                    if (wpcont != null)
                    {
                        foreach (var sp in wpcont.Children.OfType<StackPanel>())
                        {
                            string path = sp.Tag as string;
                            if (!string.IsNullOrEmpty(path))
                            {
                                bool isFolder = Directory.Exists(path) ||
                                    (System.IO.Path.GetExtension(path).ToLower() == ".lnk" &&
                                     Directory.Exists(Utility.GetShortcutTarget(path)));
                                string arguments = null;
                                if (System.IO.Path.GetExtension(path).ToLower() == ".lnk")
                                {
                                    WshShell shell = new WshShell();
                                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(path);
                                    arguments = shortcut.Arguments;
                                }
                                ClickEventAdder(sp, path, isFolder, arguments);
                                updatedItems++;
                            }
                        }
                    }
                }
                Log(LogLevel.Debug, LogCategory.Settings, $"Updated click events for {updatedItems} items");
            });
            }
            else
            {
                Log(LogLevel.Debug, LogCategory.Settings, "Application.Current is null, cannot update icon.");
            }
        }


        public static void ClickEventAdder(StackPanel sp, string path, bool isFolder, string arguments = null)
        {
            bool isLogEnabled = _options.IsLogEnabled ?? true;

            // Clear existing handler
            sp.MouseLeftButtonDown -= MouseDownHandler;

            // Store only path, isFolder, and arguments in Tag
            sp.Tag = new { FilePath = path, IsFolder = isFolder, Arguments = arguments };

            Log(LogLevel.Debug, LogCategory.General, $"Attaching handler for {path}, initial singleClickToLaunch={_options.singleClickToLaunch}");

            // Check if path is a shortcut and correct isFolder for folder shortcuts
            bool isShortcut = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
            string targetPath = isShortcut ? Utility.GetShortcutTarget(path) : path;
            if (isShortcut && System.IO.Directory.Exists(targetPath))
            {
                isFolder = true;
                Log(LogLevel.Debug, LogCategory.General, $"Corrected isFolder to true for shortcut {path} targeting folder {targetPath}");
            }

            void MouseDownHandler(object sender, MouseButtonEventArgs e)
            {
                if (e.ChangedButton != MouseButton.Left) return;

                bool singleClickToLaunch = _options.singleClickToLaunch ?? true;
                Log(LogLevel.Debug, LogCategory.General, $"MouseDown on {path}, ClickCount={e.ClickCount}, singleClickToLaunch={singleClickToLaunch}");

                try
                {
                    // Recompute isShortcut to avoid unassigned variable issue
                    bool isShortcutLocal = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
                    bool targetExists;
                    string resolvedPath = path;
                    if (isShortcutLocal)
                    {
                        resolvedPath = Utility.GetShortcutTarget(path);
                        targetExists = isFolder ? System.IO.Directory.Exists(resolvedPath) : System.IO.File.Exists(resolvedPath);
                    }
                    else
                    {
                        targetExists = isFolder ? System.IO.Directory.Exists(path) : System.IO.File.Exists(path);
                    }

                    Log(LogLevel.Debug, LogCategory.General, $"Target check: Path={path}, ResolvedPath={resolvedPath}, IsShortcut={isShortcutLocal}, IsFolder={isFolder}, TargetExists={targetExists}");

                    if (!targetExists)
                    {
                        Log(LogLevel.Warn, LogCategory.General, $"Target not found: {resolvedPath}");
                        return;
                    }

                    if (singleClickToLaunch && e.ClickCount == 1)
                    {
                        Log(LogLevel.Debug, LogCategory.General, $"Single click launching {path}");
                        LaunchItem(sp, path, isFolder, arguments);
                        e.Handled = true;
                    }
                    else if (!singleClickToLaunch && e.ClickCount == 2)
                    {
                        Log(LogLevel.Debug, LogCategory.General, $"Double click launching {path}");
                        LaunchItem(sp, path, isFolder, arguments);
                        e.Handled = true;
                    }
                }
                catch (Exception ex)
                {
                    Log(LogLevel.Error, LogCategory.General, $"Error checking target existence: {ex.Message}");
                }
            }

            sp.MouseLeftButtonDown += MouseDownHandler;
        }




        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);
            while (parent != null && !(parent is T))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            return parent as T;
        }


        private static void LaunchItem(StackPanel sp, string path, bool isFolder, string arguments)
        {


            try
            {
                // Find the fence at runtime
                NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(sp);
                dynamic fence = _fenceData.FirstOrDefault(f => f.Title == win?.Title);
                string customEffect = fence?.CustomLaunchEffect?.ToString();
                LaunchEffect defaultEffect = _options.LaunchEffect;
                LaunchEffect effect;

                if (fence == null)
                {
                    Log(LogLevel.Debug, LogCategory.General, $"Failed to find fence for {path}, using default effect");
                    effect = defaultEffect;
                }
                else if (string.IsNullOrEmpty(customEffect))
                {
                    effect = defaultEffect;
                    Log(LogLevel.Debug, LogCategory.General, $"No CustomLaunchEffect for {path} in fence '{fence.Title}', using default: {effect}");
                }
                else
                {
                    try
                    {
                        effect = (LaunchEffect)Enum.Parse(typeof(LaunchEffect), customEffect, true);
                        Log(LogLevel.Debug, LogCategory.General, $"Using CustomLaunchEffect for {path} in fence '{fence.Title}': {effect}");
                    }
                    catch (Exception ex)
                    {
                        effect = defaultEffect;
                        Log(LogLevel.Debug, LogCategory.General, $"Failed to parse CustomLaunchEffect '{customEffect}' for {path} in fence '{fence.Title}', falling back to {effect}: {ex.Message}");
                    }
                }

                Log(LogLevel.Debug, LogCategory.General, $"LaunchItem called for {path} with effect {effect}");

                // Ensure transform is set up
                if (sp.RenderTransform == null || !(sp.RenderTransform is TransformGroup))
                {
                    sp.RenderTransform = new TransformGroup
                    {
                        Children = new TransformCollection
                {
                    new ScaleTransform(1, 1),
                    new TranslateTransform(0, 0),
                    new RotateTransform(0)
                }
                    };
                    sp.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
                }
                var transformGroup = (TransformGroup)sp.RenderTransform;
                var scaleTransform = (ScaleTransform)transformGroup.Children[0];
                var translateTransform = (TranslateTransform)transformGroup.Children[1];
                var rotateTransform = (RotateTransform)transformGroup.Children[2];

                // Define animation based on selected effect
                switch (effect)
                {
                    case LaunchEffect.Zoom:
                        var zoomScale = new DoubleAnimation(1, 1.2, TimeSpan.FromSeconds(0.1)) { AutoReverse = true };
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, zoomScale);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, zoomScale);
                        break;

                    case LaunchEffect.Bounce:
                        var bounceScale = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.6)
                        };
                        bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                        bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.1))));
                        bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
                        bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.2, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
                        bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
                        bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5))));
                        bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, bounceScale);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, bounceScale);
                        break;

                    case LaunchEffect.FadeOut:
                        var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.2)) { AutoReverse = true };
                        sp.BeginAnimation(UIElement.OpacityProperty, fade);
                        break;

                    case LaunchEffect.SlideUp:
                        var slideUp = new DoubleAnimation(0, -20, TimeSpan.FromSeconds(0.2)) { AutoReverse = true };
                        translateTransform.BeginAnimation(TranslateTransform.YProperty, slideUp);
                        break;

                    case LaunchEffect.Rotate:
                        var rotate = new DoubleAnimation(0, 360, TimeSpan.FromSeconds(0.4))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };
                        rotateTransform.BeginAnimation(RotateTransform.AngleProperty, rotate);
                        break;


                    case LaunchEffect.Agitate:
                        var agitateTranslate = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.7)
                        };
                        agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                        agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(-10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.1))));
                        agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
                        agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(-10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
                        agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
                        agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(-10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5))));
                        agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
                        agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.7))));
                        translateTransform.BeginAnimation(TranslateTransform.XProperty, agitateTranslate);
                        break;


                    case LaunchEffect.GrowAndFly:
                        var growAnimation = new DoubleAnimation(1, 1.2, TimeSpan.FromSeconds(0.2));
                        growAnimation.Completed += (s, _) =>
                        {
                            // After growing, start the fly away animation
                            var shrinkAnimation = new DoubleAnimation(1.2, 0.05, TimeSpan.FromSeconds(0.3));
                            var moveUpAnimation = new DoubleAnimation(0, -50, TimeSpan.FromSeconds(0.3));

                            shrinkAnimation.Completed += (s2, _2) =>
                            {
                                // Make the icon invisible
                                sp.Opacity = 0;

                                // Remove all animations to allow direct property setting
                                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, null);
                                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, null);
                                translateTransform.BeginAnimation(TranslateTransform.YProperty, null);

                                // Reset transform values
                                scaleTransform.ScaleX = 1.0;
                                scaleTransform.ScaleY = 1.0;
                                translateTransform.Y = 0;

                                // Small delay before showing the icon again0
                                var restoreTimer = new System.Windows.Threading.DispatcherTimer
                                {
                                    Interval = TimeSpan.FromSeconds(0.1)
                                };

                                restoreTimer.Tick += (timerSender, timerArgs) =>
                                {
                                    // Restore opacity
                                    var restoreAnimation = new DoubleAnimation(0, 1, TimeSpan.FromSeconds(0.1));
                                    sp.BeginAnimation(UIElement.OpacityProperty, restoreAnimation);

                                    // Stop and cleanup timer
                                    restoreTimer.Stop();
                                };

                                restoreTimer.Start();
                            };

                            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, shrinkAnimation);
                            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, shrinkAnimation);
                            translateTransform.BeginAnimation(TranslateTransform.YProperty, moveUpAnimation);
                        };

                        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, growAnimation);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, growAnimation);
                        break;



                    case LaunchEffect.Pulse:
                        // Creates a pulsing effect with color change
                        var pulseAnimation = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.8),
                            AutoReverse = false
                        };
                        pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                        pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
                        pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0.8, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
                        pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
                        pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.8))));

                        // Optional: Add color animation if icon supports it (assuming it's a Path or Shape)
                        if (sp.Children.Count > 0 && sp.Children[0] is Shape shape)
                        {
                            var originalBrush = shape.Fill as SolidColorBrush;
                            if (originalBrush != null)
                            {
                                var colorAnimation = new ColorAnimation(
                                    Colors.Red,
                                    TimeSpan.FromSeconds(0.4))
                                {
                                    AutoReverse = true
                                };
                                shape.Fill.BeginAnimation(SolidColorBrush.ColorProperty, colorAnimation);
                            }
                        }

                        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, pulseAnimation);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, pulseAnimation);
                        break;

                    case LaunchEffect.Spiral:
                        // Combines rotation with a zoom effect
                        var spiralRotate = new DoubleAnimation(0, 720, TimeSpan.FromSeconds(0.7))
                        {
                            EasingFunction = new BackEase { Amplitude = 0.3, EasingMode = EasingMode.EaseInOut }
                        };

                        var spiralScale = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.7)
                        };
                        spiralScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                        spiralScale.KeyFrames.Add(new EasingDoubleKeyFrame(0.7, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
                        spiralScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5))));
                        spiralScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.7))));

                        rotateTransform.BeginAnimation(RotateTransform.AngleProperty, spiralRotate);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, spiralScale);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, spiralScale);
                        break;

                    case LaunchEffect.Elastic:
                        // Creates a stretchy, elastic effect
                        var elasticX = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.8)
                        };
                        elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                        elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(1.5, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
                        elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(0.8, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
                        elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(1.1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
                        elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.8))));

                        var elasticY = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.8)
                        };
                        elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                        elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(0.7, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
                        elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(1.2, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
                        elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(0.9, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
                        elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.8))));

                        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, elasticX);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, elasticY);
                        break;

                    case LaunchEffect.Flip3D:
                        // Creates a 3D flip effect
                        var flipAnimation = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.6)
                        };

                        flipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                        flipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(90, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.15))));
                        flipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(270, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.45))));
                        flipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(360, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));

                        // For X-axis flip
                        scaleTransform.CenterX = sp.ActualWidth / 2;
                        scaleTransform.CenterY = sp.ActualHeight / 2;

                        // We use a scale animation to create the flip illusion
                        var scaleFlipAnimation = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.6)
                        };
                        scaleFlipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                        scaleFlipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.15))));
                        scaleFlipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.45))));
                        scaleFlipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));

                        rotateTransform.BeginAnimation(RotateTransform.AngleProperty, flipAnimation);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleFlipAnimation);
                        break;

                        //case LaunchEffect.Explosion:



                }

                // Execution
                bool isShortcut = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
                string targetPath = isShortcut ? Utility.GetShortcutTarget(path) : path;
                bool isTargetFolder = System.IO.Directory.Exists(targetPath);
                bool targetExists = System.IO.File.Exists(targetPath) || System.IO.Directory.Exists(targetPath);

                Log(LogLevel.Debug, LogCategory.General, $"Target path resolved to: {targetPath}");
                Log(LogLevel.Debug, LogCategory.General, $"Target exists: {targetExists}, IsFolder: {isTargetFolder}");

                if (targetExists)
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = targetPath,
                        UseShellExecute = true,
                        Verb = isTargetFolder ? "open" : null
                    };
                    if (!string.IsNullOrEmpty(arguments))
                    {
                        psi.Arguments = arguments;
                        Log(LogLevel.Debug, LogCategory.General, $"Arguments: {arguments}");
                    }

                    Log(LogLevel.Debug, LogCategory.General, $"Attempting to launch {targetPath}");
                    Process.Start(psi);
                    Log(LogLevel.Debug, LogCategory.General, $"Successfully launched {targetPath}");
                }
                else
                {
                    Log(LogLevel.Error, LogCategory.General, $"Target not found: {targetPath}");
                    //    MessageBox.Show($"Target '{targetPath}' was not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Target '{targetPath}' was not found.", "Error");

                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, LogCategory.General, $"Error in LaunchItem for {path}: {ex.Message}");
                //  MessageBox.Show($"Error opening item: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Error opening item: {ex.Message}", "Error");
            }
        }
        static readonly string[] adjectives = {
        "High", "Low", "Tiny", "Vast", "Wide", "Slim", "Flat", "Bold", "Cold", "Warm",
        "Soft", "Hard", "Dark", "Pale", "Fast", "Slow", "Deep", "Tall", "Short", "Bent",
        "Thin", "Bright", "Light", "Sharp", "Dull", "Loud", "Mute", "Grim", "Kind", "Neat",
        "Rough", "Smooth", "Brave", "Fierce", "Plain", "Worn", "Dry", "Damp", "Strong", "Weak"
    };

        static readonly string[] places = {
        "Bay", "Hill", "Lake", "Cove", "Peak", "Reef", "Dune", "Glen", "Moor", "Vale",
        "Rock", "Shore", "Bank", "Ford", "Cape", "Crag", "Marsh", "Pond", "Cliff", "Wood",
        "Dell", "Pass", "Cave", "Ridge", "Knob", "Fall", "Isle", "Path", "Stream", "Creek",
        "Field", "Plain", "Bluff", "Point", "Grove", "Dock", "Harbor", "Spring", "Meadow", "Hollow"
    };

        static readonly Random random = new Random();

        public static string GenerateRandomName()
        {
            string word1 = adjectives[random.Next(adjectives.Length)];
            string word2 = places[random.Next(places.Length)];
            return word1 + " " + word2;
        }


    }
}