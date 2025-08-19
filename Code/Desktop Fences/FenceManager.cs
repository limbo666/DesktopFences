using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics.Arm;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using IWshRuntimeLibrary;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Desktop_Fences
{
    public static class FenceManager
    {

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern uint ExtractIconEx(string szFileName, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

        [DllImport("user32.dll")]

        private static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        private static List<dynamic> _fenceData;
        private static string _jsonFilePath;
        private static dynamic _options;
        private static readonly Dictionary<string, ImageSource> iconCache = new Dictionary<string, ImageSource>();
        private static Dictionary<dynamic, PortalFenceManager> _portalFences = new Dictionary<dynamic, PortalFenceManager>();

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
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error reloading fences: {ex.Message}");
                throw;
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
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Adjusted fence '{win.Title}' position to ({newLeft}, {newTop}) to fit within screen bounds.");
            }
            SaveFenceData();
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
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Creating new fence at mouse position: X={mousePosition.X}, Y={mousePosition.Y}");
                //CreateNewFence("New Fence", "Data", mousePosition.X, mousePosition.Y);
                CreateNewFence("", "Data", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newFenceItem);

            // New Portal Fence item
            var newPortalFenceItem = new MenuItem { Header = "New Portal Fence" };
            newPortalFenceItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Creating new portal fence at mouse position: X={mousePosition.X}, Y={mousePosition.Y}");
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
                    // NEW: Use BackupManager to handle the deletion backup instead of manual backup code
                    BackupManager.BackupDeletedFence(fence);

                    // Proceed with deletion - remove from data structures
                    _fenceData.Remove(fence);
                    _heartTextBlocks.Remove(fence); // Remove from heart TextBlocks dictionary

                    // Clean up portal fence if applicable
                    if (_portalFences.ContainsKey(fence))
                    {
                        _portalFences[fence].Dispose();
                        _portalFences.Remove(fence);
                    }

                    // Save updated fence data
                    SaveFenceData();

                    // Find and close the window - need to find the window associated with this fence
                    var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                    var win = windows.FirstOrDefault(w => w.Tag?.ToString() == fence.Id?.ToString());
                    if (win != null)
                    {
                        win.Close();
                    }

                    // Update all heart context menus to show restore option is now available
                    UpdateAllHeartContextMenus();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation, $"Fence '{fence.Title}' deleted successfully with backup created");
                }
            };

            menu.Items.Add(deleteThisFence);



            // Restore Fence item
            var restoreItem = new MenuItem
            {
                Header = "Restore Last Deleted Fence",
                // Visibility = _isRestoreAvailable ? Visibility.Visible : Visibility.Collapsed
                Visibility = BackupManager.IsRestoreAvailable ? Visibility.Visible : Visibility.Collapsed
            };
            //  restoreItem.Click += (s, e) => RestoreLastDeletedFence();
            restoreItem.Click += (s, e) => BackupManager.RestoreLastDeletedFence();

            menu.Items.Add(restoreItem);

            menu.Items.Add(new Separator());
            // New Export Fence menu item
            var exportItem = new MenuItem { Header = "Export this Fence" };
            // exportItem.Click += (s, e) => ExportFence(fence);
            exportItem.Click += (s, e) => BackupManager.ExportFence(fence);
            menu.Items.Add(exportItem);

            var importItem = new MenuItem { Header = "Import a Fence..." };
            //importItem.Click += (s, e) => ImportFence();
            importItem.Click += (s, e) => BackupManager.ImportFence();
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
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Updated heart ContextMenu for fence '{fence.Title}'");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Skipped update for fence '{fence.Title}': heart TextBlock is null");
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
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Error checking if {filePath} is web link: {ex.Message}");
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
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Error extracting web URL from {filePath}: {ex.Message}");
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

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Created web link URL file: {urlFilePath} -> {targetUrl}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, $"Error creating web link shortcut {shortcutPath}: {ex.Message}");
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
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Shortcut {filePath} targets {targetPath}, IsUNC: {isUncPath}");
                            return isUncPath;
                        }
                    }
                }
                else
                {
                    // For direct paths, check if it's UNC
                    bool isUncPath = filePath.StartsWith("\\\\");
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Direct path {filePath}, IsUNC: {isUncPath}");
                    return isUncPath;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Error checking if {filePath} is network path: {ex.Message}");
            }
            return false;
        }



        public static List<dynamic> GetFenceData()
        {
            return _fenceData;
        }



        // Update fence property, save to JSON, and apply runtime changes

        public static void UpdateFenceProperty(dynamic fence, string propertyName, string value, string logMessage)
        {
            try
            {


                string fenceId = fence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Fence '{fence.Title}' has no Id");
                    return;
                }

                // Skip updates if fence is in transition (except for IsRolled and UnrolledHeight which are rollup-specific)
                if (_fencesInTransition.Contains(fenceId) && propertyName != "IsRolled" && propertyName != "UnrolledHeight")
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate, $"Skipping {propertyName} update for fence '{fenceId}' - in transition");
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
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"{logMessage} for fence '{fence.Title}'");

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
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Applied color '{value ?? "Default"}' to fence '{fence.Title}' at runtime");
                        }
                        else if (propertyName == "IsHidden")
                        {
                            // Update visibility based on IsHidden
                            bool isHidden = value?.ToLower() == "true";
                            win.Visibility = isHidden ? Visibility.Hidden : Visibility.Visible;
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Set visibility to {(isHidden ? "Hidden" : "Visible")} for fence '{fence.Title}'");
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
                                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Set WrapPanel visibility to {(isRolled ? "Collapsed" : "Visible")} for fence '{actualFence.Title}'");
                                        }
                                    }
                                }
                            }
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Set height to {targetHeight} for fence '{actualFence.Title}' (IsRolled={isRolled})");
                        }
                        else if (propertyName == "UnrolledHeight")
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Set UnrolledHeight to {value} for fence '{actualFence.Title}'");
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
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate, $"Failed to find window for fence '{fence.Title}' to apply {propertyName}");
                    }
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate, $"Failed to find fence '{fence.Title}' in _fenceData for {propertyName} update");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate, $"Error updating {propertyName} for fence '{fence.Title}': {ex.Message}");
            }
        }


       
        public static void LoadAndCreateFences(TargetChecker targetChecker)
        {
            string exePath = Assembly.GetEntryAssembly().Location;
            string exeDir = System.IO.Path.GetDirectoryName(exePath);
            _jsonFilePath = System.IO.Path.Combine(exeDir, "fences.json");
            // Below added for reload function
            _currentTargetChecker = targetChecker;

            SettingsManager.LoadSettings();

            // Delete previous log file if setting is enabled
            if (SettingsManager.DeletePreviousLogOnStart)
            {
                try
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                    if (System.IO.File.Exists(logPath))
                    {
                        System.IO.File.Delete(logPath);
                        // Note: Can't log this deletion since we just deleted the log file
                    }
                }
                catch
                {
                    // Silently ignore deletion errors to avoid startup issues
                }
            }


            BackupManager.CleanLastDeletedFolder();
            InterCore.Initialize();

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

            bool jsonLoadSuccessful = false;

            if (System.IO.File.Exists(_jsonFilePath))
            {
                jsonLoadSuccessful = LoadFenceDataFromJson();
            }

            // If JSON loading failed or file doesn't exist, initialize with defaults
            if (!jsonLoadSuccessful || _fenceData == null || _fenceData.Count == 0)
            {
                InitializeDefaultFence();
            }
            else
            {
                // Only migrate if we successfully loaded existing data
                MigrateLegacyJson();
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
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation, $"Marked Portal Fence '{fence.Title}' for removal due to missing target folder: {targetPath ?? "null"}");
                    }
                }
            }

            // Remove invalid fences and save
            if (invalidFences.Any())
            {
                foreach (var fence in invalidFences)
                {
                    _fenceData.Remove(fence);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Removed Portal Fence '{fence.Title}' from _fenceData");
                }
                SaveFenceData();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Saved updated fences.json after removing {invalidFences.Count} invalid Portal Fences");
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
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate, $"Emergency cleanup: Found {_fencesInTransition.Count} fences stuck in transition state");
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

        private static bool LoadFenceDataFromJson()
        {
            try
            {
                string jsonContent = System.IO.File.ReadAllText(_jsonFilePath);

                // Check if the file is empty or contains only whitespace
                if (string.IsNullOrWhiteSpace(jsonContent))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                        "JSON file is empty or contains only whitespace. Using default fence configuration.");
                    return false;
                }

                // First, try to parse as a list of fences
                try
                {
                    _fenceData = JsonConvert.DeserializeObject<List<dynamic>>(jsonContent);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                        $"Successfully loaded {_fenceData?.Count ?? 0} fences from JSON array.");
                    return _fenceData != null;
                }
                catch (JsonSerializationException)
                {
                    // If that fails, try to parse as a single fence object
                    try
                    {
                        var singleFence = JsonConvert.DeserializeObject<dynamic>(jsonContent);
                        if (singleFence != null)
                        {
                            _fenceData = new List<dynamic> { singleFence };
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                "Successfully loaded single fence from JSON object.");
                            return true;
                        }
                    }
                    catch (JsonSerializationException innerEx)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                            $"Failed to parse JSON as single fence object: {innerEx.Message}");
                    }
                }
            }
            catch (System.IO.IOException ioEx)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"IO error reading fences.json: {ioEx.Message}");
            }
            catch (UnauthorizedAccessException accessEx)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Access denied reading fences.json: {accessEx.Message}");
            }
            catch (JsonReaderException jsonEx)
            {
                // Handle malformed JSON (syntax errors, invalid characters, etc.)
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Malformed JSON detected in fences.json: {jsonEx.Message}");

                // Optionally, create a backup of the corrupted file
                CreateCorruptedFileBackup();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Unexpected error loading fences.json: {ex.Message}");
            }

            return false;
        }

        private static void CreateCorruptedFileBackup()
        {
            try
            {
                string backupPath = _jsonFilePath + ".corrupted." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                System.IO.File.Copy(_jsonFilePath, backupPath, true);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    $"Created backup of corrupted JSON file: {backupPath}");
            }
            catch (Exception backupEx)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Failed to create backup of corrupted JSON file: {backupEx.Message}");
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
                // var validEffects = Enum.GetNames(typeof(LaunchEffect)).ToHashSet();
                var validEffects = Enum.GetNames(typeof(LaunchEffectsManager.LaunchEffect)).ToHashSet();
                for (int i = 0; i < _fenceData.Count; i++)
                {
                    dynamic fence = _fenceData[i];
                    IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();

                    // Add GUID if missing
                    if (!fenceDict.ContainsKey("Id"))
                    {
                        fenceDict["Id"] = Guid.NewGuid().ToString();
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added Id to {fence.Title}");
                    }

                    // Existing migration: Handle Portal Fence ItemsType
                    if (fence.ItemsType?.ToString() == "Portal")
                    {
                        string portalPath = fence.Items?.ToString();
                        if (!string.IsNullOrEmpty(portalPath) && !System.IO.Directory.Exists(portalPath))
                        {
                            fenceDict["IsFolder"] = true;
                            jsonModified = true;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Migrated legacy portal fence {fence.Title}: Added IsFolder=true");
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
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Migrated item in {fence.Title}: Added IsFolder for {path}");
                            }
                        }
                        fenceDict["Items"] = items; // Ensure Items updates are captured
                    }
                    // Migration: Add or validate IsLocked
                    if (!fenceDict.ContainsKey("IsLocked"))
                    {
                        fenceDict["IsLocked"] = "false"; // Default to string "false"
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added IsLocked=\"false\" to {fence.Title}");
                    }



                    // Migration: Add or validate IsRolled
                    if (!fenceDict.ContainsKey("IsRolled"))
                    {
                        fenceDict["IsRolled"] = "false";
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added IsRolled=\"false\" to {fence.Title}");
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
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid IsRolled value '{fenceDict["IsRolled"]}' in {fence.Title}, resetting to \"false\"");
                            isRolled = false;
                            jsonModified = true;
                        }
                        fenceDict["IsRolled"] = isRolled.ToString().ToLower();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Preserved IsRolled: \"{isRolled.ToString().ToLower()}\" for {fence.Title}");
                    }

                    // Migration: Add or validate IconSize
                    if (!fenceDict.ContainsKey("IconSize"))
                    {
                        fenceDict["IconSize"] = "Medium"; // Default to current size
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added IconSize=\"Medium\" to {fence.Title}");
                    }
                    else
                    {
                        string iconSize = fenceDict["IconSize"]?.ToString();
                        var validSizes = new[] { "Tiny","Small", "Medium", "Large", "Huge" };
                        if (string.IsNullOrEmpty(iconSize) || !validSizes.Contains(iconSize))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid IconSize '{iconSize}' in {fence.Title}, resetting to \"Medium\"");
                            fenceDict["IconSize"] = "Medium";
                            jsonModified = true;
                        }
                    }


                    // Migration: Add or validate UnrolledHeight
                    if (!fenceDict.ContainsKey("UnrolledHeight"))
                    {
                        double height = fenceDict.ContainsKey("Height") ? Convert.ToDouble(fenceDict["Height"]) : 130;
                        fenceDict["UnrolledHeight"] = height;
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added UnrolledHeight={height} to {fence.Title}");
                    }
                         else
                    {
                        double unrolledHeight;
                        if (!double.TryParse(fenceDict["UnrolledHeight"]?.ToString(), out unrolledHeight) || unrolledHeight <= 0)
                        {
                            unrolledHeight = fenceDict.ContainsKey("Height") ? Convert.ToDouble(fenceDict["Height"]) : 130;
                            fenceDict["UnrolledHeight"] = unrolledHeight;
                            jsonModified = true;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid UnrolledHeight '{fenceDict["UnrolledHeight"]}' in {fence.Title}, reset to {unrolledHeight}");
                        }
                    }


                    // Migration: Add or validate CustomColor
                    if (!fenceDict.ContainsKey("CustomColor"))
                    {
                        fenceDict["CustomColor"] = null;
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added CustomColor=null to {fence.Title}");
                    }
                    else
                    {
                        string customColor = fenceDict["CustomColor"]?.ToString();
                        if (!string.IsNullOrEmpty(customColor) && !validColors.Contains(customColor))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid CustomColor '{customColor}' in {fence.Title}, resetting to null");
                            fenceDict["CustomColor"] = null;
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate FenceBorderColor
                    if (!fenceDict.ContainsKey("FenceBorderColor"))
                    {
                        fenceDict["FenceBorderColor"] = null; // Default no border color
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added FenceBorderColor=null to {fence.Title}");
                    }
                    else
                    {
                        // Validate FenceBorderColor value
                        string borderColor = fenceDict["FenceBorderColor"]?.ToString();
                        if (!string.IsNullOrEmpty(borderColor) && !validColors.Contains(borderColor))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid FenceBorderColor '{borderColor}' in {fence.Title}, resetting to null");
                            fenceDict["FenceBorderColor"] = null;
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate FenceBorderThickness
                    if (!fenceDict.ContainsKey("FenceBorderThickness"))
                    {
                        fenceDict["FenceBorderThickness"] = 0; // Default border thickness
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added FenceBorderThickness=0 to {fence.Title}");
                    }
                    else
                    {
                        // Validate FenceBorderThickness value (0-5 pixels reasonable range)
                        int borderThickness;
                        if (!int.TryParse(fenceDict["FenceBorderThickness"]?.ToString(), out borderThickness) || borderThickness < 0 || borderThickness > 5)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid FenceBorderThickness '{fenceDict["FenceBorderThickness"]}' in {fence.Title}, resetting to 0");
                            fenceDict["FenceBorderThickness"] = 0;
                            jsonModified = true;
                        }
                    }





                    // Migration: Add or validate BoldTitleText
                    if (!fenceDict.ContainsKey("BoldTitleText"))
                    {
                        fenceDict["BoldTitleText"] = "false"; // Default to normal text
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added BoldTitleText=\"false\" to {fence.Title}");
                    }
                    else
                    {
                        // Validate BoldTitleText value
                        bool isBold = false;
                        if (fenceDict["BoldTitleText"] is bool boolValue)
                        {
                            isBold = boolValue;
                        }
                        else if (fenceDict["BoldTitleText"] is string stringValue)
                        {
                            isBold = stringValue.ToLower() == "true";
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid BoldTitleText value '{fenceDict["BoldTitleText"]}' in {fence.Title}, resetting to \"false\"");
                            isBold = false;
                            jsonModified = true;
                        }
                        fenceDict["BoldTitleText"] = isBold.ToString().ToLower();
                    }


                    // Migration: Add or validate TitleTextColor
                    if (!fenceDict.ContainsKey("TitleTextColor"))
                    {
                        fenceDict["TitleTextColor"] = null;
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added TitleTextColor=null to {fence.Title}");
                    }
                    else
                    {
                        string titleTextColor = fenceDict["TitleTextColor"]?.ToString();
                        if (!string.IsNullOrEmpty(titleTextColor) && !validColors.Contains(titleTextColor))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid TitleTextColor '{titleTextColor}' in {fence.Title}, resetting to null");
                            fenceDict["TitleTextColor"] = null;
                            jsonModified = true;
                        }
                    }


                    // Migration: Add or validate DisableTextShadow
                    if (!fenceDict.ContainsKey("DisableTextShadow"))
                    {
                        fenceDict["DisableTextShadow"] = "false"; // Default to shadow enabled
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added DisableTextShadow=\"false\" to {fence.Title}");
                    }
                    else
                    {
                        // Validate DisableTextShadow value
                        bool isDisabled = false;
                        if (fenceDict["DisableTextShadow"] is bool boolValue)
                        {
                            isDisabled = boolValue;
                        }
                        else if (fenceDict["DisableTextShadow"] is string stringValue)
                        {
                            isDisabled = stringValue.ToLower() == "true";
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid DisableTextShadow value '{fenceDict["DisableTextShadow"]}' in {fence.Title}, resetting to \"false\"");
                            isDisabled = false;
                            jsonModified = true;
                        }
                        fenceDict["DisableTextShadow"] = isDisabled.ToString().ToLower();
                    }


                    // Migration: Add or validate GrayscaleIcons
                    if (!fenceDict.ContainsKey("GrayscaleIcons"))
                    {
                        fenceDict["GrayscaleIcons"] = "false";
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added GrayscaleIcons=\"false\" to {fence.Title}");
                    }
                    else
                    {
                        bool isGrayscale = false;
                        if (fenceDict["GrayscaleIcons"] is bool boolValue)
                        {
                            isGrayscale = boolValue;
                        }
                        else if (fenceDict["GrayscaleIcons"] is string stringValue)
                        {
                            isGrayscale = stringValue.ToLower() == "true";
                        }
                        else
                        {
                            isGrayscale = false;
                            jsonModified = true;
                        }
                        fenceDict["GrayscaleIcons"] = isGrayscale.ToString().ToLower();
                    }

                    // Migration: Add or validate TitleTextSize
                    if (!fenceDict.ContainsKey("TitleTextSize"))
                    {
                        fenceDict["TitleTextSize"] = "Medium"; // Default title text size
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added TitleTextSize=\"Medium\" to {fence.Title}");
                    }
                    else
                    {
                        // Validate TitleTextSize value
                        string titleTextSize = fenceDict["TitleTextSize"]?.ToString();
                        var validSizes = new[] { "Small", "Medium", "Large" };
                        if (string.IsNullOrEmpty(titleTextSize) || !validSizes.Contains(titleTextSize))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid TitleTextSize '{titleTextSize}' in {fence.Title}, resetting to \"Medium\"");
                            fenceDict["TitleTextSize"] = "Medium";
                            jsonModified = true;
                        }
                    }




                    // Migration: Add or validate IconSpacing
                    if (!fenceDict.ContainsKey("IconSpacing"))
                    {
                        fenceDict["IconSpacing"] = 5; // Default spacing between icons
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added IconSpacing=5 to {fence.Title}");
                    }
                    else
                    {
                        // Validate existing IconSpacing value (0-20 pixels reasonable range)
                        int iconSpacing;
                        if (!int.TryParse(fenceDict["IconSpacing"]?.ToString(), out iconSpacing) || iconSpacing < 0 || iconSpacing > 20)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid IconSpacing '{fenceDict["IconSpacing"]}' in {fence.Title}, resetting to 5");
                            fenceDict["IconSpacing"] = 5;
                            jsonModified = true;
                        }
                        // If valid, keep the existing value unchanged
                    }

                    // Migration: Add or validate CustomLaunchEffect
                    if (!fenceDict.ContainsKey("CustomLaunchEffect"))
                    {
                        fenceDict["CustomLaunchEffect"] = null;
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added CustomLaunchEffect=null to {fence.Title}");
                    }
                    else
                    {
                        string customEffect = fenceDict["CustomLaunchEffect"]?.ToString();
                        if (!string.IsNullOrEmpty(customEffect) && !validEffects.Contains(customEffect))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid CustomLaunchEffect '{customEffect}' in {fence.Title}, resetting to null");
                            fenceDict["CustomLaunchEffect"] = null;
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate IsHidden
                    if (!fenceDict.ContainsKey("IsHidden"))
                    {
                        fenceDict["IsHidden"] = "false"; // Default to string "false"
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added IsHidden=\"false\" to {fence.Title}");
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
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid IsHidden value '{fenceDict["IsHidden"]}' in {fence.Title}, resetting to \"false\"");
                            isHidden = false;
                            jsonModified = true;
                        }
                        fenceDict["IsHidden"] = isHidden.ToString().ToLower(); // Store as "true" or "false"
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Preserved IsHidden: \"{isHidden.ToString().ToLower()}\" for {fence.Title}");
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
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added IsLink=false to existing item {filename} in {fence.Title}");
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
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid IsLink value '{isLinkToken}' in {fence.Title}, resetting to false");
                                        isLink = false;
                                        itemsModified = true;
                                    }
                                    itemDict["IsLink"] = isLink; // Store as boolean
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Preserved IsLink: {isLink} for item in {fence.Title}");
                                }
                            }
                        }




                        if (itemsModified)
                        {
                            fenceDict["Items"] = items; // Ensure Items updates are captured
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate TextColor
                    if (!fenceDict.ContainsKey("TextColor"))
                    {
                        fenceDict["TextColor"] = null;
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added TextColor=null to {fence.Title}");
                    }
                    else
                    {
                        string textColor = fenceDict["TextColor"]?.ToString();
                        if (!string.IsNullOrEmpty(textColor) && !validColors.Contains(textColor))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid TextColor '{textColor}' in {fence.Title}, resetting to null");
                            fenceDict["TextColor"] = null;
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
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added IsNetwork={isNetwork} to existing item {filename} in {fence.Title}");
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
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid IsNetwork value '{isNetworkToken}' in {fence.Title}, resetting to false");
                                        isNetwork = false;
                                        itemsModified = true;
                                    }
                                    itemDict["IsNetwork"] = isNetwork; // Store as boolean
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Preserved IsNetwork: {isNetwork} for item in {fence.Title}");
                                }
                            }
                        }

                        if (itemsModified)
                        {
                            fenceDict["Items"] = items; // Ensure Items updates are captured
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate DisplayOrder for items in Data fences
                    if (fence.ItemsType?.ToString() == "Data")
                    {
                        var items = fence.Items as JArray ?? new JArray();
                        bool itemsModified = false;
                        int orderCounter = 0; // Start from 0 for sequential ordering

                        foreach (var item in items)
                        {
                            var itemDict = item as JObject;
                            if (itemDict != null)
                            {
                                if (!itemDict.ContainsKey("DisplayOrder"))
                                {
                                    // Add DisplayOrder for new items based on current position in array
                                    itemDict["DisplayOrder"] = orderCounter;
                                    itemsModified = true;
                                    string filename = itemDict["Filename"]?.ToString() ?? "Unknown";
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added DisplayOrder={orderCounter} to item {filename} in {fence.Title}");
                                }
                                else
                                {
                                    // Validate existing DisplayOrder values
                                    int displayOrder = 0;
                                    var displayOrderToken = itemDict["DisplayOrder"];

                                    if (displayOrderToken?.Type == JTokenType.Integer)
                                    {
                                        displayOrder = displayOrderToken.Value<int>();
                                    }
                                    else if (displayOrderToken?.Type == JTokenType.String)
                                    {
                                        if (!int.TryParse(displayOrderToken.Value<string>(), out displayOrder))
                                        {
                                            displayOrder = orderCounter; // Reset to sequential if invalid
                                            itemsModified = true;
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid DisplayOrder '{displayOrderToken}' in {fence.Title}, reset to {displayOrder}");
                                        }
                                    }
                                    else
                                    {
                                        displayOrder = orderCounter; // Reset to sequential if invalid type
                                        itemsModified = true;
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid DisplayOrder type '{displayOrderToken?.Type}' in {fence.Title}, reset to {displayOrder}");
                                    }

                                    itemDict["DisplayOrder"] = displayOrder; // Store as integer
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Preserved DisplayOrder: {displayOrder} for item in {fence.Title}");
                                }
                                orderCounter++; // Increment for next item
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
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Migrated fences.json with updated fields");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, "No migration needed for fences.json");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error migrating fences.json: {ex.Message}\nStackTrace: {ex.StackTrace}");
            }
        }


        private static void UpdateLockState(TextBlock lockIcon, dynamic fence, bool? forceState = null, bool saveToJson = true)
        {
            // Get the actual fence from _fenceData using Id to ensure correct reference
            string fenceId = fence.Id?.ToString();
            if (string.IsNullOrEmpty(fenceId))
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Fence '{fence.Title}' has no Id, cannot update lock state");
                return;
            }

            int index = _fenceData.FindIndex(f => f.Id?.ToString() == fenceId);
            if (index < 0)
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Fence '{fence.Title}' not found in _fenceData, cannot update lock state");
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
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Could not find NonActivatingWindow for fence '{actualFence.Title}'");
                    return;
                }

                // Update ResizeMode
                win.ResizeMode = isLocked ? ResizeMode.NoResize : ResizeMode.CanResizeWithGrip;
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Set ResizeMode to {win.ResizeMode} for fence '{actualFence.Title}'");
            });

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Updated lock state for fence '{actualFence.Title}': IsLocked={isLocked}");
        }


        public static void CreateFence(dynamic fence, TargetChecker targetChecker)
        {



            // Check for valid Portal Fence target folder
            if (fence.ItemsType?.ToString() == "Portal")
            {
                string targetPath = fence.Path?.ToString();
                if (string.IsNullOrEmpty(targetPath) || !System.IO.Directory.Exists(targetPath))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation, $"Skipping creation of Portal Fence '{fence.Title}' due to missing target folder: {targetPath ?? "null"}");
                    _fenceData.Remove(fence);
                    SaveFenceData();
                    return;
                }

            }
            DockPanel dp = new DockPanel();
            //Border cborder = new Border
            //{
            //    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 0, 0, 0)),
            //    CornerRadius = new CornerRadius(6),
            //    Child = dp
            //};

            //// Get fence border color from fence data
            //System.Windows.Media.Brush borderBrush = null;
            //double borderThickness = 0;
            //try
            //{
            //    string borderColorName = fence.FenceBorderColor?.ToString();
            //    if (!string.IsNullOrEmpty(borderColorName))
            //    {
            //        var borderColor = Utility.GetColorFromName(borderColorName);
            //        borderBrush = new SolidColorBrush(borderColor);
            //        borderThickness = 2; // Default thickness when color is specified
            //        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Applied border color '{borderColorName}' to fence '{fence.Title}'");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Error applying border color: {ex.Message}");
            //    borderBrush = null;
            //    borderThickness = 0;
            //}

            // Get fence border color and thickness from fence data
            System.Windows.Media.Brush borderBrush = null;
            double borderThickness = 0;
            try
            {
                string borderColorName = fence.FenceBorderColor?.ToString();
                int customThickness = Convert.ToInt32(fence.FenceBorderThickness?.ToString() ?? "2");

                if (!string.IsNullOrEmpty(borderColorName))
                {
                    var borderColor = Utility.GetColorFromName(borderColorName);
                    borderBrush = new SolidColorBrush(borderColor);
                    borderThickness = customThickness; // Use custom thickness
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Applied border color '{borderColorName}' with thickness {borderThickness} to fence '{fence.Title}'");
                }
                else if (customThickness > 0)
                {
                    // Even without color, apply thickness with default color
                    borderBrush = new SolidColorBrush(System.Windows.Media.Colors.Gray);
                    borderThickness = customThickness;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Applied default border with thickness {borderThickness} to fence '{fence.Title}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Error applying border: {ex.Message}");
                borderBrush = null;
                borderThickness = 0;
            }





            Border cborder = new Border
            {
                Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 0, 0, 0)),
                CornerRadius = new CornerRadius(6),
                BorderBrush = borderBrush, // Apply border color
                BorderThickness = new Thickness(borderThickness), // Apply border thickness
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
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Could not find NonActivatingWindow for lock icon click in fence '{fence.Title}'");
                        return;
                    }

                    // Find the fence in _fenceData using the window's Tag (Id)
                    string fenceId = win.Tag?.ToString();
                    if (string.IsNullOrEmpty(fenceId))
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Fence Id is missing for window '{win.Title}'");
                        return;
                    }

                    dynamic currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    if (currentFence == null)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Fence with Id '{fenceId}' not found in _fenceData");
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
            //     var validEffects = Enum.GetNames(typeof(LaunchEffect)).ToHashSet();
            var validEffects = Enum.GetNames(typeof(LaunchEffectsManager.LaunchEffect)).ToHashSet();
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
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Initiating Peek Behind for fence '{fence.Title}'");

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
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Created countdown window for fence '{fence.Title}' at ({countdownWindow.Left}, {countdownWindow.Top})");

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
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Peek Behind completed for fence '{fence.Title}'");
                    }
                };

                // Start fade out and timer
                win.BeginAnimation(UIElement.OpacityProperty, fadeOut);
                timer.Start();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Started Peek Behind fade-out and countdown for fence '{fence.Title}'");
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
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Opened folder in Explorer: {folderPath}");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Folder path is invalid or does not exist: {folderPath}");
                            MessageBox.Show("The folder path is invalid or does not exist.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error opening folder in Explorer: {ex.Message}");
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
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error reading IsHidden: {ex.Message}");
            }

            bool isHidden = isHiddenString == "true";
            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, $"Fence '{fence.Title}' IsHidden state: {isHidden}");



            // Adjust the fence position to ensure it fits within screen bounds
            AdjustFencePositionToScreen(win);

            win.Loaded += (s, e) =>
            {
                UpdateLockState(lockIcon, fence, null, saveToJson: false);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Applied lock state for fence '{fence.Title}' on load: IsLocked={fence.IsLocked?.ToString().ToLower()}");

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
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"UnrolledHeight {parsedHeight} is invalid (non-positive) for fence '{fence.Title}', using Height={unrolledHeight}");
                            }
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Failed to parse UnrolledHeight '{fence.UnrolledHeight}' for fence '{fence.Title}', using Height={unrolledHeight}");
                        }
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"UnrolledHeight is null for fence '{fence.Title}', using Height={unrolledHeight}");
                    }
                    targetHeight = unrolledHeight;
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation, $"Applied rolled-down state for fence '{fence.Title}' on load: Height={targetHeight}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation, $"Applied rolled-up state for fence '{fence.Title}' on load: Height={targetHeight}");
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
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Set initial WrapPanel visibility to {(isRolled ? "Collapsed" : "Visible")} for fence '{fence.Title}'");
                            }
                        }
                    }
                }

                if (isHidden)
                {
                    win.Visibility = Visibility.Hidden;
                    TrayManager.AddHiddenFence(win);
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation, $"Hid fence '{fence.Title}' after loading at startup");
                }
            };
            // Step 5



            if (isHidden)
            {
                TrayManager.AddHiddenFence(win);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Added fence '{fence.Title}' to hidden list at startup");
            }

            // Hide click
            miHide.Click += (s, e) =>
            {
                UpdateFenceProperty(fence, "IsHidden", "true", $"Hid fence '{fence.Title}'");
                TrayManager.AddHiddenFence(win);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Triggered Hide Fence for '{fence.Title}'");
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
                    //  DebugLog("HEIGHT_UPDATE", fenceId, $"Set fence.Height to {newHeight:F1}");
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

                    //  DebugLog("EVENT", fenceId, "SizeChanged COMPLETED");
                }
                else
                {
                    //   DebugLog("ERROR", fenceId, "Fence not found in _fenceData during resize");
                }
            };



            win.KeyDown += (sender, e) =>
            {
                if (IconDragDropManager.IsDragging && e.Key == Key.Escape)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Escape pressed during drag, cancelling operation");
                    IconDragDropManager.CancelDrag();
                    e.Handled = true;
                }


                // Process InterCore key inputs
                //   InterCore.ProcessKeyInput(e.Key);
            };


            // Make window focusable for key events during drag
            win.Focusable = true;


            win.Show();



            miNewFence.Click += (s, e) =>
            {


                System.Windows.Point mousePosition = win.PointToScreen(new System.Windows.Point(0, 0));
                mousePosition = Mouse.GetPosition(win);
                System.Windows.Point absolutePosition = win.PointToScreen(mousePosition);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Creating new fence at position: X={absolutePosition.X}, Y={absolutePosition.Y}");
                CreateNewFence("New Fence", "Data", absolutePosition.X, absolutePosition.Y);
            };

            miNewPortalFence.Click += (s, e) =>
            {


                System.Windows.Point mousePosition = win.PointToScreen(new System.Windows.Point(0, 0));
                mousePosition = Mouse.GetPosition(win);
                System.Windows.Point absolutePosition = win.PointToScreen(mousePosition);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Creating new portal fence at position: X={absolutePosition.X}, Y={absolutePosition.Y}");
                CreateNewFence("New Portal Fence", "Portal", absolutePosition.X, absolutePosition.Y);
            };



            // Check if title should be bold
            bool isBoldTitle = false;
            try
            {
                isBoldTitle = fence.BoldTitleText?.ToString().ToLower() == "true";
            }
            catch { /* Safe fallback */ }

            // Get title text color or use default white
            System.Windows.Media.Brush titleTextBrush = System.Windows.Media.Brushes.White; // Default
            try
            {
                string titleColorName = fence.TitleTextColor?.ToString();
                if (!string.IsNullOrEmpty(titleColorName))
                {
                    var titleColor = Utility.GetColorFromName(titleColorName);
                    titleTextBrush = new SolidColorBrush(titleColor);
                }
            }
            catch
            {
                titleTextBrush = System.Windows.Media.Brushes.White; // Fallback
            }

            //Label titlelabel = new Label
            //{
            //    Content = fence.Title.ToString(),
            //    Foreground = titleTextBrush, // Changed from hardcoded White
            //    HorizontalContentAlignment = HorizontalAlignment.Center,
            //    HorizontalAlignment = HorizontalAlignment.Stretch,
            //    Cursor = Cursors.SizeAll,
            //    FontWeight = isBoldTitle ? FontWeights.Bold : FontWeights.Normal
            //};



            // Get title text size from fence data
            double titleFontSize = 12; // Default Medium size
            try
            {
                string titleSizeValue = fence.TitleTextSize?.ToString() ?? "Medium";
                switch (titleSizeValue)
                {
                    case "Small":
                        titleFontSize = 10;
                        break;
                    case "Large":
                        titleFontSize = 16;
                        break;
                    default: // Medium
                        titleFontSize = 12;
                        break;
                }
            }
            catch
            {
                titleFontSize = 12; // Fallback to Medium
            }

            Label titlelabel = new Label
            {
                Content = fence.Title.ToString(),
                Foreground = titleTextBrush, // Changed from hardcoded White
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Cursor = Cursors.SizeAll,
                FontWeight = isBoldTitle ? FontWeights.Bold : FontWeights.Normal,
                FontSize = titleFontSize // Apply custom title text size
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
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Fence Id is missing for window '{win.Title}'");
                return;
            }
            dynamic currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
            if (currentFence == null)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Fence with Id '{fenceId}' not found in _fenceData");
                return;
            }
            bool isLocked = currentFence.IsLocked?.ToString().ToLower() == "true";

            titlelabel.MouseDown += (sender, e) =>
            {
                if (e.ClickCount == 2)
                {


                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Entering edit mode for fence: {fence.Title}");

                    titletb.Text = titlelabel.Content.ToString();
                    titlelabel.Visibility = Visibility.Collapsed;
                    titletb.Visibility = Visibility.Visible;
                    win.ShowActivated = true;
                    win.Activate();
                    Keyboard.Focus(titletb);
                    titletb.SelectAll();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Focus set to title textbox for fence: {fence.Title}");
                }
                //  else if (e.LeftButton == MouseButtonState.Pressed)
                else if (e.LeftButton == MouseButtonState.Pressed)
                {
                    string fenceId = win.Tag?.ToString();
                    if (string.IsNullOrEmpty(fenceId))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Fence Id is missing for window '{win.Title}' during MouseDown");
                        return;
                    }
                    dynamic currentFence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    if (currentFence == null)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Fence with Id '{fenceId}' not found in _fenceData during MouseDown");
                        return;
                    }
                    bool isLocked = currentFence.IsLocked?.ToString().ToLower() == "true";
                    if (!isLocked)
                    {
                        win.DragMove();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Dragging fence '{currentFence.Title}'");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"DragMove blocked for locked fence '{currentFence.Title}'");
                    }
                }
            };




            titletb.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    string originalTitle = fence.Title.ToString();
                    string newTitle = titletb.Text;

                    //Process through InterCore for special triggers
                    string finalTitle = InterCore.ProcessTitleChange(fence, newTitle, originalTitle);

                    fence.Title = finalTitle;
                    titlelabel.Content = finalTitle;
                    win.Title = finalTitle;
                    titletb.Visibility = Visibility.Collapsed;
                    titlelabel.Visibility = Visibility.Visible;
                    SaveFenceData();
                    win.ShowActivated = false;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Exited edit mode via Enter, final title for fence: {fence.Title}");
                    win.Focus();
                }
            };

            titletb.LostFocus += (sender, e) =>
            {
                string originalTitle = fence.Title.ToString();
                string newTitle = titletb.Text;

                //Process through InterCore for special triggers
                string finalTitle = InterCore.ProcessTitleChange(fence, newTitle, originalTitle);

                fence.Title = finalTitle;
                titlelabel.Content = finalTitle;
                win.Title = finalTitle;
                titletb.Visibility = Visibility.Collapsed;
                titlelabel.Visibility = Visibility.Visible;
                SaveFenceData();
                win.ShowActivated = false;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Exited edit mode via click, final title for fence: {fence.Title}");
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
                    double opacity = (SettingsManager.PortalBackgroundOpacity / 100.0);
                    wpcontscr.Background = new ImageBrush
                    {
                        ImageSource = new BitmapImage(new Uri("pack://application:,,,/Resources/portal.png")),
                        //  Opacity = 0.2,
                        Opacity = opacity,
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
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Set initial WrapPanel visibility to {(isRolled ? "Collapsed" : "Visible")} for fence '{fence.Title}' in InitContent");

                wpcont.Children.Clear();
                if (fence.ItemsType?.ToString() == "Data")
                {
                    var items = fence.Items as JArray;
                    if (items != null)
                    {
                        // Sort items by DisplayOrder before adding to WrapPanel
                        var sortedItems = items
                            .OfType<JObject>()
                            .OrderBy(item =>
                            {
                                var orderToken = item["DisplayOrder"];
                                if (orderToken?.Type == JTokenType.Integer)
                                {
                                    return orderToken.Value<int>();
                                }
                                // Fallback to 0 if DisplayOrder is missing or invalid
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Missing or invalid DisplayOrder for item {item["Filename"]}, using 0");
                                return 0;
                            })
                            .ToList();

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Loading {sortedItems.Count} items in display order for fence '{fence.Title}'");

                        foreach (dynamic icon in sortedItems)
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

                                // Only add to TargetChecker based on type and settings
                                bool itemIsLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];
                                bool itemIsNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"];
                                bool allowNetworkChecking = _options.CheckNetworkPaths ?? false;

                                if (!itemIsLink && (!itemIsNetwork || allowNetworkChecking))
                                {
                                    targetChecker.AddCheckAction(filePath, () => UpdateIcon(sp, filePath, isFolder), isFolder);
                                    if (itemIsNetwork && allowNetworkChecking)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Added network path {filePath} to target checking (user enabled)");
                                    }
                                }
                                else
                                {
                                    if (itemIsLink)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Excluded web link {filePath} from target checking");
                                    }
                                    if (itemIsNetwork && !allowNetworkChecking)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Excluded network path {filePath} from target checking (user setting)");
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

                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Removing icon for {filePath} from fence");
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
                                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Deleted shortcut: {shortcutPath}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to delete shortcut {shortcutPath}: {ex.Message}");
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
                                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Deleted backup shortcut: {backupPath}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to delete backup shortcut {backupPath}: {ex.Message}");
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
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Run as admin with arguments: {runArguments} for {filePath}");
                                    }
                                    try
                                    {
                                        Process.Start(psi);
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Successfully launched {targetPath} as admin");
                                    }
                                    catch (Exception ex)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to launch {targetPath} as admin: {ex.Message}");
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
                    // string[] droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
                    string[] droppedFiles = null;

                    try
                    {
                        droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                            $"Drop handler received {droppedFiles?.Length ?? 0} files");
                    }
                    catch (Exception dataEx)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                            $"Error getting drop data: {dataEx.Message}");
                        return;
                    }

                    if (droppedFiles == null) return;



                    foreach (string droppedFile in droppedFiles)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                            $"Processing dropped item: '{droppedFile}'");
                        try
                        {

                            // Enhanced Unicode path validation for folders and files
                            bool fileExists = false;
                            bool directoryExists = false;

                            try
                            {
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                    $"Checking existence of: '{droppedFile}'");

                                fileExists = System.IO.File.Exists(droppedFile);
                                directoryExists = System.IO.Directory.Exists(droppedFile);

                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                    $"Existence check results - File: {fileExists}, Directory: {directoryExists}");
                            }
                            catch (Exception pathEx)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                    $"Path validation error for Unicode path '{droppedFile}': {pathEx.Message}");
                                continue;
                            }

                            if (!fileExists && !directoryExists)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                    $"Invalid file or directory: '{droppedFile}'");
                                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Invalid file or directory: {droppedFile}", "Error");
                                continue;
                            }

                            if (directoryExists)
                            {
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                    $"Successfully validated Unicode folder: '{droppedFile}'");
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
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Detected web link: {droppedFile} -> {webUrl}");
                                }


                                else
                                {
                                    // Handle regular file/folder shortcuts with Unicode support
                                    if (isDroppedShortcut)
                                    {
                                        // Use enhanced Unicode-safe shortcut target resolution
                                        targetPath = GetShortcutTargetUnicodeSafe(droppedFile);

                                        // Enhanced logging for debugging
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                            $"Unicode-safe target resolution for dropped shortcut: {droppedFile} -> '{targetPath}'");

                                        if (string.IsNullOrEmpty(targetPath))
                                        {
                                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                                                $"Failed to resolve target for Unicode shortcut: {droppedFile}");
                                            // Default to treating as file if we can't resolve
                                            isFolder = false;
                                        }
                                        else
                                        {
                                            // Determine if target is a folder
                                            isFolder = System.IO.Directory.Exists(targetPath);

                                            // Additional check for folder shortcuts without extensions
                                            if (!isFolder && string.IsNullOrEmpty(System.IO.Path.GetExtension(targetPath)))
                                            {
                                                isFolder = System.IO.Directory.Exists(targetPath);
                                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                                    $"Checked extensionless target for folder: {targetPath} -> isFolder={isFolder}");
                                            }

                                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                $"Shortcut classification: {droppedFile} -> target='{targetPath}', isFolder={isFolder}");
                                        }
                                    }
                                    else
                                    {
                                        // Handle direct files/folders
                                        targetPath = droppedFile;
                                        isFolder = System.IO.Directory.Exists(targetPath);
                                    }
                                }

                                if (!isDroppedShortcut && !isDroppedUrlFile)
                                {
                                    // Create shortcut for dropped file/folder
                                    shortcutName = baseShortcutName + ".lnk";
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{System.IO.Path.GetFileNameWithoutExtension(droppedFile)} ({counter++}).lnk");
                                    }
          
                                    try
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                            $"Creating shortcut: '{droppedFile}' -> '{shortcutName}'");

                                        // Check if shortcut name or target contains Unicode characters
                                        bool shortcutHasUnicode = shortcutName.Any(c => c > 127);
                                        bool targetHasUnicode = droppedFile.Any(c => c > 127);

                                        if (shortcutHasUnicode || targetHasUnicode)
                                        {
                                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                $"Handling Unicode shortcut - name: {shortcutHasUnicode}, target: {targetHasUnicode}");

                                            // Create shortcut with ASCII-safe temporary name
                                            string tempShortcutName = System.IO.Path.Combine("Shortcuts", $"temp_{DateTime.Now.Ticks}.lnk");

                                            // For target with Unicode, use a known working folder first
                                            string tempTarget = targetHasUnicode ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : droppedFile;

                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                                $"Creating temp shortcut: '{tempTarget}' -> '{tempShortcutName}'");

                                            WshShell shell = new WshShell();
                                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(tempShortcutName);
                                            shortcut.TargetPath = tempTarget;
                                            shortcut.Save();

                                 

                                            // If target has Unicode, use improved Unicode shortcut creation
                                            if (targetHasUnicode && System.IO.File.Exists(tempShortcutName))
                                            {
                                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                                    $"Creating Unicode-compatible shortcut from '{tempTarget}' to '{droppedFile}'");

                                                try
                                                {
                                                    // Delete the temp shortcut and create a new one with proper Unicode handling
                                                    System.IO.File.Delete(tempShortcutName);

                                                    // Use .NET's approach for Unicode folders - create shortcut to explorer with folder argument
                                                    WshShell unicodeShell = new WshShell();
                                                    IWshShortcut unicodeShortcut = (IWshShortcut)unicodeShell.CreateShortcut(tempShortcutName);

                                                    if (directoryExists)
                                                    {
                                                        // For Unicode folders, use explorer.exe with quoted path argument
                                                        unicodeShortcut.TargetPath = "explorer.exe";
                                                        unicodeShortcut.Arguments = "\"" + droppedFile + "\"";
                                                        unicodeShortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(droppedFile);

                                                        // Set icon to folder icon
                                                        unicodeShortcut.IconLocation = "shell32.dll,3";

                                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                            $"Created Unicode folder shortcut using explorer.exe method");
                                                    }
                                                    else
                                                    {
                                                        // For Unicode files, try direct approach first
                                                        unicodeShortcut.TargetPath = droppedFile;
                                                        unicodeShortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(droppedFile);

                                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                            $"Created Unicode file shortcut using direct method");
                                                    }

                                                    unicodeShortcut.Save();

                                                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                        $"Successfully created Unicode shortcut with proper target resolution");
                                                }
                                                catch (Exception unicodeEx)
                                                {
                                                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                                        $"Unicode shortcut creation failed: {unicodeEx.Message}");

                                                    // Final fallback - recreate shortcut pointing directly to parent folder
                                                    try
                                                    {
                                                        string parentFolder = System.IO.Path.GetDirectoryName(droppedFile);
                                                        WshShell fallbackShell = new WshShell();
                                                        IWshShortcut fallbackShortcut = (IWshShortcut)fallbackShell.CreateShortcut(tempShortcutName);
                                                        fallbackShortcut.TargetPath = "explorer.exe";
                                                        fallbackShortcut.Arguments = "\"" + parentFolder + "\"";
                                                        fallbackShortcut.Save();

                                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                            $"Created fallback shortcut to parent folder");
                                                    }
                                                    catch (Exception fallbackEx)
                                                    {
                                                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                                            $"Fallback shortcut creation failed: {fallbackEx.Message}");
                                                    }
                                                }
                                            }

                                            // Rename to final Unicode name if needed
                                            if (shortcutHasUnicode)
                                            {
                                                System.IO.File.Move(tempShortcutName, shortcutName);
                                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                    $"Renamed to Unicode filename: {shortcutName}");
                                            }
                                            else
                                            {
                                                System.IO.File.Move(tempShortcutName, shortcutName);
                                            }

                                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                $"Successfully created Unicode shortcut: {shortcutName}");
                                        }
                                        else
                                        {
                                            // Regular shortcut creation for non-Unicode paths
                                            WshShell shell = new WshShell();
                                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutName);
                                            shortcut.TargetPath = droppedFile;

                                            if (directoryExists)
                                            {
                                                shortcut.WorkingDirectory = droppedFile;
                                            }

                                            shortcut.Save();

                                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                                $"Successfully created regular shortcut: {shortcutName}");
                                        }
                                    }
                                    catch (Exception shortcutEx)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                            $"Failed to create shortcut for '{droppedFile}': {shortcutEx.Message}");
                                        continue;
                                    }

                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Created unique shortcut: {shortcutName}");
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
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Created web link shortcut: {shortcutName} -> {webUrl}");
                                    }
                                    else
                                    {
                                        // Copy regular shortcut
                                        System.IO.File.Copy(droppedFile, shortcutName, false);
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Copied unique shortcut: {shortcutName}");
                                    }
                                }

            
                                dynamic newItem = new System.Dynamic.ExpandoObject();
                                IDictionary<string, object> newItemDict = newItem;

                                // Enhanced Unicode handling for folder items
                                if (System.IO.Directory.Exists(droppedFile))
                                {
                                    // This is a folder - use Unicode-safe validation
                                    string validatedPath;
                                    if (ValidateUnicodeFolderPath(droppedFile, out validatedPath))
                                    {
                                        newItemDict["Filename"] = shortcutName;
                                        newItemDict["IsFolder"] = true;
                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                            $"Unicode folder processed successfully: {droppedFile} -> {shortcutName}");
                                    }
                                    else
                                    {
                                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                            $"Failed Unicode folder validation: {droppedFile}");
                                        continue; // Skip this item
                                    }
                                }
                                else
                                {
                                    // Regular file/shortcut handling
                                    newItemDict["Filename"] = shortcutName;
                                    newItemDict["IsFolder"] = isFolder;
                                }







                                newItemDict["IsLink"] = isWebLink; // Set IsLink property
                                newItemDict["IsNetwork"] = IsNetworkPath(shortcutName); // Set IsNetwork property
                                newItemDict["DisplayName"] = System.IO.Path.GetFileNameWithoutExtension(droppedFile);


                                // Add DisplayOrder as the next available order number
                                var items = fence.Items as JArray ?? new JArray();
                                int nextDisplayOrder = items.Count; // Use array length as next order
                                newItemDict["DisplayOrder"] = nextDisplayOrder;
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Assigned DisplayOrder={nextDisplayOrder} to new dropped item {shortcutName}");


                                // Log the detection results
                                if ((bool)newItemDict["IsNetwork"])
                                {
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Detected network path for new item: {shortcutName}");
                                }

                                //  var items = fence.Items as JArray ?? new JArray();
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
                                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Removing icon for {shortcutName} from fence");
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
                                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Deleted shortcut: {shortcutPath}");
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to delete shortcut {shortcutPath}: {ex.Message}");
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
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Copied target folder path to clipboard: {folderPath}");
                                    };

                                    miCopyFullPath.Click += (sender, e) =>
                                    {
                                        string targetPath = Utility.GetShortcutTarget(shortcutName);
                                        Clipboard.SetText(targetPath);
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Copied target full path to clipboard: {targetPath}");
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
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Added shortcut to Data Fence: {shortcutName}");
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
                                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation, $"No Path defined for Portal Fence: {fence.Title}");
                                    continue;
                                }

                                if (!System.IO.Directory.Exists(destinationFolder))
                                {
                                    // MessageBox.Show($"The destination folder '{destinationFolder}' no longer exists. Please update the Portal Fence settings.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    TrayManager.Instance.ShowOKOnlyMessageBoxForm($"The destination folder '{destinationFolder}' no longer exists. Please update the Portal Fence settings.", "Error");
                                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation, $"Destination folder missing for Portal Fence: {destinationFolder}");
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
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Copied file to Portal Fence: {destinationPath}");
                                }
                                else if (System.IO.Directory.Exists(droppedFile))
                                {
                                    BackupManager.CopyDirectory(droppedFile, destinationPath);
                                    //   CopyDirectory(droppedFile, destinationPath);
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Copied directory to Portal Fence: {destinationPath}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            //  MessageBox.Show($"Failed to add {droppedFile}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to add {droppedFile}: {ex.Message}", "Error");
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation, $"Failed to add {droppedFile}: {ex.Message}");
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
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate, $"Fence Id missing during position change for window '{win.Title}'");
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
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate, $"Position updated for fence '{currentFence.Title}' to X={win.Left}, Y={win.Top}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate, $"Fence with Id '{fenceId}' not found during position change");
                }
            };

            InitContent();
            win.Show();


            IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
            SnapManager.AddSnapping(win, fenceDict);
            // Apply custom color if present, otherwise use global
            string customColor = fence.CustomColor?.ToString();
            Utility.ApplyTintAndColorToFence(win, string.IsNullOrEmpty(customColor) ? SettingsManager.SelectedColor : customColor);
            targetChecker.Start();
        }


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
        



            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";

            bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
            bool isLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"]; // Add IsLink detection
            bool isNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"]; // Add IsNetwork detection

            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
            string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;
            string arguments = iconDict.ContainsKey("Arguments") ? (string)iconDict["Arguments"] : null;


            //StackPanel sp = new StackPanel { Margin = new Thickness(5), Width = 60 };


            // Get icon spacing from fence data (default 5 if not found)
            int iconSpacing = 5;
            try
            {
                foreach (var fenceData in _fenceData)
                {
                    if (fenceData.ItemsType?.ToString() == "Data")
                    {
                        var fenceItems = fenceData.Items as JArray;
                        if (fenceItems?.Any(i => i["Filename"]?.ToString() == filePath) == true)
                        {
                            iconSpacing = Convert.ToInt32(fenceData.IconSpacing?.ToString() ?? "5");
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Using IconSpacing={iconSpacing} for {filePath}");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Error getting IconSpacing for {filePath}: {ex.Message}, using default 5");
                iconSpacing = 5; // Fallback to default
            }

            StackPanel sp = new StackPanel
            {
                Margin = new Thickness(iconSpacing), // Use fence-specific spacing
                Width = 60
            };
//      IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
         //   string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";


            //  System.Windows.Controls.Image ico = new System.Windows.Controls.Image { Width = 40, Height = 40, Margin = new Thickness(5) };

            // Determine icon size
            double iconWidth = 40, iconHeight = 40; // Default Medium
            foreach (var fenceData in _fenceData)
            {
                if (fenceData.ItemsType?.ToString() == "Data")
                {
                    var fenceItems = fenceData.Items as JArray;
                    if (fenceItems?.Any(i => i["Filename"]?.ToString() == filePath) == true)
                    {
                        string sizeValue = fenceData.IconSize?.ToString() ?? "Medium";
                        switch (sizeValue)
                        {
                            case "Tiny":
                                iconWidth = iconHeight = 24;
                                break;
                            case "Small":
                                iconWidth = iconHeight = 32;
                                break;
                            case "Large":
                                iconWidth = iconHeight = 48;
                                break;
                            case "Huge":
                                iconWidth = iconHeight = 64;
                                break;
                            default: // Medium
                                iconWidth = iconHeight = 40;
                                break;
                        }
                        break;
                    }
                }
            }

            System.Windows.Controls.Image ico = new System.Windows.Controls.Image { Width = iconWidth, Height = iconHeight, Margin = new Thickness(5) };


      

            // Apply icon effect if enabled
            if (SettingsManager.IconVisibilityEffect != IconVisibilityEffect.None)
            {
                ico.Effect = Utility.CreateIconEffect(SettingsManager.IconVisibilityEffect);
            }







            ImageSource shortcutIcon = null;
            // Handle web link icons first (before other shortcut processing)
            if (isLink)
            {
                shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Using link-White.png for web link {filePath}");
            }
            else if (isShortcut)

            {
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                // targetPath = shortcut.TargetPath;






                targetPath = shortcut.TargetPath;

                // Enhanced handling for Unicode shortcut filenames
                if (string.IsNullOrEmpty(targetPath))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                        $"WshShell returned empty TargetPath for shortcut: {filePath}");

                    // Try Unicode-safe resolution as fallback
                    targetPath = GetShortcutTargetUnicodeSafe(filePath);

                    if (!string.IsNullOrEmpty(targetPath))
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                            $"Unicode-safe method resolved target for AddIcon: {filePath} -> {targetPath}");
                    }
                }

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
                                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling, $"Extracted icon at index {iconIndex} from {iconPath} for {filePath}");
                                }
                                finally
                                {
                                    DestroyIcon(hIcon[0]);
                                }
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Failed to extract icon at index {iconIndex} from {iconPath} for {filePath}. Result: {result}");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, $"Error extracting icon from {iconPath} at index {iconIndex} for {filePath}: {ex.Message}");
                        }
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling, $"Icon file not found: {iconPath} for {filePath}");
                    }
                }


                // Fallback only if no custom icon was successfully extracted
                if (shortcutIcon == null)
                {
                    // Enhanced target path handling for Unicode shortcuts
                    string effectiveTargetPath = targetPath;

                    // If targetPath is empty, try to resolve it again with Unicode-safe method
                    if (string.IsNullOrEmpty(targetPath))
                    {
                        effectiveTargetPath = GetShortcutTargetUnicodeSafe(filePath);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                            $"Re-resolved target for AddIcon: {filePath} -> {effectiveTargetPath}");
                    }

                    if (!string.IsNullOrEmpty(effectiveTargetPath) && System.IO.Directory.Exists(effectiveTargetPath))
                    {
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Using folder-White.png for shortcut {filePath} targeting folder {effectiveTargetPath}");
                    }
                    else if (!string.IsNullOrEmpty(effectiveTargetPath) && System.IO.File.Exists(effectiveTargetPath))
                    {
                        try
                        {
                            shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(effectiveTargetPath).ToImageSource();
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                                $"Successfully extracted icon for Unicode shortcut {filePath}: {effectiveTargetPath}");
                        }
                        catch (Exception ex)
                        {
                            shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                                $"Failed to extract target file icon for {filePath}: {ex.Message}");
                        }
                    }
                    else
                    {
                        shortcutIcon = isFolder
                            ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                            : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")} (target: '{effectiveTargetPath}')");
                    }
                }
            }
            else
            {


                // Enhanced Unicode folder icon handling
                bool folderExists = false;
                try
                {
                    folderExists = System.IO.Directory.Exists(filePath);
                    if (!folderExists && isFolder)
                    {
                        // For items marked as folders, try DirectoryInfo approach
                        var dirInfo = new System.IO.DirectoryInfo(filePath);
                        folderExists = dirInfo.Exists;
                    }
                }
                catch (Exception folderEx)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                        $"Folder existence check failed for Unicode path {filePath}: {folderEx.Message}");
                }

                if (folderExists)
                {
                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                        $"Using folder-White.png for Unicode folder {filePath}");
                }
                else if (isFolder)
                {
                    // Folder marked as folder but doesn't exist - use missing folder icon
                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                        $"Using folder-WhiteX.png for missing Unicode folder {filePath}");
                }
                else if (System.IO.File.Exists(filePath))
                {
                    try
                    {
                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Using file icon for {filePath}");
                    }
                    catch (Exception ex)
                    {
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Failed to extract icon for {filePath}: {ex.Message}");
                    }
                }
                else
                {
                    shortcutIcon = isFolder
                        ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                        : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")}");
                }
            }

            ico.Source = shortcutIcon;

            // Apply grayscale effect if enabled - AFTER setting the source
            try
            {
                bool shouldApplyGrayscale = false;
                foreach (var fenceData in _fenceData)
                {
                    if (fenceData.ItemsType?.ToString() == "Data")
                    {
                        var fenceItems = fenceData.Items as JArray;
                        if (fenceItems?.Any(i => i["Filename"]?.ToString() == filePath) == true)
                        {
                            shouldApplyGrayscale = fenceData.GrayscaleIcons?.ToString().ToLower() == "true";
                            break;
                        }
                    }
                }

                if (shouldApplyGrayscale && shortcutIcon is BitmapSource bitmapSource)
                {
                    // Convert to grayscale using FormatConvertedBitmap
                    var grayscaleImage = new FormatConvertedBitmap(bitmapSource, PixelFormats.Gray8, BitmapPalettes.Gray256, 0);
                    ico.Source = grayscaleImage;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Applied grayscale conversion to {filePath}");
                }
                else if (shouldApplyGrayscale)
                {
                    // Fallback: Strong desaturation effect for non-bitmap sources
                    ico.Effect = new System.Windows.Media.Effects.DropShadowEffect
                    {
                        Color = System.Windows.Media.Colors.Gray,
                        Direction = 0,
                        ShadowDepth = 0,
                        BlurRadius = 0,
                        Opacity = 0.8
                    };
                    ico.Opacity = 0.6; // Reduce overall opacity to simulate grayscale
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Applied grayscale effect fallback to {filePath}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Error applying grayscale effect: {ex.Message}");
            }




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

                // Create network indicator (🔗 symbol in blue)
                TextBlock networkIndicator = new TextBlock
                {
                    Text = "🔗",
                    FontSize = 14,

                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(65, 135, 225)), //  Blue #4169E1
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

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Added network indicator to {filePath}");
            }
            else
            {
                // For non-network items, add icon directly as before
                sp.Children.Add(ico);
            }

            string displayText = (!iconDict.ContainsKey("DisplayName") || iconDict["DisplayName"] == null || string.IsNullOrEmpty((string)iconDict["DisplayName"]))
                ? System.IO.Path.GetFileNameWithoutExtension(filePath)
                : (string)iconDict["DisplayName"];
            //if (displayText.Length > 20)
            //{
            //    displayText = displayText.Substring(0, 20) + "...";
            //}

            if (displayText.Length > SettingsManager.MaxDisplayNameLength)
            {
                displayText = displayText.Substring(0, SettingsManager.MaxDisplayNameLength) + "...";
            }


            // Get text color from fence data or use default white
            string textColorName = null;
            try
            {
                // Find the fence this icon belongs to
                foreach (var fenceData in _fenceData)
                {
                    if (fenceData.ItemsType?.ToString() == "Data")
                    {
                        var fenceItems = fenceData.Items as JArray;
                        if (fenceItems?.Any(i => i["Filename"]?.ToString() == filePath) == true)
                        {
                            textColorName = fenceData.TextColor?.ToString();
                            break;
                        }
                    }
                }
            }
            catch { /* Safe fallback */ }

            System.Windows.Media.Brush textBrush = System.Windows.Media.Brushes.White; // Default
            if (!string.IsNullOrEmpty(textColorName))
            {
                try
                {
                    var textColor = Utility.GetColorFromName(textColorName);
                    textBrush = new SolidColorBrush(textColor);
                }
                catch
                {
                    textBrush = System.Windows.Media.Brushes.White; // Fallback
                }
            }


            //TextBlock lbl = new TextBlock
            //{
            //    TextWrapping = TextWrapping.Wrap,
            //    TextTrimming = TextTrimming.None,
            //    HorizontalAlignment = HorizontalAlignment.Center,
            //    //Foreground = System.Windows.Media.Brushes.White,
            //    Foreground = textBrush, // Changed from hardcoded White
            //    MaxWidth = double.MaxValue,
            //    Width = double.NaN,
            //    TextAlignment = TextAlignment.Center,
            //    Text = displayText
            //};
            //lbl.Effect = new DropShadowEffect
            //{
            //    Color = Colors.Black,
            //    Direction = 315,
            //    ShadowDepth = 2,
            //    BlurRadius = 3,
            //    Opacity = 0.8
            //};

            TextBlock lbl = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                TextTrimming = TextTrimming.None,
                HorizontalAlignment = HorizontalAlignment.Center,
                Foreground = textBrush,
                MaxWidth = double.MaxValue,
                Width = double.NaN,
                TextAlignment = TextAlignment.Center,
                Text = displayText
            };

            // Apply shadow effect unless disabled
            bool disableTextShadow = false;
            try
            {
                // Find the fence this icon belongs to
                foreach (var fenceData in _fenceData)
                {
                    if (fenceData.ItemsType?.ToString() == "Data")
                    {
                        var fenceItems = fenceData.Items as JArray;
                        if (fenceItems?.Any(i => i["Filename"]?.ToString() == filePath) == true)
                        {
                            disableTextShadow = fenceData.DisableTextShadow?.ToString().ToLower() == "true";
                            break;
                        }
                    }
                }
            }
            catch { /* Safe fallback */ }

            if (!disableTextShadow)
            {
                lbl.Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 315,
                    ShadowDepth = 2,
                    BlurRadius = 3,
                    Opacity = 0.8
                };
            }

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
            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, $"Saved fences.json with consistent IsHidden, IsRolled, and UnrolledHeight string format");
        }

        private static void CreateNewFence(string title, string itemsType, double x = 20, double y = 20, string customColor = null, string customLaunchEffect = null)
        {
            // Generate random name instead of using the passed title
            string fenceName = GenerateRandomName();

            dynamic newFence = new System.Dynamic.ExpandoObject();
            newFence.Id = Guid.NewGuid().ToString();
            IDictionary<string, object> newFenceDict = newFence;
            // newFenceDict["Title"] = title;
            
            
          //  newFenceDict["Title"] = fenceName; // Use random name
          //Option to set Portal fences fodler nane
            // Only use random name for non-Portal fences
            if (itemsType != "Portal")
            {
                newFenceDict["Title"] = fenceName; // Use random name
            }

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

            // Initialize ALL fence properties with defaults to match JSON structure
            newFenceDict["IsLocked"] = "false";
            newFenceDict["IsHidden"] = "false";
            newFenceDict["CustomColor"] = customColor;
            newFenceDict["CustomLaunchEffect"] = customLaunchEffect;
            newFenceDict["IsRolled"] = "false";
            newFenceDict["UnrolledHeight"] = 130;
            newFenceDict["TextColor"] = null;
            newFenceDict["BoldTitleText"] = "false";
            newFenceDict["TitleTextColor"] = null;
            newFenceDict["DisableTextShadow"] = "false";
            newFenceDict["IconSize"] = "Medium";
            newFenceDict["GrayscaleIcons"] = "false";
            newFenceDict["IconSpacing"] = 5;
            newFenceDict["TitleTextSize"] = "Medium";
            newFenceDict["FenceBorderColor"] = null;
            newFenceDict["FenceBorderThickness"] = 2;

            if (itemsType == "Portal")
                if (itemsType == "Portal")
                {
                    using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                    {
                        dialog.Description = "Select the folder to monitor for this Portal Fence";
                        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            newFenceDict["Path"] = dialog.SelectedPath;

                            // Get the folder name from the selected path
                            // should add and option to be checked for the below
                            // Set title to folder name for Portal fences
                            string folderName = System.IO.Path.GetFileName(dialog.SelectedPath);
                            newFenceDict["Title"] = folderName;

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

                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling, $"Moving item {filename} from {sourceFence.Title} to {fence.Title}");

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
                                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Item {filename} not found in source fence '{sourceFence.Title}'");
                            }

                            moveWindow.Close();
                    
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
                                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling, $"Item moved successfully to {fence.Title}");
                                }
                                else
                                {
                                    waitWindow.Close();
                                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Skipped fence reload due to application shutdown");
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
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"FindWrapPanel: Reached max depth {maxDepth} or null parent at depth {depth}");
                return null;
            }

            // Check if current element is a WrapPanel
            if (parent is WrapPanel wrapPanel)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"FindWrapPanel: Found WrapPanel at depth {depth}");
                return wrapPanel;
            }

            // Recurse through visual tree
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"FindWrapPanel: Checking child {i} at depth {depth}, type: {child?.GetType()?.Name ?? "null"}");
                var result = FindWrapPanel(child, depth + 1, maxDepth);
                if (result != null)
                {
                    return result;
                }
            }

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"FindWrapPanel: No WrapPanel found under parent {parent?.GetType()?.Name ?? "null"} at depth {depth}");
            return null;
        }

        public static void EditItem(dynamic icon, dynamic fence, NonActivatingWindow win)
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
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Updated DisplayName for {filePath} to {newDisplayName}");

                // Update fence data
                var items = fence.Items as JArray;
                if (items != null)
                {
                    var itemToUpdate = items.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                    if (itemToUpdate != null)
                    {
                        itemToUpdate["DisplayName"] = newDisplayName;
                        SaveFenceData();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Fence data updated for {filePath}");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Failed to find item {filePath} in fence items for update");
                    }
                }

                // Update UI
                if (win == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Window is null for fence when updating {filePath}");
                    return;
                }

                // Attempt to find WrapPanel directly
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Starting UI update for {filePath}. Window content type: {win.Content?.GetType()?.Name ?? "null"}");
                WrapPanel wpcont = null;
                var border = win.Content as Border;
                if (border != null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Found Border. Child type: {border.Child?.GetType()?.Name ?? "null"}");
                    var dockPanel = border.Child as DockPanel;
                    if (dockPanel != null)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Found DockPanel. Checking for ScrollViewer...");
                        var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                        if (scrollViewer != null)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Found ScrollViewer. Content type: {scrollViewer.Content?.GetType()?.Name ?? "null"}");
                            wpcont = scrollViewer.Content as WrapPanel;
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"ScrollViewer not found in DockPanel for {filePath}");
                        }
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"DockPanel not found in Border for {filePath}");
                    }
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Border not found in window content for {filePath}");
                }

                // Fallback to visual tree traversal if direct lookup fails
                if (wpcont == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Direct WrapPanel lookup failed, attempting visual tree traversal for {filePath}");
                    wpcont = FindWrapPanel(win);
                }

                if (wpcont != null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Found WrapPanel. Checking for StackPanel with Tag.FilePath: {filePath}");
                    // Access Tag.FilePath to match ClickEventAdder's anonymous object
                    var sp = wpcont.Children.OfType<StackPanel>()
                        .FirstOrDefault(s => s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == filePath);
                    if (sp != null)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Found StackPanel for {filePath}. Tag: {sp.Tag?.ToString() ?? "null"}");
                        var lbl = sp.Children.OfType<TextBlock>().FirstOrDefault();
                        if (lbl != null)
                        {
                            // Apply truncation logic (same as in AddIcon)
                            string displayText = string.IsNullOrEmpty(newDisplayName)
                                ? System.IO.Path.GetFileNameWithoutExtension(filePath)
                                : newDisplayName;
                            //if (displayText.Length > 20)
                            //{
                            //    displayText = displayText.Substring(0, 20) + "...";
                            //}
                            if (displayText.Length > SettingsManager.MaxDisplayNameLength)
                            {
                                displayText = displayText.Substring(0, SettingsManager.MaxDisplayNameLength) + "...";
                            }
                            lbl.Text = displayText;
                            lbl.InvalidateVisual(); // Force UI refresh
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Updated TextBlock for {filePath} to '{displayText}'");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"TextBlock not found in StackPanel for {filePath}. Children: {string.Join(", ", sp.Children.OfType<FrameworkElement>().Select(c => c.GetType().Name))}");
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
                                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Extracted icon at index {iconIndex} from {iconPath} for {filePath}");
                                            }
                                            finally
                                            {
                                                DestroyIcon(hIcon[0]); // Clean up icon handle
                                            }
                                        }
                                        else
                                        {
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Failed to extract icon at index {iconIndex} from {iconPath} for {filePath}. Result: {result}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, $"Error extracting icon from {iconPath} at index {iconIndex} for {filePath}: {ex.Message}");
                                    }
                                }
                                else
                                {
                                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, $"Icon file not found: {iconPath} for {filePath}");
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
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"No custom icon, using target icon for {filePath}: {targetPath}");
                                    }
                                    catch (Exception ex)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Failed to extract target icon for {filePath}: {ex.Message}");
                                    }
                                }
                                else if (System.IO.Directory.Exists(targetPath))
                                {
                                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"No custom icon, using folder-White.png for {filePath}");
                                }
                                else
                                {
                                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"No custom icon or valid target, using file-WhiteX.png for {filePath}");
                                }
                            }

                            ico.Source = shortcutIcon;
                            iconCache[filePath] = shortcutIcon;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Updated icon for {filePath}");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, $"Image not found in StackPanel for {filePath}");
                        }
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"StackPanel not found for {filePath} in WrapPanel. Available Tag.FilePaths: {string.Join(", ", wpcont.Children.OfType<StackPanel>().Select(s => s.Tag != null ? (s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() ?? "null") : "null"))}");

                        // Fallback: Rebuild single icon
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Attempting to rebuild icon for {filePath}");
                        // Remove existing StackPanel if present (in case Tag lookup failed due to corruption)
                        var oldSp = wpcont.Children.OfType<StackPanel>()
                            .FirstOrDefault(s => s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == filePath);
                        if (oldSp != null)
                        {
                            wpcont.Children.Remove(oldSp);
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Removed old StackPanel for {filePath}");
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
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Rebuilt icon for {filePath} with new StackPanel");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Failed to rebuild icon for {filePath}: No new StackPanel created");
                        }
                    }
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"WrapPanel not found for {filePath}. UI update skipped.");
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
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Created TempShortcuts directory: {tempShortcutsDir}");
                }

                if (!targetExists)
                {
                    // Backup shortcut if target is missing and not already backed up
                    if (System.IO.File.Exists(filePath) && !System.IO.File.Exists(backupPath))
                    {
                        System.IO.File.Copy(filePath, backupPath, true);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Backed up shortcut {filePath} to {backupPath}");
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
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Restored shortcut {filePath} from {backupPath} with custom icon");
                        }
                        // Delete backup
                        System.IO.File.Delete(backupPath);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Deleted backup {backupPath} after restoration");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error in BackupOrRestoreShortcut for {filePath}: {ex.Message}");
            }
        }




        /// <summary>
        /// Parses LNK file binary data to extract target path
        /// Used as fallback when COM methods fail with Unicode filenames
        /// </summary>

        private static string ParseLnkFileBinary(byte[] lnkData)
        {
            try
            {
                if (lnkData.Length < 0x4C) return string.Empty;

                // Verify LNK file header
                if (BitConverter.ToUInt32(lnkData, 0) != 0x0000004C) return string.Empty;

                // Look for paths in the binary data (enhanced for folders)
                string unicodeContent = System.Text.Encoding.Unicode.GetString(lnkData);
                string ansiContent = System.Text.Encoding.Default.GetString(lnkData);

                // Search both encodings for paths (folders and executables)
                foreach (string content in new[] { unicodeContent, ansiContent })
                {
                    // Look for executable paths
                    var exeMatches = System.Text.RegularExpressions.Regex.Matches(content,
                        @"[A-Za-z]:\\[^<>:""|?*\x00-\x1f]*\.(exe|dll|bat|cmd|com)",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                    foreach (System.Text.RegularExpressions.Match match in exeMatches)
                    {
                        string candidatePath = match.Value.Trim('\0', ' ', '\t', '\r', '\n');
                        if (candidatePath.Length > 3 && System.IO.File.Exists(candidatePath))
                        {
                            return candidatePath;
                        }
                    }

                    // Look for folder paths (enhanced for Unicode folders)
                    var folderMatches = System.Text.RegularExpressions.Regex.Matches(content,
                        @"[A-Za-z]:\\[^<>:""|?*\x00-\x1f\x20]+",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                    foreach (System.Text.RegularExpressions.Match match in folderMatches)
                    {
                        string candidatePath = match.Value.Trim('\0', ' ', '\t', '\r', '\n');

                        // Clean up the path
                        candidatePath = candidatePath.Replace("\0", "").Trim();

                        // Check if this looks like a valid path and exists
                        if (candidatePath.Length > 5 && candidatePath.Contains("\\"))
                        {
                            // Try to validate as directory
                            try
                            {
                                if (System.IO.Directory.Exists(candidatePath))
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

                    // Special handling for Unicode characters
                    if (content.Any(c => c > 127))
                    {
                        // Look for Desktop paths with Unicode
                        var unicodeMatches = System.Text.RegularExpressions.Regex.Matches(content,
                            @"[A-Za-z]:[\\\/][^<>:""|?*\x00-\x1f]*[^\x00-\x20]",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        foreach (System.Text.RegularExpressions.Match match in unicodeMatches)
                        {
                            string candidatePath = match.Value;

                            // Clean and normalize the path
                            candidatePath = candidatePath.Replace('/', '\\').Trim('\0', ' ', '\t', '\r', '\n');

                            if (candidatePath.Length > 5)
                            {
                                try
                                {
                                    if (System.IO.Directory.Exists(candidatePath) || System.IO.File.Exists(candidatePath))
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                            $"Found valid Unicode path: {candidatePath}");
                                        return candidatePath;
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"No valid paths found in LNK binary data");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Error parsing LNK binary: {ex.Message}");
                return string.Empty;
            }
        }




        private static void UpdateIcon(StackPanel sp, string filePath, bool isFolder)
        {


            if (Application.Current == null)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, "Application.Current is null, cannot update icon.");
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
                        //   LogManager.Log(LogManager.LogLevel.Debug,LogManager.LogCategory.IconHandling, $"Skipping icon update for web link: {filePath}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Error checking IsLink for {filePath}: {ex.Message}");
                }
                System.Windows.Controls.Image ico = sp.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                if (ico == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"No Image found in StackPanel for {filePath}");
                    return;
                }

                bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
                //  string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;

                string targetPath = filePath;
                if (isShortcut)
                {
                    targetPath = GetShortcutTargetUnicodeSafe(filePath);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                        $"UpdateIcon Unicode-safe resolution: {filePath} -> {targetPath}");
                }



                bool targetExists = System.IO.File.Exists(targetPath) || System.IO.Directory.Exists(targetPath);
                bool isTargetFolder = System.IO.Directory.Exists(targetPath);

                // Correct isFolder for shortcut to folder
                if (isShortcut && isTargetFolder)
                {
                    isFolder = true;
                    // do not remove the below commented log line, it may needed for future debugging
                    // LogManager.Log(LogManager.LogLevel.Debug,LogManager.LogCategory.IconHandling, $"Corrected isFolder to true for shortcut {filePath} targeting folder {targetPath}");
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
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")}");
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
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Extracted custom icon at index {iconIndex} from {iconPath} for {filePath}");
                                    }
                                    finally
                                    {
                                        DestroyIcon(hIcon[0]);
                                    }
                                }
                                else
                                {
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Failed to extract custom icon at index {iconIndex} from {iconPath} for {filePath}. Result: {result}");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Error extracting custom icon from {iconPath} at index {iconIndex} for {filePath}: {ex.Message}");
                            }
                        }
                        else
                        {
                            // do not remove the below commented log line, it may needed for future debugging
                            // LogManager.Log(LogManager.LogLevel.Debug,LogManager.LogCategory.IconHandling, $"Custom icon file not found: {iconPath} for {filePath}");
                        }


                    }



                    // Fallback if no custom icon
                    if (newIcon == null)
                    {
                        if (isTargetFolder)
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                $"Using folder-White.png for shortcut {filePath} targeting folder {targetPath}");
                        }
                        else
                        {
                            // Enhanced target resolution for Unicode shortcuts
                            string iconTargetPath = targetPath;

                            // If targetPath is empty or invalid, use Unicode-safe resolution
                            if (string.IsNullOrEmpty(targetPath) || !System.IO.File.Exists(targetPath))
                            {
                                iconTargetPath = GetShortcutTargetUnicodeSafe(filePath);
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                    $"UpdateIcon using Unicode-safe resolution: {filePath} -> {iconTargetPath}");
                            }

                            if (!string.IsNullOrEmpty(iconTargetPath) && System.IO.File.Exists(iconTargetPath))
                            {
                                try
                                {
                                    newIcon = System.Drawing.Icon.ExtractAssociatedIcon(iconTargetPath).ToImageSource();
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                        $"UpdateIcon extracted target icon for Unicode shortcut {filePath}: {iconTargetPath}");
                                }
                                catch (Exception ex)
                                {
                                    newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                                        $"UpdateIcon failed to extract icon for {filePath}: {ex.Message}");
                                }
                            }
                            else
                            {
                                newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                    $"UpdateIcon using missing file icon for {filePath} (target: '{iconTargetPath}')");
                            }
                        }
                    }
                }
                //else
                //{
                //    // Non-shortcut handling
                //    if (isTargetFolder)
                //    {
                //        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                //        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Using folder-White.png for {filePath}");
                //    }
                //    else
                //    {
                //        try
                //        {
                //            newIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource();
                //            // do not remove the below commented log line, it may needed for future debugging
                //            // LogManager.Log(LogManager.LogLevel.Debug,LogManager.LogCategory.IconHandling, $"Using file icon for {filePath}");
                //        }
                //        catch (Exception ex)
                //        {
                //            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                //            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, $"Failed to extract icon for {filePath}: {ex.Message}");
                //        }
                //    }
                //}

                else
                {
                    // Non-shortcut handling with Unicode folder support

                    // Enhanced Unicode folder handling in UpdateIcon
                    bool folderExistsInUpdate = false;
                    try
                    {
                        folderExistsInUpdate = System.IO.Directory.Exists(filePath);
                        if (!folderExistsInUpdate && isFolder)
                        {
                            // Try alternative validation for Unicode folders
                            var dirInfo = new System.IO.DirectoryInfo(filePath);
                            folderExistsInUpdate = dirInfo.Exists;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                $"UpdateIcon used DirectoryInfo for Unicode folder: {filePath} -> exists: {folderExistsInUpdate}");
                        }
                    }
                    catch (Exception updateFolderEx)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"UpdateIcon folder check failed for {filePath}: {updateFolderEx.Message}");
                    }

                    if (folderExistsInUpdate || (isFolder && folderExistsInUpdate))
                    {
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"UpdateIcon using folder-White.png for existing Unicode folder {filePath}");
                    }
                    else if (isFolder)
                    {
                        // Missing folder
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"UpdateIcon using folder-WhiteX.png for missing Unicode folder {filePath}");
                    }
                    else
                    {
                        // Regular file handling
                        try
                        {
                            newIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource();
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                $"UpdateIcon using file icon for {filePath}");
                        }
                        catch (Exception ex)
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                                $"UpdateIcon failed to extract icon for {filePath}: {ex.Message}");
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
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Updated icon cache for {filePath}");
                    }

                    // do not remove the below commented log line, it may needed for future debugging
                    // LogManager.Log(LogManager.LogLevel.Debug,LogManager.LogCategory.IconHandling, $"Icon updated for {filePath}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"No icon update needed for {filePath}: same icon");
                }
            });
        }

        // Safety method to ensure no fences are stuck in transition state
        public static void ClearAllTransitionStates()
        {
            if (_fencesInTransition.Count > 0)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate, $"Clearing {_fencesInTransition.Count} stuck transition states");
                _fencesInTransition.Clear();
            }
        }

        public static void UpdateOptionsAndClickEvents()
        {


            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, $"Updating options, new singleClickToLaunch={SettingsManager.SingleClickToLaunch}");

            // Update _options
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
                CheckNetworkPaths = false // Keep this as is
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
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, $"Updated click events for {updatedItems} items");
            });
            }
            else
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, "Application.Current is null, cannot update icon.");
            }
        }




        /// <summary>
        /// Enhanced shortcut target resolution with Unicode support
        /// Handles both direct shortcuts and explorer.exe-based folder shortcuts
        /// </summary>
        private static string GetShortcutTargetUnicodeSafe(string shortcutPath)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Attempting Unicode-safe shortcut resolution for: {shortcutPath}");

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
                        if (!string.IsNullOrEmpty(folderPath) && System.IO.Directory.Exists(folderPath))
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                                $"WshShell resolved Unicode folder shortcut: {shortcutPath} -> {folderPath}");
                            return folderPath;
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                                $"Unicode folder path from arguments not found: '{folderPath}'");
                        }
                    }
                    else if (!string.IsNullOrEmpty(targetPath))
                    {
                        // Regular shortcut target
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"WshShell successfully resolved regular shortcut: {shortcutPath} -> {targetPath}");
                        return targetPath;
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                            $"WshShell returned empty TargetPath for Unicode shortcut: {shortcutPath}");
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                        $"WshShell failed for Unicode shortcut {shortcutPath}: {ex.Message}");
                }

                // Method 2: Binary parsing with enhanced Unicode folder detection
                try
                {
                    byte[] shortcutBytes = System.IO.File.ReadAllBytes(shortcutPath);
                    string targetPath = ExtractTargetFromLnkBytes(shortcutBytes, shortcutPath);

                    if (!string.IsNullOrEmpty(targetPath))
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                            $"Binary parsing successfully resolved Unicode shortcut: {shortcutPath} -> {targetPath}");
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



        /// <summary>
        /// Unicode-safe folder validation and processing
        /// Handles folders with Unicode characters in their names
        /// </summary>
        private static bool ValidateUnicodeFolderPath(string folderPath, out string sanitizedPath)
        {
            sanitizedPath = folderPath;

            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                    $"Validating Unicode folder path: {folderPath}");

                // Method 1: Direct validation
                if (System.IO.Directory.Exists(folderPath))
                {
                    // Get the full path to normalize it
                    sanitizedPath = System.IO.Path.GetFullPath(folderPath);
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                        $"Unicode folder validation successful: {folderPath} -> {sanitizedPath}");
                    return true;
                }

                // Method 2: Try with different encoding approaches
                try
                {
                    var dirInfo = new System.IO.DirectoryInfo(folderPath);
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
                return false;
            }
        }


        /// <summary>
        /// Extracts target path from .lnk file binary data
        /// Enhanced for Unicode folder shortcuts created with explorer.exe method
        /// </summary>
        private static string ExtractTargetFromLnkBytes(byte[] lnkData, string shortcutPath)
        {


            try
            {
                // LNK file format: Look for LinkInfo structure and path strings
                // Enhanced to handle Unicode folders via explorer.exe shortcuts

                if (lnkData.Length < 0x4C)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                        $"LNK file too short for parsing: {shortcutPath}");
                    return string.Empty;
                }

                // Check if this is a valid LNK file (starts with LinkHeader)
                if (lnkData[0] != 0x4C || lnkData[1] != 0x00 || lnkData[2] != 0x00 || lnkData[3] != 0x00)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                        $"Invalid LNK file header: {shortcutPath}");
                    return string.Empty;
                }

                // First, check if this is an explorer.exe shortcut (our Unicode folder shortcuts)
                string fileContentUnicode = System.Text.Encoding.Unicode.GetString(lnkData);
                string fileContentAnsi = System.Text.Encoding.Default.GetString(lnkData);

                // Look for explorer.exe followed by quoted folder path (our Unicode folder method)
                var explorerMatches = System.Text.RegularExpressions.Regex.Matches(fileContentUnicode,
                    @"explorer\.exe\s*""([^""]+)""",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                foreach (System.Text.RegularExpressions.Match match in explorerMatches)
                {
                    if (match.Groups.Count > 1)
                    {
                        string candidatePath = match.Groups[1].Value.Trim('\0', ' ', '\t', '\r', '\n');
                        if (!string.IsNullOrEmpty(candidatePath) && System.IO.Directory.Exists(candidatePath))
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                                $"Found Unicode folder path via explorer.exe method: {candidatePath}");
                            return candidatePath;
                        }
                    }
                }

                // Try ANSI encoding for explorer.exe shortcuts too
                explorerMatches = System.Text.RegularExpressions.Regex.Matches(fileContentAnsi,
             @"explorer\.exe\s*""([^""]+)""",
             System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                foreach (System.Text.RegularExpressions.Match match in explorerMatches)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
       $"Found explorer.exe match: '{match.Value}', Groups: {match.Groups.Count}");


                    if (match.Groups.Count > 1)
                    {
                        string candidatePath = match.Groups[1].Value.Trim('\0', ' ', '\t', '\r', '\n');
                        if (!string.IsNullOrEmpty(candidatePath) && System.IO.Directory.Exists(candidatePath))
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                                $"Found folder path via explorer.exe method (ANSI): {candidatePath}");
                            return candidatePath;
                        }
                    }
                }

                // Fallback: Look for regular executable paths
                var pathMatches = System.Text.RegularExpressions.Regex.Matches(fileContentUnicode,
                    @"[A-Za-z]:\\[^<>:""|?*\x00-\x1f]*\.exe",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                foreach (System.Text.RegularExpressions.Match match in pathMatches)
                {
                    string candidatePath = match.Value.Trim('\0', ' ', '\t', '\r', '\n');
                    if (System.IO.File.Exists(candidatePath))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Found valid executable path in LNK binary: {candidatePath}");
                        return candidatePath;
                    }
                }

                // Try ANSI encoding for executables
                pathMatches = System.Text.RegularExpressions.Regex.Matches(fileContentAnsi,
                    @"[A-Za-z]:\\[^<>:""|?*\x00-\x1f]*\.exe",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                foreach (System.Text.RegularExpressions.Match match in pathMatches)
                {
                    string candidatePath = match.Value.Trim('\0', ' ', '\t', '\r', '\n');
                    if (System.IO.File.Exists(candidatePath))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Found valid executable path in LNK binary (ANSI): {candidatePath}");
                        return candidatePath;
                    }
                }

                // Enhanced folder path detection for direct Unicode folder shortcuts
                var folderMatches = System.Text.RegularExpressions.Regex.Matches(fileContentUnicode,
                    @"[A-Za-z]:\\[^<>:""|?*\x00-\x1f]+(?:\\[^<>:""|?*\x00-\x1f]*)*",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                foreach (System.Text.RegularExpressions.Match match in folderMatches)
                {
                    string candidatePath = match.Value.Trim('\0', ' ', '\t', '\r', '\n');
                    if (candidatePath.Length > 5 && System.IO.Directory.Exists(candidatePath))
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                            $"Found valid Unicode folder path in LNK binary: {candidatePath}");
                        return candidatePath;
                    }
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"No valid target path found in LNK binary data: {shortcutPath}");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error extracting target from LNK binary: {ex.Message}");
                return string.Empty;
            }
        }









        public static void ClickEventAdder(StackPanel sp, string path, bool isFolder, string arguments = null)
        {
            bool isLogEnabled = _options.IsLogEnabled ?? true;

            // Store only path, isFolder, and arguments in Tag
            sp.Tag = new { FilePath = path, IsFolder = isFolder, Arguments = arguments };

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Attaching fresh handler for {path}, arguments='{arguments ?? "null"}'");

            // Clear existing handler
            sp.MouseLeftButtonDown -= MouseDownHandler;

            // Store only path, isFolder, and arguments in Tag
            sp.Tag = new { FilePath = path, IsFolder = isFolder, Arguments = arguments };

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Attaching handler for {path}, initial singleClickToLaunch={_options.singleClickToLaunch}");

            // Check if path is a shortcut and correct isFolder for folder shortcuts
            bool isShortcut = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
            string targetPath = isShortcut ? Utility.GetShortcutTarget(path) : path;
            if (isShortcut && System.IO.Directory.Exists(targetPath))
            {
                isFolder = true;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Corrected isFolder to true for shortcut {path} targeting folder {targetPath}");
            }



            void MouseDownHandler(object sender, MouseButtonEventArgs e)
            {
                if (e.ChangedButton != MouseButton.Left) return;

                // Check for Ctrl+Click to start drag operation
                if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                {
                    // Start drag operation for icon reordering
                    System.Windows.Point mousePosition = e.GetPosition(sp);
                    // StartIconDrag(sp, mousePosition);
                    IconDragDropManager.StartIconDrag(sp, mousePosition);

                    e.Handled = true; // Prevent normal click handling
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Ctrl+Click detected on {path}, starting drag operation");
                    return;
                }

                // Continue with normal click handling for launching
                bool singleClickToLaunch = _options.singleClickToLaunch ?? true;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"MouseDown on {path}, ClickCount={e.ClickCount}, singleClickToLaunch={singleClickToLaunch}");

                try
                {
                    // Recompute isShortcut to avoid unassigned variable issue
                    bool isShortcutLocal = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
                    bool targetExists;
                    string resolvedPath = path;
                    //if (isShortcutLocal)
                    //{
                    //    resolvedPath = Utility.GetShortcutTarget(path);
                    //    targetExists = isFolder ? System.IO.Directory.Exists(resolvedPath) : System.IO.File.Exists(resolvedPath);
                    //}

                    if (isShortcutLocal)
                    {
                        // Use the same Unicode-safe resolution for consistency
                        resolvedPath = GetShortcutTargetUnicodeSafe(path);

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                            $"ClickEventAdder target resolution: {path} -> {resolvedPath}");

                        if (string.IsNullOrEmpty(resolvedPath))
                        {
                            targetExists = false;
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                                $"Failed to resolve shortcut target in ClickEventAdder: {path}");
                        }
                        else
                        {
                            targetExists = isFolder ? System.IO.Directory.Exists(resolvedPath) : System.IO.File.Exists(resolvedPath);

                            // Correct isFolder if needed (folder shortcut but was marked as file)
                            if (!isFolder && System.IO.Directory.Exists(resolvedPath))
                            {
                                isFolder = true;
                                targetExists = true;
                            }
                        }
                    }
                    else
                    {
                        targetExists = isFolder ? System.IO.Directory.Exists(path) : System.IO.File.Exists(path);
                    }

                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Target check: Path={path}, ResolvedPath={resolvedPath}, IsShortcut={isShortcutLocal}, IsFolder={isFolder}, TargetExists={targetExists}");

                    if (!targetExists)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Target not found: {resolvedPath}");
                        return;
                    }

                    if (singleClickToLaunch && e.ClickCount == 1)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Single click launching {path}");
                        LaunchItem(sp, path, isFolder, arguments);
                        e.Handled = true;
                    }
                    else if (!singleClickToLaunch && e.ClickCount == 2)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Double click launching {path}");
                        LaunchItem(sp, path, isFolder, arguments);
                        e.Handled = true;
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error checking target existence: {ex.Message}");
                }
            }


            sp.MouseLeftButtonDown += MouseDownHandler;



            sp.MouseMove += (sender, e) =>
            {
                if (IconDragDropManager.IsDragging)
                {
                    try
                    {
                        // Update drag preview position to follow cursor
                        System.Windows.Point screenPosition = sp.PointToScreen(e.GetPosition(sp));
                        IconDragDropManager.HandleDragMove(screenPosition);
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error during drag move: {ex.Message}");
                    }
                }
            };

            sp.MouseLeftButtonUp += (sender, e) =>
            {
                if (IconDragDropManager.IsDragging)
                {
                    try
                    {
                        // Calculate final drop position
                        WrapPanel wrapPanel = FindVisualParent<WrapPanel>(sp);
                        if (wrapPanel != null)
                        {
                            System.Windows.Point finalPosition = e.GetPosition(wrapPanel);
                            IconDragDropManager.CompleteDrag(finalPosition);
                        }
                        else
                        {
                            IconDragDropManager.CancelDrag();
                        }
                        e.Handled = true;
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error completing drag operation: {ex.Message}");
                        IconDragDropManager.CancelDrag();
                    }
                }
            };


            // Add key up handler to cancel drag on Ctrl release
            sp.KeyUp += (sender, e) =>
            {
                if (IconDragDropManager.IsDragging &&
                    (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl) &&
                    !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Ctrl key released during drag, cancelling operation");
                    IconDragDropManager.CancelDrag();
                    e.Handled = true;
                }
            };
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


                LaunchEffectsManager.LaunchEffect defaultEffect = (LaunchEffectsManager.LaunchEffect)_options.LaunchEffect;
                LaunchEffectsManager.LaunchEffect effect;

                if (fence == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Failed to find fence for {path}, using default effect");
                    effect = defaultEffect;
                }
                else if (string.IsNullOrEmpty(customEffect))
                {
                    effect = defaultEffect;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"No CustomLaunchEffect for {path} in fence '{fence.Title}', using default: {effect}");
                }
                else
                {

                    try
                    {
                        effect = (LaunchEffectsManager.LaunchEffect)Enum.Parse(typeof(LaunchEffectsManager.LaunchEffect), customEffect, true);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Using CustomLaunchEffect for {path} in fence '{fence.Title}': {effect}");
                    }
                    catch (Exception ex)
                    {
                        effect = defaultEffect;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Failed to parse CustomLaunchEffect '{customEffect}' for {path} in fence '{fence.Title}', falling back to {effect}: {ex.Message}");
                    }
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"LaunchItem called for {path} with effect {effect}, providedArguments='{arguments ?? "null"}'");


                // Execute launch effect animation using LaunchEffectsManager
                LaunchEffectsManager.ExecuteLaunchEffect(sp, (LaunchEffectsManager.LaunchEffect)effect);


                // Enhanced launch logic with proper argument handling
                string extension = System.IO.Path.GetExtension(path).ToLower();
                bool isLnkShortcut = extension == ".lnk";
                bool isUrlFile = extension == ".url";
                bool isShortcut = isLnkShortcut || isUrlFile;
                string targetPath = path;
                string finalArguments = arguments ?? "";
                string workingDirectory = "";

                if (isShortcut)
                {
                    if (!System.IO.File.Exists(path))
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Shortcut file not found: {path}");
                        TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Shortcut file not found: {path}", "Error");
                        return;
                    }

                    try
                    {
                        if (isUrlFile)
                        {
                            // Handle .url files
                            targetPath = ExtractUrlFromUrlFile(path);
                            if (string.IsNullOrEmpty(targetPath))
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Could not extract URL from .url file: {path}");
                                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Invalid .url file: {path}", "Error");
                                return;
                            }

                            // .url files don't typically have arguments, but use provided ones if available
                            finalArguments = arguments ?? "";
                            workingDirectory = ""; // Not applicable for URLs

                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"URL file resolved: '{path}' -> URL='{targetPath}'");
                        }


                        else if (isLnkShortcut)
                        {
                            // Handle .lnk files with Unicode support
                            string resolvedTarget = GetShortcutTargetUnicodeSafe(path);

                            if (!string.IsNullOrEmpty(resolvedTarget))
                            {
                                // Successfully resolved with Unicode-safe method
                                targetPath = resolvedTarget;
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                    $"LaunchItem Unicode resolution: {path} -> {targetPath}");

                                // Enhanced argument and working directory handling for Unicode shortcuts
                                try
                                {
                                    WshShell shell = new WshShell();
                                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(path);
                                    string shortcutTarget = shortcut.TargetPath?.Trim();
                                    string shortcutArgs = shortcut.Arguments?.Trim();

                                    // Check if this is a Unicode folder shortcut (explorer.exe method)
                                    if (!string.IsNullOrEmpty(shortcutTarget) &&
                                        shortcutTarget.ToLower().EndsWith("explorer.exe") &&
                                        !string.IsNullOrEmpty(shortcutArgs))
                                    {
                                        // This is our Unicode folder shortcut - don't pass explorer arguments to launch
                                        workingDirectory = System.IO.Path.GetDirectoryName(targetPath) ?? "";
                                        finalArguments = arguments ?? ""; // Use provided arguments, not explorer args

                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                            $"Detected Unicode folder shortcut via explorer.exe method: {targetPath}");
                                    }
                                    else
                                    {
                                        // Regular shortcut handling
                                        workingDirectory = shortcut.WorkingDirectory ?? "";

                                        // Use provided arguments if available, otherwise try shortcut's arguments
                                        if (string.IsNullOrEmpty(arguments))
                                        {
                                            // Only use shortcut arguments if target paths match (not explorer.exe case)
                                            if (!string.IsNullOrEmpty(shortcutTarget) &&
                                                shortcutTarget.Equals(targetPath, StringComparison.OrdinalIgnoreCase))
                                            {
                                                finalArguments = shortcutArgs ?? "";
                                            }
                                            else
                                            {
                                                finalArguments = "";
                                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                                    $"Target mismatch, using empty arguments for shortcut: {path}");
                                            }
                                        }
                                        else
                                        {
                                            finalArguments = arguments;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Fallback for working directory - handle folders properly
                                    if (System.IO.Directory.Exists(targetPath))
                                    {
                                        workingDirectory = targetPath; // For folders, working directory is the folder itself
                                    }
                                    else if (System.IO.File.Exists(targetPath))
                                    {
                                        workingDirectory = System.IO.Path.GetDirectoryName(targetPath) ?? "";
                                    }
                                    else
                                    {
                                        workingDirectory = "";
                                    }

                                    finalArguments = arguments ?? "";
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                        $"WshShell failed for working directory, using fallback for {path}: {ex.Message}");
                                }
                            }
                            else
                            {
                                // Fallback to original WshShell method
                                try
                                {
                                    WshShell shell = new WshShell();
                                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(path);
                                    targetPath = shortcut.TargetPath ?? "";
                                    workingDirectory = shortcut.WorkingDirectory ?? "";

                                    if (string.IsNullOrEmpty(arguments))
                                    {
                                        finalArguments = shortcut.Arguments ?? "";
                                    }
                                    else
                                    {
                                        finalArguments = arguments;
                                    }

                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                        $"Fallback WshShell method used for LaunchItem: {path} -> {targetPath}");
                                }
                                catch (Exception ex)
                                {
                                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                                        $"Both Unicode-safe and WshShell methods failed for {path}: {ex.Message}");
                                    targetPath = "";
                                    workingDirectory = "";
                                    finalArguments = arguments ?? "";
                                }
                            }

                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                $"LNK shortcut resolved: '{path}' -> target='{targetPath}', args='{finalArguments}'");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error reading shortcut {path}: {ex.Message}");
                        TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Error reading shortcut: {ex.Message}", "Error");
                        return;
                    }
                }
                else
                {
                    // For direct files, set working directory
                    if (System.IO.File.Exists(targetPath))
                    {
                        try
                        {
                            workingDirectory = System.IO.Path.GetDirectoryName(targetPath) ?? "";
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Could not determine working directory for {targetPath}: {ex.Message}");
                            workingDirectory = "";
                        }
                    }
                }

                // Check if target exists
                bool isTargetFolder = System.IO.Directory.Exists(targetPath);
                bool targetExists = System.IO.File.Exists(targetPath) || isTargetFolder;
                bool isUrl = IsWebUrl(targetPath);
                bool isSpecialPath = IsSpecialWindowsPath(targetPath);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Target analysis: path='{targetPath}', exists={targetExists}, isFolder={isTargetFolder}, isUrl={isUrl}, isSpecial={isSpecialPath}");

                if (!targetExists && !isUrl && !isSpecialPath)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Target not found: {targetPath}");
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Target '{targetPath}' was not found.", "Error");
                    return;
                }

                // Create and configure ProcessStartInfo
                ProcessStartInfo psi = new ProcessStartInfo();

                if (isUrl)
                {
                    // Handle URLs - don't add arguments to URLs
                    psi.FileName = targetPath;
                    psi.UseShellExecute = true;
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Launching URL: {targetPath}");
                }
                else if (isSpecialPath)
                {
                    // Handle special Windows paths
                    psi.FileName = targetPath;
                    psi.UseShellExecute = true;
                    if (!string.IsNullOrEmpty(finalArguments))
                    {
                        psi.Arguments = finalArguments;
                    }
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Launching special path: {targetPath} with args: '{finalArguments}'");
                }
                else if (isTargetFolder)
                {
                    // Handle folders
                    psi.FileName = "explorer.exe";
                    psi.Arguments = $"\"{targetPath}\"";
                    psi.UseShellExecute = true;
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Opening folder: {targetPath}");
                }
                else
                {
                    // Handle regular executables
                    psi.FileName = targetPath;
                    psi.UseShellExecute = true;

                    if (!string.IsNullOrEmpty(finalArguments))
                    {
                        psi.Arguments = finalArguments;
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Launching: '{targetPath}' with arguments: '{finalArguments}'");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Launching: '{targetPath}' (no arguments)");
                    }

                    // Set working directory if available
                    if (!string.IsNullOrEmpty(workingDirectory) && System.IO.Directory.Exists(workingDirectory))
                    {
                        psi.WorkingDirectory = workingDirectory;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Set working directory: {workingDirectory}");
                    }
                }

                // Execute the process
                try
                {
                    Process process = Process.Start(psi);
                    if (process != null)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Successfully started process for: {targetPath}");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Process.Start returned null for: {targetPath}");
                        TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to start: {targetPath}", "Launch Error");
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Exception starting process: {ex.Message}\nTarget: {targetPath}\nArguments: {finalArguments}");
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Error launching: {ex.Message}", "Launch Error");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Critical error in LaunchItem for {path}: {ex.Message}");
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Error opening item: {ex.Message}", "Error");
            }
        }




        // Determines if a path is a web URL

        private static bool IsWebUrl(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;

            return path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase);
        }


        // Determines if a path is a special Windows path

        private static bool IsSpecialWindowsPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;

            return path.StartsWith("shell:", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("ms-settings:", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("ms-", StringComparison.OrdinalIgnoreCase) ||
                   path.Contains("::") ||
                   path.StartsWith("control", StringComparison.OrdinalIgnoreCase);
        }


        // Extracts the URL from a .url file

        private static string ExtractUrlFromUrlFile(string urlFilePath)
        {
            try
            {
                if (!System.IO.File.Exists(urlFilePath))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"URL file not found: {urlFilePath}");
                    return null;
                }

                string[] lines = System.IO.File.ReadAllLines(urlFilePath);
                foreach (string line in lines)
                {
                    if (line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                    {
                        string url = line.Substring(4).Trim(); // Remove "URL=" prefix
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Extracted URL from {urlFilePath}: {url}");
                        return url;
                    }
                }

                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"No URL= line found in .url file: {urlFilePath}");
                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error reading .url file {urlFilePath}: {ex.Message}");
                return null;
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

        // Public method to refresh an icon's click handlers after shortcut editing
        // Called by EditShortcutWindow to ensure immediate argument updates

        public static void RefreshIconClickHandlers(string shortcutPath, string newDisplayName)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Refreshing click handlers for edited shortcut: {shortcutPath}");

                // Find all fence windows and locate the icon
                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                bool iconFound = false;

                foreach (var window in windows)
                {
                    // Find the WrapPanel containing icons
                    var wrapPanel = FindWrapPanel(window);
                    if (wrapPanel == null) continue;

                    // Find the specific icon StackPanel
                    foreach (StackPanel iconPanel in wrapPanel.Children.OfType<StackPanel>())
                    {
                        var tagData = iconPanel.Tag;
                        if (tagData != null)
                        {
                            string filePath = tagData.GetType().GetProperty("FilePath")?.GetValue(tagData)?.ToString();
                            if (!string.IsNullOrEmpty(filePath) &&
                                string.Equals(System.IO.Path.GetFullPath(filePath), System.IO.Path.GetFullPath(shortcutPath), StringComparison.OrdinalIgnoreCase))
                            {
                                // Found it! Refresh this icon completely
                                RefreshSingleIconComplete(iconPanel, shortcutPath, newDisplayName, window);
                                iconFound = true;
                                break;
                            }
                        }
                    }
                    if (iconFound) break;
                }

                if (!iconFound)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                        $"Could not find icon to refresh for: {shortcutPath}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error refreshing icon click handlers: {ex.Message}");
            }
        }


        // Completely refreshes a single icon with fresh data from the .lnk fil

        private static void RefreshSingleIconComplete(StackPanel iconPanel, string shortcutPath, string newDisplayName, NonActivatingWindow parentWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Performing complete refresh of icon: {shortcutPath}");

                // Read fresh data from the updated shortcut file
                string freshTargetPath = shortcutPath;
                string freshArguments = "";
                bool isFolder = false;
                string workingDirectory = "";

                if (System.IO.Path.GetExtension(shortcutPath).ToLower() == ".lnk")
                {
                    try
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                        freshTargetPath = shortcut.TargetPath ?? shortcutPath;
                        freshArguments = shortcut.Arguments ?? "";
                        workingDirectory = shortcut.WorkingDirectory ?? "";
                        isFolder = System.IO.Directory.Exists(freshTargetPath);

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Fresh data loaded - Target: '{freshTargetPath}', Args: '{freshArguments}', IsFolder: {isFolder}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error reading shortcut: {ex.Message}");
                        return;
                    }
                }

                //  Update the Tag with completely fresh data (MOST IMPORTANT)
                iconPanel.Tag = new
                {
                    FilePath = shortcutPath,
                    IsFolder = isFolder,
                    Arguments = freshArguments
                };

                // 2. Update display name if provided
                if (!string.IsNullOrEmpty(newDisplayName))
                {
                    var textBlock = iconPanel.Children.OfType<TextBlock>().FirstOrDefault();
                    if (textBlock != null)
                    {
                        string displayText = newDisplayName.Length > 20 ? newDisplayName.Substring(0, 20) + "..." : newDisplayName;
                        textBlock.Text = displayText;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Updated display text: '{displayText}'");
                    }
                }

                //  Update tooltip with fresh information
                string toolTipText = $"{System.IO.Path.GetFileName(shortcutPath)}\nTarget: {freshTargetPath}";
                if (!string.IsNullOrEmpty(freshArguments))
                {
                    toolTipText += $"\nArguments: {freshArguments}";
                }
                if (!string.IsNullOrEmpty(workingDirectory))
                {
                    toolTipText += $"\nWorking Directory: {workingDirectory}";
                }
                iconPanel.ToolTip = new ToolTip { Content = toolTipText };

                // CRITICAL: Re-attach fresh event handlers
                ClickEventAdder(iconPanel, shortcutPath, isFolder, freshArguments);

                //  Update fence JSON data
                UpdateFenceDataForIcon(shortcutPath, newDisplayName, parentWindow);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Complete icon refresh successful. New arguments: '{freshArguments}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in complete icon refresh: {ex.Message}");
            }
        }

        /// Simplified approach to handle old event handlers

        private static void ClearIconEventHandlers(StackPanel iconPanel)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Preparing to refresh event handlers (simplified approach)");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Event handler preparation: {ex.Message}");
            }
        }



        // Updates fence JSON data after icon refresh

        private static void UpdateFenceDataForIcon(string shortcutPath, string newDisplayName, NonActivatingWindow parentWindow)
        {
            try
            {
                if (string.IsNullOrEmpty(newDisplayName)) return;

                string fenceId = parentWindow.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                var fence = _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (fence?.ItemsType?.ToString() != "Data") return;

                var items = fence.Items as JArray;
                if (items == null) return;

                foreach (var item in items)
                {
                    string itemFilename = item["Filename"]?.ToString();
                    if (!string.IsNullOrEmpty(itemFilename) &&
                        string.Equals(System.IO.Path.GetFullPath(itemFilename), System.IO.Path.GetFullPath(shortcutPath), StringComparison.OrdinalIgnoreCase))
                    {
                        item["DisplayName"] = newDisplayName;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Updated JSON data for: {itemFilename}");
                        break;
                    }
                }

                SaveFenceData();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error updating fence JSON: {ex.Message}");
            }
        }












    }



}