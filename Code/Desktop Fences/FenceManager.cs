using IWshRuntimeLibrary;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics.Arm;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using static System.Net.Mime.MediaTypeNames;
using Color = System.Drawing.Color;

namespace Desktop_Fences
{
    public static class FenceManager
    {

        // --- NEW: Centralized Presets ---
        private static readonly Dictionary<string, string> _standardPresets = new Dictionary<string, string>
        {
            { "Images", "*.jpg; *.jpeg; *.png; *.gif; *.bmp; *.webp" },
            { "Documents", "*.doc*; *.pdf; *.txt; *.rtf; *.xls*; *.ppt*" },
            { "Executables", "*.exe; *.bat; *.msi; *.cmd; *.ps1" },
            { "Archives", "*.zip; *.rar; *.7z; *.tar; *.gz" },
            { "Media", "*.mp3; *.wav; *.mp4; *.mkv; *.avi" },
            { "Hide System", ">*.tmp; >desktop.ini; >~$*" }
        };
        // -------------------------------


        // --- WM_GETMINMAXINFO Implementation ---
        private const int WM_GETMINMAXINFO = 0x0024;

        private const int WM_SYSCOMMAND = 0x0112; // <--- NEW
        private const int SC_MAXIMIZE = 0xF030;   // <--- NEW

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MINMAXINFO
        {
            public POINT ptReserved;
            public POINT ptMaxSize;
            public POINT ptMaxPosition;
            public POINT ptMinTrackSize;
            public POINT ptMaxTrackSize;
        }


        // Tracks temporary navigation paths for Portal Fences. Key: FenceID, Value: CurrentPath
        private static Dictionary<string, string> _portalNavigationStates = new Dictionary<string, string>();

        // Track icon state to prevent GDI leaks from constant re-extraction
        // Key: FilePath
        // Value: (LastWriteTime of the .lnk/file, IsBroken state)
        private static Dictionary<string, (DateTime LastWrite, bool IsBroken)> _iconStates = new Dictionary<string, (DateTime, bool)>();

        private static dynamic _options;
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
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Reloading all fences...");

                // 1. Life Support: Create a temporary window to prevent app shutdown
                // when we close all the fences.
                Window lifeSupport = MessageBoxesManager.CreateWaitWindow("Desktop Fences +", "Refreshing configuration...");
                lifeSupport.Show();

                // Allow UI to render the wait window
                System.Windows.Forms.Application.DoEvents();

                // 2. Close All Existing Fence Windows
                var openFences = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
                foreach (var win in openFences)
                {
                    win.Close();
                }

                // 3. Clear Static References & Memory
                _heartTextBlocks.Clear();

                foreach (var portal in _portalFences.Values)
                {
                    try { portal.Dispose(); } catch { }
                }
                _portalFences.Clear();

                FenceDataManager.FenceData?.Clear();

                _currentTargetChecker?.Stop();

                // 4. Re-Initialize
                _currentTargetChecker = new TargetChecker(1000);

                // Reload settings (in case they changed during reset) and create new fences
                LoadAndCreateFences(_currentTargetChecker);

                // 5. Remove Life Support
                lifeSupport.Close();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Fences reloaded successfully.");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error reloading fences: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error reloading fences: {ex.Message}", "Error");
            }
        }



        /// <summary>
        /// Helper to detect UNC Roots (e.g. \\Server or \\192.168.1.10).
        /// These technically aren't "Directories" in .NET, but act as Folders in Windows.
        /// </summary>
        private static bool IsUncRoot(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;

            // Must start with \\ and NOT contain a drive letter (like C:\)
            if (path.StartsWith(@"\\") && !path.Contains(@":\"))
            {
                string clean = path.Substring(2);
                // If no more slashes (Server) or just one trailing slash (Server\), it's a root.
                int slash = clean.IndexOf('\\');
                return (slash < 0 || slash == clean.Length - 1);
            }
            return false;
        }

        // PORTAL FEATURE: Logic to switch folders
        public static void NavigatePortalFence(dynamic fence, string newPath)
        {
            try
            {
                if (string.IsNullOrEmpty(newPath) || !System.IO.Directory.Exists(newPath)) return;

                string fenceId = fence.Id?.ToString();
                string basePath = fence.Path?.ToString(); // The permanent "Home"

                // Update State
                if (string.Equals(newPath, basePath, StringComparison.OrdinalIgnoreCase))
                {
                    _portalNavigationStates.Remove(fenceId); // We returned Home
                }
                else
                {
                    _portalNavigationStates[fenceId] = newPath; // We are Deep
                }

                // Tell the Manager to switch
                if (_portalFences.ContainsKey(fence))
                {
                    _portalFences[fence].NavigateTo(newPath);
                }

                // Update UI (Show/Hide Bar)
                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                var win = windows.FirstOrDefault(w => w.Tag?.ToString() == fenceId);
                if (win != null)
                {
                    RefreshPortalNavBar(win, fence);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Portal Navigation Error: {ex.Message}");
            }
        }

        // PORTAL FEATURE: The UI Strip [ <  Path...  ⚓ ]
        // v2.5.4 .186: Precision Alignment (Anchor under Filter)
        public static void RefreshPortalNavBar(NonActivatingWindow fenceWindow, dynamic fence)
        {
            try
            {
                var border = fenceWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                // 1. Cleanup Old Bar
                var oldBar = dockPanel.Children.OfType<Grid>()
                    .FirstOrDefault(g => g.Tag?.ToString() == "PORTAL_NAV_BAR");
                if (oldBar != null) dockPanel.Children.Remove(oldBar);

                // 2. Check Logic
                if (fence.ItemsType?.ToString() != "Portal") return;

                string fenceId = fence.Id?.ToString();
                string basePath = fence.Path?.ToString();
                string currentPath = _portalNavigationStates.ContainsKey(fenceId) ? _portalNavigationStates[fenceId] : basePath;

                // Only show bar if we are NOT at home
                bool isNavigating = !string.Equals(currentPath, basePath, StringComparison.OrdinalIgnoreCase);
                if (!isNavigating) return;

                // Path Logic
                string displayPath = new System.IO.DirectoryInfo(currentPath).Name;
                try
                {
                    if (currentPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
                    {
                        string rootName = new System.IO.DirectoryInfo(basePath).Name;
                        string relativePart = currentPath.Substring(basePath.Length).TrimStart(System.IO.Path.DirectorySeparatorChar);
                        displayPath = string.IsNullOrEmpty(relativePart) ? rootName : System.IO.Path.Combine(rootName, relativePart);
                    }
                }
                catch { }

                // 3. Create UI (Styled & Aligned)
                Grid navGrid = new Grid
                {
                    Tag = "PORTAL_NAV_BAR",
                    Height = 24,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(60, 0, 0, 0)),
                    // FIX: Right Margin 35 (30px Lock + 5px Gap) aligns the bar perfectly under the Filter Icon
                    Margin = new Thickness(5, 0, 35, 2)
                };

                navGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(25) }); // Back Arrow
                navGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Path Text
                navGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(25) }); // Anchor Icon

                // [ < ] Back Button
                Button btnBack = new Button
                {
                    Content = "‹",
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = System.Windows.Media.Brushes.Transparent,
                    Foreground = System.Windows.Media.Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    ToolTip = "Go Up"
                };
                btnBack.Click += (s, e) => NavigatePortalFence(fence, System.IO.Directory.GetParent(currentPath)?.FullName);

                // [ Path ] Label (Left Aligned)
                TextBlock lblPath = new TextBlock
                {
                    Text = displayPath,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = currentPath,
                    Margin = new Thickness(5, 0, 5, 0)
                };

                // [ ⚓ ] Anchor Button
                Button btnSetBase = new Button
                {
                    Content = "⚓",
                    FontSize = 10,
                    Background = System.Windows.Media.Brushes.Transparent,
                    Foreground = System.Windows.Media.Brushes.Orange,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    ToolTip = "Set current view as the new Home for this fence"
                };
                btnSetBase.Click += (s, e) =>
                {
                    if (MessageBoxesManager.ShowCustomYesNoMessageBox($"Set this folder as the new home?\n\n{currentPath}", "Update Portal Fence"))
                    {
                        UpdateFenceProperty(fence, "Path", currentPath, "Updated Portal Path");
                        _portalNavigationStates.Remove(fenceId);
                        RefreshPortalNavBar(fenceWindow, fence);
                    }
                };

                navGrid.Children.Add(btnBack); Grid.SetColumn(btnBack, 0);
                navGrid.Children.Add(lblPath); Grid.SetColumn(lblPath, 1);
                navGrid.Children.Add(btnSetBase); Grid.SetColumn(btnSetBase, 2);

                // 4. Smart Insertion (Under Title)
                DockPanel.SetDock(navGrid, Dock.Top);

                int insertIndex = 0;
                bool titleFound = false;

                for (int i = 0; i < dockPanel.Children.Count; i++)
                {
                    if (dockPanel.Children[i] is Grid g &&
                        g.Children.OfType<TextBlock>().Any(tb => tb.Name == "FenceLockIcon"))
                    {
                        insertIndex = i + 1;
                        titleFound = true;
                        break;
                    }
                }

                if (!titleFound) insertIndex = 1;
                if (insertIndex > dockPanel.Children.Count) insertIndex = dockPanel.Children.Count;

                dockPanel.Children.Insert(insertIndex, navGrid);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing Nav Bar: {ex.Message}");
            }
        }
        /// <summary>
        /// Surgically reloads only specific fences by closing and recreating their windows
        /// Used by: ItemMoveDialog to prevent duplicate windows during move operations
        /// Category: Selective Window Management
        /// </summary>
        public static void ReloadSpecificFences(dynamic sourceFence, dynamic targetFence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Reloading specific fences: '{sourceFence?.Title}' and '{targetFence?.Title}'");

                // Get fence IDs for identification
                string sourceFenceId = sourceFence?.Id?.ToString();
                string targetFenceId = targetFence?.Id?.ToString();

                // Check if we're about to close all fences (which would terminate the app)
                bool isSourceTargetSame = sourceFenceId == targetFenceId;
                int totalFenceCount = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().Count();
                int fencesToClose = isSourceTargetSame ? 1 : 2;
                bool needsLifeSupport = (fencesToClose >= totalFenceCount);

                Window lifeSupportWindow = null;

                // Show "life support" wait window if we're about to close all fences
                if (needsLifeSupport)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Creating life support window - about to close {fencesToClose} of {totalFenceCount} total fences");

                    lifeSupportWindow = MessageBoxesManager.CreateWaitWindow("Desktop Fences +", "Refreshing fences, please wait...");
                    lifeSupportWindow.Show();
                }

                // Find and close only the specific fence windows
                var allWindows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
                var windowsToClose = new List<NonActivatingWindow>();

                foreach (var window in allWindows)
                {
                    string windowFenceId = window.Tag?.ToString();
                    if (!string.IsNullOrEmpty(windowFenceId) &&
                        (windowFenceId == sourceFenceId || windowFenceId == targetFenceId))
                    {
                        windowsToClose.Add(window);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"Marking fence window for reload: {window.Title} (ID: {windowFenceId})");
                    }
                }

                // Close the specific windows
                foreach (var window in windowsToClose)
                {
                    try
                    {
                        window.Close();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"Closed fence window: {window.Title}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                            $"Error closing window {window.Title}: {ex.Message}");
                    }
                }

                // Clean up portal fence managers for the specific fences
                var portalFencesToRemove = _portalFences.Keys
                    .Where(fence => fence?.Id?.ToString() == sourceFenceId || fence?.Id?.ToString() == targetFenceId)
                    .ToList();

                foreach (var fence in portalFencesToRemove)
                {
                    _portalFences.Remove(fence);
                }

                // Recreate only the specific fences with proper TargetChecker
                var fenceData = FenceDataManager.FenceData;
                var fencesToRecreate = fenceData.Where(f =>
                    f.Id?.ToString() == sourceFenceId || f.Id?.ToString() == targetFenceId).ToList();

                foreach (var fence in fencesToRecreate)
                {
                    try
                    {
                        CreateFence(fence, _currentTargetChecker ?? new TargetChecker(1000));
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                            $"Recreated fence: {fence.Title}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                            $"Error recreating fence {fence.Title}: {ex.Message}");
                    }
                }

                // Close life support window if we used it
                if (lifeSupportWindow != null)
                {
                    try
                    {
                        lifeSupportWindow.Close();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            "Closed life support window");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                            $"Error closing life support window: {ex.Message}");
                    }
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    "Specific fence reload completed successfully");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in ReloadSpecificFences: {ex.Message}");
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
            FenceDataManager.SaveFenceData();
  
        }

        private static int _registryMonitorTickCount = 0;







        // Builds the heart ContextMenu for a fence with consistent items and dynamic state
        // v2.5.4 .183: Swapped Tabs/Delete position for better UX safety
        private static ContextMenu BuildHeartContextMenu(dynamic fence, bool showTabsOption = false)
        {
            // DEBUG: Log menu building
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                $"Building heart context menu for fence '{fence.Title}'");

            var menu = new ContextMenu();

            // --- AUTO-CLOSE TIMER ---
            System.Windows.Threading.DispatcherTimer menuTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(4)
            };

            menuTimer.Tick += (s, e) =>
            {
                if (menu.IsOpen && !menu.IsMouseOver)
                {
                    menu.IsOpen = false;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Heart Menu auto-closed by timer");
                }
                menuTimer.Stop();
            };

            menu.Opened += (s, e) => { menuTimer.Start(); };
            menu.Closed += (s, e) => menuTimer.Stop();
            menu.MouseEnter += (s, e) => menuTimer.Stop();
            menu.MouseLeave += (s, e) => { menuTimer.Stop(); menuTimer.Start(); };
            // ------------------------

            // About item
            var aboutItem = new MenuItem { Header = "About..." };
            aboutItem.Click += (s, e) => AboutFormManager.ShowAboutForm();
            menu.Items.Add(aboutItem);

            // Options item
            var optionsItem = new MenuItem { Header = "Options..." };
            optionsItem.Click += (s, e) => OptionsFormManager.ShowOptionsForm();
            menu.Items.Add(optionsItem);

            // Separator
            menu.Items.Add(new Separator());

            // New Fence items
            var newFenceItem = new MenuItem { Header = "New Fence" };
            newFenceItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                CreateNewFence("", "Data", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newFenceItem);

            var newPortalFenceItem = new MenuItem { Header = "New Portal Fence" };
            newPortalFenceItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                CreateNewFence("New Portal Fence", "Portal", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newPortalFenceItem);

            MenuItem newNoteFenceItem = new MenuItem { Header = "New Note Fence" };
            newNoteFenceItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                CreateNewFence("", "Note", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newNoteFenceItem);

            menu.Items.Add(new Separator());

            // --- REORDERED: Tabs Option First ---
            // TABS FEATURE (Native Checkbox Logic)
            bool isDataFence = fence.ItemsType?.ToString() == "Data";
            if (isDataFence)
            {
                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

                var enableTabsItem = new MenuItem
                {
                    Header = "Enable Tabs On This Fence",
                    IsCheckable = true,   // Shows checkbox gutter
                    IsChecked = tabsEnabled // Visual checkmark
                };

                enableTabsItem.Click += (s, e) => ToggleFenceTabs(fence);
                menu.Items.Add(enableTabsItem);

                // Separator AFTER tabs to separate it from the Delete option
                menu.Items.Add(new Separator());
            }

            // --- REORDERED: Delete Option Second ---
            // Delete this fence
            var deleteThisFence = new MenuItem { Header = "Delete this Fence" };
            deleteThisFence.Click += (s, e) =>
            {
                bool result = MessageBoxesManager.ShowCustomMessageBoxForm();
                if (result == true)
                {
                    if (SettingsManager.ExportShortcutsOnFenceDeletion && fence.ItemsType?.ToString() == "Data")
                    {
                        ExportAllIconsToDesktop(fence, false);
                    }

                    BackupManager.BackupDeletedFence(fence);

                    FenceDataManager.FenceData.Remove(fence);
                    _heartTextBlocks.Remove(fence);

                    if (_portalFences.ContainsKey(fence))
                    {
                        _portalFences[fence].Dispose();
                        _portalFences.Remove(fence);
                    }

                    FenceDataManager.SaveFenceData();

                    var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                    var win = windows.FirstOrDefault(w => w.Tag?.ToString() == fence.Id?.ToString());
                    if (win != null) win.Close();

                    UpdateAllHeartContextMenus();
                }
            };
            menu.Items.Add(deleteThisFence);

            menu.Items.Add(new Separator());

            // Export/Import Group
            var exportItem = new MenuItem { Header = "Export this Fence" };
            exportItem.Click += (s, e) => BackupManager.ExportFence(fence);
            menu.Items.Add(exportItem);

            var importItem = new MenuItem { Header = "Import a Fence..." };
            importItem.Click += (s, e) => BackupManager.ImportFence();
            menu.Items.Add(importItem);

            // Restore Fence item
            var restoreItem = new MenuItem
            {
                Header = "Restore Last Deleted Fence",
                Visibility = BackupManager.IsRestoreAvailable ? Visibility.Visible : Visibility.Collapsed
            };
            restoreItem.Click += (s, e) => BackupManager.RestoreLastDeletedFence();
            menu.Items.Add(restoreItem);

            // Separator
            menu.Items.Add(new Separator());

            // Exit item
            var exitItem = new MenuItem { Header = "Exit" };
            exitItem.Click += (s, e) => System.Windows.Application.Current.Shutdown();
            menu.Items.Add(exitItem);

            return menu;
        }


        /// <summary>
        /// Centralized method to attach the standard Context Menu to an icon.
        /// Layout: Edit/Move/Remove -> Copy -> Admin -> Path
        /// Includes LIVE data lookup to ensure Checkable items stay synced.
        /// </summary>
        public static void AttachIconContextMenu(StackPanel sp, dynamic item, dynamic fence, NonActivatingWindow window)
        {
            try
            {
                // Initial snapshot for static properties (Path, Folder status)
                IDictionary<string, object> iconDict = item is IDictionary<string, object> dict ?
                    dict : ((JObject)item).ToObject<IDictionary<string, object>>();

                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];

                // HELPER: Function to find the LIVE item in memory (Handles Tabs vs Main)
                // We need this because 'item' becomes stale after an update.
                Func<dynamic> GetLiveItem = () =>
                {
                    try
                    {
                        string fenceId = fence.Id?.ToString();
                        var liveFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                        if (liveFence == null) return null;

                        JArray itemsArray = liveFence.Items as JArray ?? new JArray();

                        // Handle Tabs
                        bool tabsEnabled = liveFence.TabsEnabled?.ToString().ToLower() == "true";
                        if (tabsEnabled)
                        {
                            var tabs = liveFence.Tabs as JArray ?? new JArray();
                            int currentTabIndex = Convert.ToInt32(liveFence.CurrentTab?.ToString() ?? "0");
                            if (currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                            {
                                var activeTab = tabs[currentTabIndex] as JObject;
                                itemsArray = activeTab?["Items"] as JArray ?? itemsArray;
                            }
                        }

                        // Find specific item by filename
                        return itemsArray.FirstOrDefault(i => string.Equals(
                            System.IO.Path.GetFullPath(i["Filename"]?.ToString() ?? ""),
                            System.IO.Path.GetFullPath(filePath),
                            StringComparison.OrdinalIgnoreCase));
                    }
                    catch { return null; }
                };

                ContextMenu iconContextMenu = new ContextMenu();

                // --- GROUP 1: MANIPULATION ---
                MenuItem miEdit = new MenuItem { Header = "Edit..." };
                MenuItem miMove = new MenuItem { Header = "Move..." };
                MenuItem miRemove = new MenuItem { Header = "Remove" };

                iconContextMenu.Items.Add(miEdit);
                iconContextMenu.Items.Add(miMove);
                iconContextMenu.Items.Add(miRemove);

                // --- GROUP 2: CLIPBOARD ---
                iconContextMenu.Items.Add(new Separator());

                MenuItem miCopyItem = new MenuItem { Header = "Copy Item" };
                iconContextMenu.Items.Add(miCopyItem);

                // --- GROUP 3: EXECUTION (Conditional) ---
                bool isEligibleForAdmin = !isFolder && !string.IsNullOrEmpty(filePath) &&
                    (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk" || System.IO.Path.GetExtension(filePath).ToLower() == ".exe");

                MenuItem miAlwaysAdmin = null;

                if (isEligibleForAdmin)
                {
                    iconContextMenu.Items.Add(new Separator());

                    MenuItem miRunAsAdmin = new MenuItem { Header = "Run as administrator" };
                    miAlwaysAdmin = new MenuItem
                    {
                        Header = "Always run as administrator",
                        IsCheckable = true
                    };

                    iconContextMenu.Items.Add(miRunAsAdmin);
                    iconContextMenu.Items.Add(miAlwaysAdmin);

                 
                    // Run as Admin Logic
                    miRunAsAdmin.Click += (s, e) => {
                        string target = Utility.GetShortcutTarget(filePath);
                        string args = Utility.GetShortcutArguments(filePath); // Extract arguments

                        ProcessStartInfo psi = new ProcessStartInfo { FileName = target, UseShellExecute = true, Verb = "runas" };

                        if (!string.IsNullOrEmpty(args))
                        {
                            psi.Arguments = args; // Pass arguments to admin launch
                        }

                        if (System.IO.File.Exists(target))
                        {
                            string wd = System.IO.Path.GetDirectoryName(target);
                            if (!string.IsNullOrEmpty(wd)) psi.WorkingDirectory = wd;
                        }
                        try { Process.Start(psi); }
                        catch (Exception ex)
                        {
                            MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error running as admin: {ex.Message}", "Error");
                        }
                    };

                    // Toggle Always Admin Logic
                    miAlwaysAdmin.Click += (sender, e) => {
                        var liveItem = GetLiveItem();
                        if (liveItem != null)
                        {
                            bool newVal = !Convert.ToBoolean(liveItem["AlwaysRunAsAdmin"] ?? false);
                            liveItem["AlwaysRunAsAdmin"] = newVal;

                            // Visual feedback immediately
                            miAlwaysAdmin.IsChecked = newVal;

                            FenceDataManager.SaveFenceData();
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Toggled AlwaysRunAsAdmin to {newVal} for {filePath}");
                        }
                    };
                }

                // --- GROUP 4: PATH / TARGET ---
                iconContextMenu.Items.Add(new Separator());

                MenuItem miCopyPathRoot = new MenuItem { Header = "Copy path" };
                MenuItem miCopyFolder = new MenuItem { Header = "Folder path" };
                MenuItem miCopyFullPath = new MenuItem { Header = "Full path" };

                miCopyPathRoot.Items.Add(miCopyFolder);
                miCopyPathRoot.Items.Add(miCopyFullPath);

                MenuItem miFindTarget = new MenuItem { Header = "Open target folder..." };

                iconContextMenu.Items.Add(miCopyPathRoot);
                iconContextMenu.Items.Add(miFindTarget);

                // --- EVENT HANDLERS ---
                miEdit.Click += (s, e) => EditItem(item, fence, window);
                miMove.Click += (s, e) => ItemMoveDialog.ShowMoveDialog(item, fence, window.Dispatcher);

                miRemove.Click += (s, e) =>
                {
                    try
                    {
                        // Use the LIVE lookup to find the item to remove
                        var liveItem = GetLiveItem(); // Returns the JObject from the array

                        // We need the Array itself to remove the item
                        string fenceId = fence.Id?.ToString();
                        var liveFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);

                        if (liveFence != null && liveItem != null)
                        {
                            // Logic to locate the array containing liveItem
                            JArray targetArray = liveFence.Items as JArray; // Default to main

                            bool tabsEnabled = liveFence.TabsEnabled?.ToString().ToLower() == "true";
                            if (tabsEnabled)
                            {
                                var tabs = liveFence.Tabs as JArray;
                                int tabIdx = Convert.ToInt32(liveFence.CurrentTab?.ToString() ?? "0");
                                if (tabs != null && tabIdx < tabs.Count)
                                {
                                    targetArray = tabs[tabIdx]["Items"] as JArray;
                                }
                            }

                            if (targetArray != null)
                            {
                                targetArray.Remove(liveItem);
                                FenceDataManager.SaveFenceData();
                                var wp = VisualTreeHelper.GetParent(sp) as WrapPanel;
                                if (wp != null) wp.Children.Remove(sp);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error removing: {ex.Message}");
                    }
                };

                miCopyItem.Click += (s, e) => CopyPasteManager.CopyItem(item, fence);

                miFindTarget.Click += (s, e) => {
                    string target = Utility.GetShortcutTarget(filePath);
                    if (!string.IsNullOrEmpty(target) && (System.IO.File.Exists(target) || System.IO.Directory.Exists(target)))
                        Process.Start("explorer.exe", $"/select,\"{target}\"");
                };

                miCopyFolder.Click += (s, e) => {
                    string target = Utility.GetShortcutTarget(filePath);
                    if (!string.IsNullOrEmpty(target)) Clipboard.SetText(System.IO.Path.GetDirectoryName(target));
                };
                miCopyFullPath.Click += (s, e) => {
                    string target = Utility.GetShortcutTarget(filePath);
                    if (!string.IsNullOrEmpty(target)) Clipboard.SetText(target);
                };

                // --- DYNAMIC UPDATES (On Open) ---
                MenuItem miSendToDesktop = null;

                iconContextMenu.Opened += (s, e) => {
                    // 1. Send to Desktop (Ctrl)
                    bool isCtrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                    if (miSendToDesktop != null && iconContextMenu.Items.Contains(miSendToDesktop))
                        iconContextMenu.Items.Remove(miSendToDesktop);

                    if (isCtrl)
                    {
                        miSendToDesktop = new MenuItem { Header = "Send to Desktop" };
                        miSendToDesktop.Click += (sender, args) => CopyPasteManager.SendToDesktop(item);
                        int idx = iconContextMenu.Items.IndexOf(miCopyItem);
                        if (idx != -1) iconContextMenu.Items.Insert(idx + 1, miSendToDesktop);
                    }

                    // 2. REFRESH CHECK STATE from Live Data
                    if (miAlwaysAdmin != null)
                    {
                        var liveItem = GetLiveItem();
                        if (liveItem != null)
                        {
                            bool always = Convert.ToBoolean(liveItem["AlwaysRunAsAdmin"] ?? false);
                            miAlwaysAdmin.IsChecked = always;
                        }
                    }
                };

                sp.ContextMenu = iconContextMenu;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error attaching context menu: {ex.Message}");
            }
        }





        // Updates all heart ContextMenus across all fences using stored TextBlock references
        public static void UpdateAllHeartContextMenus()
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                foreach (var entry in _heartTextBlocks)
                {
                    var fence = entry.Key;
                    var heart = entry.Value;
                    if (heart != null)
                    {
                        // heart.ContextMenu = BuildHeartContextMenu(fence);
                        heart.ContextMenu = BuildHeartContextMenu(fence, false); // Normal menu by default
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Updated heart ContextMenu for fence '{fence.Title}'");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Skipped update for fence '{fence.Title}': heart TextBlock is null");
                    }
                }
            });
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


        /// <summary>
        /// Exports all icons from a Data fence to the desktop
        /// </summary>
        /// <param name="fence">The fence to export icons from</param>
        /// <param name="showConfirmation">If true, shows a message box on completion. False for silent automation.</param>
        private static void ExportAllIconsToDesktop(dynamic fence, bool showConfirmation = true)
        {
            try
            {
                // Verify this is a Data fence
                if (fence.ItemsType?.ToString() != "Data")
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                        $"Export all icons attempted on non-Data fence: {fence.ItemsType}");
                    return;
                }

                int exportedCount = 0;
                int totalCount = 0;

                // Check if fence has tabs
                bool hasTabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

                if (hasTabsEnabled && fence.Tabs != null)
                {
                    // Handle tabbed fence - export from all tabs
                    var tabs = fence.Tabs as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray();

                    foreach (var tab in tabs)
                    {
                        var tabObj = tab as Newtonsoft.Json.Linq.JObject;
                        var tabItems = tabObj?["Items"] as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray();

                        foreach (var item in tabItems)
                        {
                            totalCount++;
                            try
                            {
                                CopyPasteManager.SendToDesktop(item);
                                exportedCount++;
                            }
                            catch (Exception itemEx)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                                    $"Error exporting item to desktop: {itemEx.Message}");
                            }
                        }
                    }
                }
                else
                {
                    // Handle regular fence (no tabs)
                    var items = fence.Items as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray();

                    foreach (var item in items)
                    {
                        totalCount++;
                        try
                        {
                            CopyPasteManager.SendToDesktop(item);
                            exportedCount++;
                        }
                        catch (Exception itemEx)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                                $"Error exporting item to desktop: {itemEx.Message}");
                        }
                    }
                }

                // Show result message after a brief delay to allow desktop refresh
                string resultMessage = $"Exported {exportedCount} of {totalCount} icons to desktop.";
                if (exportedCount != totalCount)
                {
                    resultMessage += $" {totalCount - exportedCount} items failed to export.";
                }

                // Use DispatcherTimer for UI thread delay
                var delayTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(1700) // 1.7 second delay
                };
                delayTimer.Tick += (s, e) =>
                {
                    delayTimer.Stop();

                    // FIX: Check flag before showing message
                    if (showConfirmation)
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm(resultMessage, "Export Complete");
                    }
                };
                delayTimer.Start();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Export all icons completed for fence '{fence.Title}': {exportedCount}/{totalCount} successful");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in ExportAllIconsToDesktop: {ex.Message}");

                // Always show errors, even in silent mode, so user knows why it failed
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error exporting icons: {ex.Message}", "Export Error");
            }
        }

        /// <summary>
        /// Extracts URL from different browser drop data formats
        /// </summary>
        private static string ExtractUrlFromDropData(IDataObject dataObject)
        {
            try
            {
                // Try UniformResourceLocator format first (most browsers)
                if (dataObject.GetDataPresent("UniformResourceLocator"))
                {
                    object urlData = dataObject.GetData("UniformResourceLocator");
                    if (urlData is System.IO.MemoryStream stream)
                    {
                        byte[] bytes = new byte[stream.Length];
                        stream.Read(bytes, 0, (int)stream.Length);
                        string url = System.Text.Encoding.ASCII.GetString(bytes).Trim('\0');
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                            $"Extracted URL from UniformResourceLocator: {url}");
                        return url;
                    }
                }

                // Try Firefox format
                if (dataObject.GetDataPresent("text/x-moz-url"))
                {
                    string mozUrl = dataObject.GetData("text/x-moz-url") as string;
                    if (!string.IsNullOrEmpty(mozUrl))
                    {
                        string[] parts = mozUrl.Split('\n');
                        if (parts.Length > 0 && !string.IsNullOrEmpty(parts[0]))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                $"Extracted URL from text/x-moz-url: {parts[0]}");
                            return parts[0].Trim();
                        }
                    }
                }

                // Try HTML format  
                if (dataObject.GetDataPresent(DataFormats.Html))
                {
                    string html = dataObject.GetData(DataFormats.Html) as string;
                    if (!string.IsNullOrEmpty(html))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(html, @"href=['""]([^'""]+)['""]");
                        if (match.Success)
                        {
                            string url = match.Groups[1].Value;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                $"Extracted URL from HTML: {url}");
                            return url;
                        }
                    }
                }

                // Try plain text format (last resort)
                if (dataObject.GetDataPresent(DataFormats.Text))
                {
                    string text = dataObject.GetData(DataFormats.Text) as string;
                    if (!string.IsNullOrEmpty(text) && IsValidWebUrl(text.Trim()))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                            $"Extracted URL from text: {text.Trim()}");
                        return text.Trim();
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error extracting URL from drop data: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Validates if a string is a valid web URL
        /// </summary>
        private static bool IsValidWebUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return false;

            return Uri.TryCreate(url, UriKind.Absolute, out Uri validUri) &&
                   (validUri.Scheme == Uri.UriSchemeHttp || validUri.Scheme == Uri.UriSchemeHttps ||
                    validUri.Scheme.Equals("steam", StringComparison.OrdinalIgnoreCase));
        }






        /// <summary>
        /// Adds a URL shortcut to the fence from dropped URL
        /// </summary>
        private static void AddUrlShortcutToFence(string url, dynamic fence, WrapPanel wrapPanel)
        {
            try
            {
                // Create shortcuts directory if it doesn't exist
                if (!System.IO.Directory.Exists("Shortcuts"))
                {
                    System.IO.Directory.CreateDirectory("Shortcuts");
                }

                // Generate shortcut name from URL
                Uri uri = new Uri(url);
                string displayName = uri.Host.Replace("www.", "");
                string baseShortcutName = System.IO.Path.Combine("Shortcuts", $"{displayName}.url");
                string shortcutPath = baseShortcutName;

                // Ensure unique filename
                int counter = 1;
                while (System.IO.File.Exists(shortcutPath))
                {
                    shortcutPath = System.IO.Path.Combine("Shortcuts", $"{displayName} ({counter++}).url");
                }

                // Create the URL shortcut file  
                CreateWebLinkShortcut(url, shortcutPath, displayName);

                // Create item data for the fence
                dynamic newItem = new System.Dynamic.ExpandoObject();
                IDictionary<string, object> newItemDict = newItem;
                newItemDict["Filename"] = shortcutPath;
                newItemDict["IsFolder"] = false;
                newItemDict["IsLink"] = true;
                newItemDict["IsNetwork"] = false;
                newItemDict["DisplayName"] = displayName;
                newItemDict["AlwaysRunAsAdmin"] = false;

                // Find the fence in FenceDataManager.FenceData and update it properly
                string fenceId = fence.Id?.ToString();
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);

                if (fenceIndex >= 0)
                {
                    // Get the actual fence from FenceDataManager.FenceData
                    dynamic actualFence = FenceDataManager.FenceData[fenceIndex];
                    var actualItems = actualFence.Items as JArray ?? new JArray();

                    // Handle tabs if enabled
                    bool tabsEnabled = actualFence.TabsEnabled?.ToString().ToLower() == "true";
                    JArray targetList = actualItems;

                    if (tabsEnabled)
                    {
                        var tabs = actualFence.Tabs as JArray ?? new JArray();
                        int currentTab = Convert.ToInt32(actualFence.CurrentTab?.ToString() ?? "0");

                        if (currentTab >= 0 && currentTab < tabs.Count)
                        {
                            var activeTab = tabs[currentTab] as JObject;
                            if (activeTab != null)
                            {
                                var tabItems = activeTab["Items"] as JArray ?? new JArray();
                                tabItems.Add(JObject.FromObject(newItem));
                                targetList = tabItems; // Mark as target for ordering
                            }
                        }
                    }
                    else
                    {
                        actualItems.Add(JObject.FromObject(newItem));
                    }

                    // Set DisplayOrder
                    newItemDict["DisplayOrder"] = targetList.Count - 1;

                    // Save fence data BEFORE creating the icon
                    FenceDataManager.SaveFenceData();

                    // Add icon to UI using FenceManager.AddIcon for proper sizing
                    FenceManager.AddIcon(newItem, wrapPanel);
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                        $"Could not find fence with ID {fenceId} in FenceDataManager.FenceData");
                }

                // Add event handlers to the new icon
                StackPanel sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as StackPanel;
                if (sp != null)
                {
                    ClickEventAdder(sp, shortcutPath, false);

                    // FIX: Use the FULL Context Menu (Edit, Move, Copy, etc.)
                    // We need to find the parent window to pass it to the menu builder
                    NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(wrapPanel);
                    if (parentWindow != null)
                    {
                        AttachIconContextMenu(sp, newItem, fence, parentWindow);
                    }
                    else
                    {
                        // Fallback (rare) - Should not happen if UI is loaded
                        CreateBasicContextMenu(sp, newItem, shortcutPath);
                    }
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    $"Successfully added URL shortcut: {displayName} -> {url}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error adding URL shortcut to fence: {ex.Message}");
            }
        }



       /// <summary>
        /// Gets the next display order for new items
        /// </summary>
        private static int GetNextDisplayOrder(dynamic fence)
        {
            try
            {
                var items = fence.Items as JArray ?? new JArray();
                int maxOrder = items.Count > 0 ? items.Max(i => i["DisplayOrder"]?.Value<int>() ?? 0) : -1;
                return maxOrder + 1;
            }
            catch { return 0; }
        }

        /// <summary>
        /// Creates basic context menu for URL shortcuts
        /// </summary>
        private static void CreateBasicContextMenu(StackPanel sp, dynamic item, string filePath)
        {
            try
            {
                ContextMenu iconContextMenu = new ContextMenu();
                MenuItem miRemove = new MenuItem { Header = "Remove" };

        
                miRemove.Click += (sender, e) =>
                {
                    try
                    {
                        NonActivatingWindow parentWin = FindVisualParent<NonActivatingWindow>(sp);
                        if (parentWin != null)
                        {
                            string fenceId = parentWin.Tag?.ToString();
                            var fence = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == fenceId);
                            if (fence != null)
                            {
                                var items = fence.Items as JArray;
                                var itemToRemove = items?.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                                if (itemToRemove != null)
                                {
                                    items.Remove(itemToRemove);
                                    FenceDataManager.SaveFenceData();

                                    WrapPanel wrapPanel = FindVisualParent<WrapPanel>(sp);
                                    wrapPanel?.Children.Remove(sp);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                            $"Error removing URL shortcut: {ex.Message}");
                    }
                };




                iconContextMenu.Items.Add(miRemove);

                bool alwaysAdmin = Convert.ToBoolean(item["AlwaysRunAsAdmin"] ?? false);
                MenuItem miAlwaysAdmin = new MenuItem
                {
                    Header = "Always run as administrator",
                    IsCheckable = true,
                    IsChecked = alwaysAdmin
                };
                miAlwaysAdmin.Click += (sender, e) => {
                    try
                    {
                        NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(sp);
                        if (parentWindow != null)
                        {
                            string fenceId = parentWindow.Tag?.ToString();
                            if (!string.IsNullOrEmpty(fenceId))
                            {
                                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                                if (currentFence != null && currentFence.ItemsType?.ToString() == "Data")
                                {
                                    var items = currentFence.Items as JArray ?? new JArray();
                                    bool tabsEnabled = currentFence.TabsEnabled?.ToString().ToLower() == "true";
                                    if (tabsEnabled)
                                    {
                                        var tabs = currentFence.Tabs as JArray ?? new JArray();
                                        int currentTabIndex = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");
                                        if (currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                                        {
                                            var currentTab = tabs[currentTabIndex] as JObject;
                                            items = currentTab?["Items"] as JArray ?? items;
                                        }
                                    }
                                    var matchingItem = items.FirstOrDefault(i => string.Equals(
                                        System.IO.Path.GetFullPath(i["Filename"]?.ToString() ?? ""),
                                        System.IO.Path.GetFullPath(filePath),
                                        StringComparison.OrdinalIgnoreCase));
                                    if (matchingItem != null)
                                    {
                                        bool currentValue = Convert.ToBoolean(matchingItem["AlwaysRunAsAdmin"] ?? false);
                                        bool newValue = !currentValue;
                                        matchingItem["AlwaysRunAsAdmin"] = newValue;
                                        miAlwaysAdmin.IsChecked = newValue;
                                        FenceDataManager.SaveFenceData();
                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Toggled AlwaysRunAsAdmin to {newValue} for item: {filePath}");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error toggling AlwaysRunAsAdmin for {filePath}: {ex.Message}");
                    }
                };
                // Refresh IsChecked on menu open
                iconContextMenu.Opened += (s, ev) => {
                    try
                    {
                        NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(sp);
                        if (parentWindow != null)
                        {
                            string fenceId = parentWindow.Tag?.ToString();
                            if (!string.IsNullOrEmpty(fenceId))
                            {
                                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                                if (currentFence != null)
                                {
                                    var items = currentFence.Items as JArray ?? new JArray();
                                    bool tabsEnabled = currentFence.TabsEnabled?.ToString().ToLower() == "true";
                                    if (tabsEnabled)
                                    {
                                        var tabs = currentFence.Tabs as JArray ?? new JArray();
                                        int currentTabIndex = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");
                                        if (currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                                        {
                                            var currentTab = tabs[currentTabIndex] as JObject;
                                            items = currentTab?["Items"] as JArray ?? items;
                                        }
                                    }
                                    var matchingItem = items.FirstOrDefault(i => string.Equals(
                                        System.IO.Path.GetFullPath(i["Filename"]?.ToString() ?? ""),
                                        System.IO.Path.GetFullPath(filePath),
                                        StringComparison.OrdinalIgnoreCase));
                                    if (matchingItem != null)
                                    {
                                        miAlwaysAdmin.IsChecked = Convert.ToBoolean(matchingItem["AlwaysRunAsAdmin"] ?? false);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Error refreshing AlwaysRunAsAdmin IsChecked for {filePath}: {ex.Message}");
                    }
                };
                iconContextMenu.Items.Add(miAlwaysAdmin);











                sp.ContextMenu = iconContextMenu;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error creating context menu: {ex.Message}");
            }
        }
             
        public static List<dynamic> GetFenceData()
        {
            return FenceDataManager.FenceData;
        }

                /// <summary>
        /// Public access to portal fences dictionary for CustomizeFenceForm
        /// </summary>
        public static Dictionary<dynamic, PortalFenceManager> GetPortalFences()
        {
            return _portalFences;
        }




 
        /// <summary>
        /// Updates the filter history for a fence using LRU (Least Recently Used) logic.
        /// Max 5 items. Duplicates move to top.
        /// </summary>
        private static void UpdateFilterHistory(dynamic fence, string newFilter)
        {
            if (string.IsNullOrWhiteSpace(newFilter)) return;

            // Don't save if it matches a Standard Preset value exactly
            if (_standardPresets.ContainsValue(newFilter)) return;

            try
            {
                // 1. GET FRESH DATA (Fixes the "Replace" bug)
                // We cannot rely on the 'fence' parameter because it might be stale 
                // (referencing the state from when the window opened).
                string fenceId = fence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                int index = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (index < 0) return;

                // Grab the LIVE object from the master list
                dynamic liveFence = FenceDataManager.FenceData[index];

                // 2. Convert to Dictionary for modification
                IDictionary<string, object> fenceDict = liveFence is IDictionary<string, object> dict
                    ? dict
                    : ((JObject)liveFence).ToObject<IDictionary<string, object>>();

                // 3. Get existing history
                JArray history;
                if (fenceDict.ContainsKey("FilterHistory") && fenceDict["FilterHistory"] is JArray existingArr)
                {
                    history = existingArr;
                }
                else
                {
                    history = new JArray();
                }

                // 4. Update Logic (LRU)
                // Remove existing if present (to bump to top)
                for (int i = history.Count - 1; i >= 0; i--)
                {
                    if (string.Equals(history[i].ToString(), newFilter, StringComparison.OrdinalIgnoreCase))
                    {
                        history.RemoveAt(i);
                    }
                }

                // Insert at top
                history.Insert(0, newFilter);

                // Cap at 5 items
                while (history.Count > 5)
                {
                    history.RemoveAt(history.Count - 1);
                }

                // 5. Save back to Master List
                fenceDict["FilterHistory"] = history;
                FenceDataManager.FenceData[index] = JObject.FromObject(fenceDict);
                FenceDataManager.SaveFenceData();

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings,
                    $"Updated Filter History for {fenceId}: {string.Join(", ", history)}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings, $"History Error: {ex.Message}");
            }
        }


        /// <summary>
        /// Resets visual customizations for all fences to default values.
        /// Preserves position, size, content, and system state (Locked/Rolled).
        /// </summary>
        public static void ResetAllCustomizations()
        {
            try
            {
                bool modificationsMade = false;

                // FIX: Use .ToList() to create a snapshot copy for iteration.
                // This prevents "Collection was modified" errors when we update the real list inside the loop.
                var fencesSnapshot = FenceDataManager.FenceData.ToList();

                foreach (dynamic fence in fencesSnapshot)
                {
                    // Convert to JObject/Dictionary for safe property access
                    IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict
                        ? dict
                        : ((JObject)fence).ToObject<IDictionary<string, object>>();

                    // --- RESET VISUALS TO DEFAULTS ---

                    // Fence Appearance
                    fenceDict["CustomColor"] = null;
                    fenceDict["CustomLaunchEffect"] = null;
                    fenceDict["FenceBorderColor"] = null;
                    fenceDict["FenceBorderThickness"] = 2; // Default thickness

                    // Title Appearance
                    fenceDict["TitleTextColor"] = null;
                    fenceDict["TitleTextSize"] = "Medium";
                    fenceDict["BoldTitleText"] = "false";

                    // Icon Appearance
                    fenceDict["TextColor"] = null;
                    fenceDict["DisableTextShadow"] = "false";
                    fenceDict["IconSize"] = "Medium";
                    fenceDict["IconSpacing"] = 5;
                    fenceDict["GrayscaleIcons"] = "false";

                    // Update the object in the REAL list
                    int index = FenceDataManager.FenceData.IndexOf(fence);
                    if (index >= 0)
                    {
                        FenceDataManager.FenceData[index] = JObject.FromObject(fenceDict);
                        modificationsMade = true;
                    }
                }

                if (modificationsMade)
                {
                    FenceDataManager.SaveFenceData();

                    // Reload to reflect changes immediately
                    ReloadFences();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.Settings,
                        "All fence customizations have been reset to defaults.");

                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(
                        "All visual customizations have been reset.\nYour icons and fence positions were preserved.",
                        "Reset Complete");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Error,
                    $"Error resetting customizations: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error resetting customizations: {ex.Message}", "Error");
            }
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
                int index = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (index >= 0)
                {
                    // Get the fence from FenceDataManager.FenceData
                    dynamic actualFence = FenceDataManager.FenceData[index];





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

                    // Update the fence in FenceDataManager.FenceData
                    FenceDataManager.FenceData[index] = JObject.FromObject(fenceDict);
                    FenceDataManager.SaveFenceData();
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
                            double targetHeight = isRolled ? 28 : Convert.ToDouble(actualFence.UnrolledHeight?.ToString() ?? "130"); //rolled height
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
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate, $"Failed to find fence '{fence.Title}' in FenceDataManager.FenceData for {propertyName} update");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate, $"Error updating {propertyName} for fence '{fence.Title}': {ex.Message}");
            }
        }



        // TABS FEATURE: Toggle tabs with strict Data Migration and Effect Cleanup
        private static void ToggleFenceTabs(dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Toggling tabs for fence '{fence.Title}'");

                string fenceId = fence.Id?.ToString();
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (fenceIndex < 0) return;

                dynamic liveFence = FenceDataManager.FenceData[fenceIndex];
                IDictionary<string, object> fenceDict = liveFence is IDictionary<string, object> dict
                    ? dict : ((JObject)liveFence).ToObject<IDictionary<string, object>>();

                bool currentTabsEnabled = fenceDict.ContainsKey("TabsEnabled") && fenceDict["TabsEnabled"]?.ToString().ToLower() == "true";
                bool newTabsEnabled = !currentTabsEnabled;

                var mainItems = fenceDict["Items"] as JArray ?? new JArray();
                var tabs = fenceDict["Tabs"] as JArray ?? new JArray();

                // 1. Data Migration Logic
                if (newTabsEnabled)
                {
                    // ENABLING TABS: Move Main -> Tab
                    if (tabs.Count == 0)
                    {
                        // Create first tab with Herb Naming
                        dynamic newTab = new JObject();
                        string herbName = FenceUtilities.GenerateRandomHerbName();
                        newTab.TabName = $"0. {herbName}";
                        newTab.Items = mainItems;
                        tabs.Add(newTab);
                        fenceDict["CurrentTab"] = 0;
                    }
                    else
                    {
                        int targetIdx = Convert.ToInt32(fenceDict["CurrentTab"]?.ToString() ?? "0");
                        if (targetIdx < 0 || targetIdx >= tabs.Count) targetIdx = 0;

                        var targetTab = tabs[targetIdx] as JObject;
                        var targetItems = targetTab["Items"] as JArray ?? new JArray();

                        foreach (var item in mainItems) targetItems.Add(item);
                        targetTab["Items"] = targetItems;
                        fenceDict["CurrentTab"] = targetIdx;
                    }
                    // Clear main items after moving them to tab
                    fenceDict["Items"] = new JArray();
                    fenceDict["Tabs"] = tabs;
                }
                else
                {
                    // DISABLING TABS: Move Active Tab -> Main
                    int currentTabIdx = Convert.ToInt32(fenceDict["CurrentTab"]?.ToString() ?? "0");
                    if (tabs.Count > 0 && currentTabIdx < tabs.Count && currentTabIdx >= 0)
                    {
                        var activeTab = tabs[currentTabIdx] as JObject;
                        var activeItems = activeTab["Items"] as JArray;
                        if (activeItems != null)
                        {
                            // FIX: Clear mainItems first! 
                            // This prevents duplication because mainItems might already contain stale copies 
                            // (especially from Tab 0 syncing).
                            mainItems.Clear();

                            foreach (var item in activeItems) mainItems.Add(item);
                            activeTab["Items"] = new JArray();
                        }
                    }
                    fenceDict["Items"] = mainItems;
                }

                // 2. Save Updates
                fenceDict["TabsEnabled"] = newTabsEnabled.ToString().ToLower();

                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                var currentWindow = windows.FirstOrDefault(w => w.Tag?.ToString() == fenceId);
                if (currentWindow != null)
                {
                    fenceDict["X"] = currentWindow.Left;
                    fenceDict["Y"] = currentWindow.Top;
                    fenceDict["Width"] = currentWindow.Width;
                    fenceDict["Height"] = currentWindow.Height;
                }

                FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                FenceDataManager.SaveFenceData();

                // 3. Reload Window & CLEANUP EFFECTS
                if (currentWindow != null)
                {
                    // Cleanup Legendary Effects
                    string currentTitle = fence.Title?.ToString() ?? "";
                    if (currentTitle == "Nikos" || currentTitle == "Nikos Georgousis" || currentTitle.Contains(">:"))
                    {
                        InterCore.ProcessTitleChange(liveFence, "RESET_EFFECT_TEMP", "");
                    }

                    _heartTextBlocks.Remove(fence);
                    if (_portalFences.ContainsKey(fence))
                    {
                        _portalFences[fence].Dispose();
                        _portalFences.Remove(fence);
                    }
                    currentWindow.Close();
                }

                var updatedFence = FenceDataManager.FenceData[fenceIndex];

                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    CreateFence(updatedFence, new TargetChecker(1000));
                }));
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error toggling tabs: {ex.Message}");
            }
        }




        // TABS FEATURE: Dynamically refresh the tab strip UI (Ghost Arrows)
        // v2.5. 4.181: Adds smart left/right chevrons that appear only when tabs are hidden.
        public static void RefreshTabStripUI(NonActivatingWindow fenceWindow, dynamic fence)
        {
            try
            {
                var border = fenceWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                // 1. CAPTURE SCROLL STATE
                double previousScrollOffset = 0;
                var oldContainer = dockPanel.Children.OfType<Grid>()
                    .FirstOrDefault(g => g.Tag?.ToString() == "TAB_STRIP_CONTAINER");
                if (oldContainer != null)
                {
                    var oldScroll = oldContainer.Children.OfType<ScrollViewer>().FirstOrDefault();
                    if (oldScroll != null) previousScrollOffset = oldScroll.HorizontalOffset;
                }

                // 2. CLEANUP
                var existingStrips = dockPanel.Children.OfType<FrameworkElement>()
                    .Where(c => c is Grid g && g.Tag?.ToString() == "TAB_STRIP_CONTAINER" ||
                                c is StackPanel sp && sp.Height == 20)
                    .ToList();
                foreach (var oldStrip in existingStrips) dockPanel.Children.Remove(oldStrip);

                // 3. CHECK ENABLED
                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                if (!tabsEnabled) return;

                // 4. COLOR ANALYSIS (For Arrow Visibility)
                string fenceColorName = fence.CustomColor?.ToString() ?? SettingsManager.SelectedColor;
                string effectiveColor = !string.IsNullOrEmpty(fenceColorName) ? fenceColorName : SettingsManager.SelectedColor;

                System.Windows.Media.Color baseColor = System.Windows.Media.Colors.Gray;
                try
                {
                    var drawingColor = Utility.GetColorFromName(effectiveColor);
                    baseColor = System.Windows.Media.Color.FromArgb(255, drawingColor.R, drawingColor.G, drawingColor.B);
                }
                catch { }

                string c = effectiveColor?.ToLower() ?? "";
                bool isExplicitDark = c.Contains("blue") || c.Contains("teal") || c.Contains("black") ||
                                      c.Contains("red") || c.Contains("green") || c.Contains("purple") ||
                                      c.Contains("bismark") || c.Contains("fuchsia") || c.Contains("default");

                double brightness = Math.Sqrt(
                    (0.299 * baseColor.R * baseColor.R) +
                    (0.587 * baseColor.G * baseColor.G) +
                    (0.114 * baseColor.B * baseColor.B)
                );

                bool isDarkTheme = isExplicitDark || brightness < 160;
                SolidColorBrush arrowBrush = isDarkTheme ? System.Windows.Media.Brushes.White : new SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50));

                // 5. CREATE NEW GRID STRUCTURE
                // Cols: [LeftArrow] [ScrollViewer*] [RightArrow] [PlusButton]
                Grid containerGrid = new Grid
                {
                    Tag = "TAB_STRIP_CONTAINER",
                    Height = 20,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(20, 0, 0, 0)),
                    Margin = new Thickness(0, 1, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top
                };

                containerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // 0: Left Arrow
                containerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // 1: Tabs
                containerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // 2: Right Arrow
                containerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // 3: Plus Button

                // 6. CREATE SCROLLVIEWER
                ScrollViewer scrollViewer = new ScrollViewer
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    PanningMode = PanningMode.HorizontalOnly,
                    CanContentScroll = true
                };

                scrollViewer.PreviewMouseWheel += (s, e) =>
                {
                    if (e.Delta > 0) scrollViewer.LineLeft();
                    else scrollViewer.LineRight();
                    e.Handled = true;
                };

                StackPanel tabStack = new StackPanel { Orientation = Orientation.Horizontal };
                scrollViewer.Content = tabStack;

                // 7. CREATE GHOST ARROWS
                TextBlock leftArrow = new TextBlock
                {
                    Text = "‹", // Elegant chevron
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Foreground = arrowBrush,
                    Opacity = 0.6,
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2, -2, 2, 0), // Slight nudge up
                    Visibility = Visibility.Collapsed // Hidden by default
                };
                leftArrow.MouseLeftButtonDown += (s, e) => { scrollViewer.LineLeft(); e.Handled = true; };

                TextBlock rightArrow = new TextBlock
                {
                    Text = "›",
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Foreground = arrowBrush,
                    Opacity = 0.6,
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2, -2, 2, 0),
                    Visibility = Visibility.Collapsed
                };
                rightArrow.MouseLeftButtonDown += (s, e) => { scrollViewer.LineRight(); e.Handled = true; };

                // 8. WIRE UP SCROLL LOGIC
                // This updates visibility whenever the scroll position or size changes
                scrollViewer.ScrollChanged += (s, e) =>
                {
                    // Only show left if we have scrolled right
                    leftArrow.Visibility = scrollViewer.HorizontalOffset > 0
                        ? Visibility.Visible : Visibility.Collapsed;

                    // Only show right if there is scrollable content remaining
                    rightArrow.Visibility = scrollViewer.HorizontalOffset < scrollViewer.ScrollableWidth
                        ? Visibility.Visible : Visibility.Collapsed;
                };

                // Add Elements to Grid
                containerGrid.Children.Add(leftArrow); Grid.SetColumn(leftArrow, 0);
                containerGrid.Children.Add(scrollViewer); Grid.SetColumn(scrollViewer, 1);
                containerGrid.Children.Add(rightArrow); Grid.SetColumn(rightArrow, 2);

                // 9. POPULATE TABS
                var tabs = fence.Tabs as JArray ?? new JArray();
                int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");

                for (int i = 0; i < tabs.Count; i++)
                {
                    var tab = tabs[i] as JObject;
                    if (tab == null) continue;

                    string tabName = tab["TabName"]?.ToString() ?? $"Tab {i + 1}";
                    bool isActiveTab = (i == currentTab);
                    int capturedIndex = i;

                    Button tabButton = new Button
                    {
                        Content = tabName,
                        Tag = i,
                        Height = 18,
                        MinWidth = 50,
                        Margin = new Thickness(1, 0, 1, 0),
                        Padding = new Thickness(10, 2, 10, 2),
                        FontSize = 10,
                        FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                        Cursor = Cursors.Hand,
                        Focusable = false
                    };

                    ApplyTabStyle(tabButton, isActiveTab, fenceColorName);

                    tabButton.PreviewMouseLeftButtonDown += (s, e) =>
                    {
                        fenceWindow.Focus();
                        System.Windows.Input.Keyboard.ClearFocus();

                        if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                        {
                            e.Handled = true;
                            RenameTab(fence, capturedIndex, fenceWindow);
                        }
                        else
                        {
                            SwitchTabByFence(fence, capturedIndex, fenceWindow);
                            e.Handled = true;
                        }
                    };

                    tabButton.PreviewMouseRightButtonDown += (s, e) => SwitchTabByFence(fence, capturedIndex, fenceWindow);

                    // Context Menu
                    ContextMenu tabContextMenu = new ContextMenu();
                    MenuItem miAddTab = new MenuItem { Header = "Add New Tab" };
                    MenuItem miRenameTab = new MenuItem { Header = "Rename Tab" };
                    MenuItem miDeleteTab = new MenuItem { Header = "Delete Tab" };
                    MenuItem miMoveLeft = new MenuItem { Header = "Move Left" };
                    MenuItem miMoveRight = new MenuItem { Header = "Move Right" };

                    tabContextMenu.Items.Add(miAddTab);
                    tabContextMenu.Items.Add(new Separator());
                    tabContextMenu.Items.Add(miRenameTab);
                    tabContextMenu.Items.Add(miDeleteTab);
                    tabContextMenu.Items.Add(new Separator());
                    tabContextMenu.Items.Add(miMoveLeft);
                    tabContextMenu.Items.Add(miMoveRight);

                    miAddTab.Click += (s, e) => AddNewTab(fence, fenceWindow);
                    miRenameTab.Click += (s, e) => RenameTab(fence, capturedIndex, fenceWindow);
                    miDeleteTab.Click += (s, e) => DeleteTab(fence, capturedIndex, fenceWindow);
                    miMoveLeft.Click += (s, e) => MoveTab(fence, capturedIndex, -1, fenceWindow);
                    miMoveRight.Click += (s, e) => MoveTab(fence, capturedIndex, 1, fenceWindow);

                    tabContextMenu.Opened += (s, e) =>
                    {
                        miMoveLeft.IsEnabled = capturedIndex > 0;
                        miMoveRight.IsEnabled = capturedIndex < tabs.Count - 1;
                    };

                    tabButton.ContextMenu = tabContextMenu;
                    tabStack.Children.Add(tabButton);
                }

                // 10. POPULATE PINNED [+] BUTTON
                Button addTabButton = new Button
                {
                    Content = "+",
                    Tag = "ADD_TAB",
                    Height = 18,
                    Width = 25,
                    Margin = new Thickness(3, 0, 1, 0),
                    FontSize = 12,
                    FontWeight = FontWeights.Bold,
                    FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                    BorderThickness = new Thickness(1),
                    Cursor = Cursors.Hand,
                    ToolTip = "Add new tab (Ctrl+Click to Import)", // Updated Tooltip
                    Focusable = false
                };

                ApplyTabStyle(addTabButton, false, fenceColorName, true);

                bool isAddingTab = false;
                addTabButton.PreviewMouseLeftButtonDown += async (s, e) =>
                {
                    fenceWindow.Focus();
                    e.Handled = true;

                    // --- NEW IMPORT LOGIC ---
                    if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                    {
                        // Advanced: Import Tab
                        ImportTabManager.HandleImportRequest(fence, fenceWindow);
                        return;
                    }
                    // ------------------------

                    if (isAddingTab) return;
                    isAddingTab = true;
                    try { AddNewTab(fence, fenceWindow); }
                    finally { await System.Threading.Tasks.Task.Delay(500); isAddingTab = false; }
                };

                containerGrid.Children.Add(addTabButton);
                Grid.SetColumn(addTabButton, 3); // Col 3 is for the Button

                // 11. SURGICAL INSERTION
                DockPanel.SetDock(containerGrid, Dock.Top);

                int insertIndex = 0;
                bool titleFound = false;
                for (int i = 0; i < dockPanel.Children.Count; i++)
                {
                    if (dockPanel.Children[i] is Grid g)
                    {
                        if (g.Children.OfType<TextBlock>().Any(tb => tb.Name == "FenceLockIcon"))
                        {
                            insertIndex = i + 1;
                            titleFound = true;
                            break;
                        }
                    }
                }
                if (!titleFound) insertIndex = 0;

                if (insertIndex < dockPanel.Children.Count &&
                    dockPanel.Children[insertIndex] is Grid potentialFilter &&
                    potentialFilter.Children.OfType<ComboBox>().Any())
                {
                    insertIndex++;
                }

                if (insertIndex > dockPanel.Children.Count) insertIndex = dockPanel.Children.Count;
                dockPanel.Children.Insert(insertIndex, containerGrid);

                // 12. RESTORE SCROLL (Async to allow layout pass)
                if (previousScrollOffset > 0)
                {
                    fenceWindow.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
                    {
                        scrollViewer.ScrollToHorizontalOffset(previousScrollOffset);
                    }));
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Refreshed Tab Strip with Ghost Arrows");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing tab strip UI: {ex.Message}");
            }
        }      // TABS FEATURE: Add new tab with random herb name
        public static void AddNewTab(dynamic fence, NonActivatingWindow fenceWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"AddNewTab called for fence '{fence.Title}'");

                // Get fresh fence data
                string fenceId = fence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot add tab: fence ID missing");
                    return;
                }

                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Cannot add tab: fence with ID '{fenceId}' not found");
                    return;
                }

                var tabs = currentFence.Tabs as JArray ?? new JArray();

                // Generate random herb name with tab index
                string herbName = FenceUtilities.GenerateRandomHerbName();
                string newTabName = $"{tabs.Count}. {herbName}";

                // Create new tab object
                var newTab = new JObject();
                newTab["TabName"] = newTabName;
                newTab["Items"] = new JArray();

                // Add to tabs array
                tabs.Add(newTab);

                // Update fence data properly
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (fenceIndex >= 0)
                {
                    IDictionary<string, object> fenceDict = currentFence is IDictionary<string, object> dict ?
                        dict : ((JObject)currentFence).ToObject<IDictionary<string, object>>();

                    fenceDict["Tabs"] = tabs; // Store JArray directly
                    fenceDict["CurrentTab"] = tabs.Count - 1; // Switch to new tab

                    FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                    FenceDataManager.SaveFenceData();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Added new tab '{newTabName}' to fence '{currentFence.Title}'");

                    // Get updated fence and refresh the display
                    var updatedFence = FenceDataManager.FenceData[fenceIndex];
                    int newTabIndex = tabs.Count - 1;

                    // Refresh content and styling
                    RefreshFenceContentSimple(fenceWindow, updatedFence, newTabIndex);
                    RefreshTabStyling(fenceWindow, newTabIndex);

                    // Refresh the entire tab strip to show new tab
                    RefreshTabStripUI(fenceWindow, updatedFence);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "New tab added successfully");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error adding new tab: {ex.Message}");
            }
        }


        // TABS FEATURE: Rename tab with inline editing (in-button, focus-enabled)
        // v2.5.4.187: Fixed Visual Tree traversal to support ScrollViewer/Ghost Arrows structure
        public static void RenameTab(dynamic fence, int tabIndex, NonActivatingWindow fenceWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"RenameTab called for fence '{fence.Title}', tab {tabIndex}");

                // 1. Find the main DockPanel
                var border = fenceWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                // 2. FIX: Find the new Grid Container first (Tag: TAB_STRIP_CONTAINER)
                var containerGrid = dockPanel.Children.OfType<Grid>()
                    .FirstOrDefault(g => g.Tag?.ToString() == "TAB_STRIP_CONTAINER");

                if (containerGrid == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "RenameTab: Tab strip container not found.");
                    return;
                }

                // 3. FIX: Find the ScrollViewer inside the Grid
                var scrollViewer = containerGrid.Children.OfType<ScrollViewer>().FirstOrDefault();
                if (scrollViewer == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "RenameTab: ScrollViewer not found.");
                    return;
                }

                // 4. FIX: Get the StackPanel from the ScrollViewer content
                var tabStrip = scrollViewer.Content as StackPanel;
                if (tabStrip == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "RenameTab: Tab StackPanel not found.");
                    return;
                }

                // 5. Find the button for this tab index
                Button targetButton = null;
                foreach (Button btn in tabStrip.Children.OfType<Button>())
                {
                    if (btn.Tag is int buttonTabIndex && buttonTabIndex == tabIndex)
                    {
                        targetButton = btn;
                        break;
                    }
                }

                if (targetButton == null) return;

                // Get current tab data
                string fenceId = fence.Id?.ToString();
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                var tabs = currentFence.Tabs as JArray ?? new JArray();
                if (tabIndex < 0 || tabIndex >= tabs.Count) return;

                var tab = tabs[tabIndex] as JObject;
                if (tab == null) return;

                string currentName = tab["TabName"]?.ToString() ?? $"Tab {tabIndex}";

                // CRITICAL: Disable NonActivatingWindow focus prevention during editing
                fenceWindow.EnableFocusPrevention(false);

                // Temporarily increase button height to accommodate TextBox properly
                double originalHeight = targetButton.Height;
                targetButton.Height = 22; // Slightly taller for TextBox

                // Create TextBox for inline editing
                TextBox editTextBox = new TextBox
                {
                    Text = currentName,
                    FontSize = targetButton.FontSize,
                    FontFamily = targetButton.FontFamily,
                    FontWeight = FontWeights.Normal,
                    Background = System.Windows.Media.Brushes.White,
                    Foreground = System.Windows.Media.Brushes.Black,
                    BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 70, 130, 180)),
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(4, 2, 4, 2),
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Stretch,
                    VerticalContentAlignment = VerticalAlignment.Center
                };

                // Store original button properties
                object originalContent = targetButton.Content;
                var originalBackground = targetButton.Background;
                var originalForeground = targetButton.Foreground;
                var originalBorderBrush = targetButton.BorderBrush;

                targetButton.Content = editTextBox;
                targetButton.Background = System.Windows.Media.Brushes.White;
                targetButton.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 70, 130, 180));

                // Focus and select text with proper timing
                System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Background,
                    new Action(() =>
                    {
                        editTextBox.Focus();
                        editTextBox.SelectAll();

                        // Additional focus call to ensure it takes
                        System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                            System.Windows.Threading.DispatcherPriority.Background,
                            new Action(() =>
                            {
                                if (!editTextBox.IsFocused)
                                {
                                    editTextBox.Focus();
                                }
                            })
                        );
                    })
                );

                bool editingComplete = false;

                // Action to complete editing and restore NonActivatingWindow behavior
                Action<bool> completeEditing = (save) =>
                {
                    if (editingComplete) return;
                    editingComplete = true;

                    try
                    {
                        if (save && !string.IsNullOrWhiteSpace(editTextBox.Text))
                        {
                            string newName = editTextBox.Text.Trim();

                            // Validate name length
                            if (newName.Length > 30)
                            {
                                newName = newName.Substring(0, 30);
                            }

                            // Update tab name in data
                            tab["TabName"] = newName;

                            // Save to JSON
                            int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                            if (fenceIndex >= 0)
                            {
                                IDictionary<string, object> fenceDict = currentFence is IDictionary<string, object> dict ?
                                    dict : ((JObject)currentFence).ToObject<IDictionary<string, object>>();

                                fenceDict["Tabs"] = tabs;
                                FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                                FenceDataManager.SaveFenceData();

                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                                    $"Renamed tab from '{currentName}' to '{newName}'");

                                // Update button with new name
                                targetButton.Content = newName;
                            }
                        }
                        else
                        {
                            // Cancel - restore original content
                            targetButton.Content = originalContent;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Tab rename cancelled");
                        }

                        // Restore original button properties
                        targetButton.Background = originalBackground;
                        targetButton.Foreground = originalForeground;
                        targetButton.BorderBrush = originalBorderBrush;
                        targetButton.Height = originalHeight;

                        // CRITICAL: Re-enable NonActivatingWindow focus prevention
                        fenceWindow.EnableFocusPrevention(true);

                        // Return focus to fence window
                        fenceWindow.Focus();
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                            $"Error completing tab rename: {ex.Message}");

                        // Restore everything on error
                        targetButton.Content = originalContent;
                        targetButton.Background = originalBackground;
                        targetButton.Foreground = originalForeground;
                        targetButton.BorderBrush = originalBorderBrush;
                        targetButton.Height = originalHeight;
                        fenceWindow.EnableFocusPrevention(true);
                        fenceWindow.Focus();
                    }
                };

                // Handle Enter key (save) and Escape (cancel)
                editTextBox.KeyDown += (s, e) =>
                {
                    if (e.Key == Key.Enter)
                    {
                        completeEditing(true);
                        e.Handled = true;
                    }
                    else if (e.Key == Key.Escape)
                    {
                        completeEditing(false);
                        e.Handled = true;
                    }
                };

                // Handle focus loss (save)
                editTextBox.LostFocus += (s, e) =>
                {
                    System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        new Action(() => completeEditing(true))
                    );
                };

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Started inline editing for tab '{currentName}' with focus enabled");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error starting tab rename: {ex.Message}");

                // Ensure focus prevention is restored on any error
                fenceWindow.EnableFocusPrevention(true);
            }
        }

        // TABS FEATURE: Delete tab with confirmation if it has items
        // TABS FEATURE: Delete tab with confirmation, Auto-Export, and Sync Fix
        public static void DeleteTab(dynamic fence, int tabIndex, NonActivatingWindow fenceWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"DeleteTab called for fence '{fence.Title}', tab {tabIndex}");

                // 1. Get Fresh Data
                string fenceId = fence.Id?.ToString();
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                var tabs = currentFence.Tabs as JArray ?? new JArray();

                // 2. Validation: Don't allow deleting the last tab
                if (tabs.Count <= 1)
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Cannot delete the last remaining tab.", "Delete Tab");
                    return;
                }

                if (tabIndex < 0 || tabIndex >= tabs.Count) return;

                var tab = tabs[tabIndex] as JObject;
                if (tab == null) return;

                string tabName = tab["TabName"]?.ToString() ?? $"Tab {tabIndex}";
                var items = tab["Items"] as JArray ?? new JArray();

                // 3. Confirmation
                if (items.Count > 0)
                {
                    bool result = MessageBoxesManager.ShowTabDeleteConfirmationForm(tabName, items.Count);
                    if (!result)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"User cancelled tab deletion for '{tabName}'");
                        return;
                    }
                }

                // --- Auto-Export Logic (Tabs follow Fence Settings) ---
                if (SettingsManager.ExportShortcutsOnFenceDeletion && items.Count > 0)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Auto-exporting {items.Count} items from tab '{tabName}' before deletion.");

                    int exportCount = 0;
                    foreach (var item in items)
                    {
                        try
                        {
                            CopyPasteManager.SendToDesktop(item);
                            exportCount++;
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Failed to export item from tab: {ex.Message}");
                        }
                    }
                }
                // -----------------------------------------------------------

                // 4. Remove Tab
                tabs.RemoveAt(tabIndex);

                // 5. Calculate New Active Tab
                int currentTab = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");
                int newCurrentTab = currentTab;

                // Shift logic:
                // If we deleted the active tab -> Go to previous (or 0)
                // If we deleted a tab BEFORE the active one -> Active tab shifts down by 1
                // If we deleted a tab AFTER the active one -> Active index stays same
                if (tabIndex <= currentTab)
                {
                    newCurrentTab = Math.Max(0, currentTab - 1);
                }

                // 6. Save Data
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (fenceIndex >= 0)
                {
                    IDictionary<string, object> fenceDict = currentFence is IDictionary<string, object> dict ?
                        dict : ((JObject)currentFence).ToObject<IDictionary<string, object>>();

                    fenceDict["Tabs"] = tabs;
                    fenceDict["CurrentTab"] = newCurrentTab;

                    FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                    FenceDataManager.SaveFenceData();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Deleted tab '{tabName}' from fence '{currentFence.Title}'. New Active Tab: {newCurrentTab}");

                    // --- BUG FIX: FORCE SYNC WITH MAIN ITEMS ---
                    // Now that the tabs have shifted, "Tab 0" might be different.
                    // We must force 'Main.Items' to mirror the NEW 'Tab 0' to prevent "Ghost Icons" 
                    // from appearing if the user later disables tabs.
                    if (tabs.Count > 0)
                    {
                        SynchronizeTab0Content(fenceId, "tab0", "full");
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Forced Tab0 synchronization after deletion.");
                    }
                    // -------------------------------------------

                    // 7. REFRESH UI
                    var updatedFence = FenceDataManager.FenceData[fenceIndex];

                    // A. Refresh the Icons (Content)
                    RefreshFenceContentSimple(fenceWindow, updatedFence, newCurrentTab);

                    // B. Refresh the Tab Buttons (Strip)
                    RefreshTabStripUI(fenceWindow, updatedFence);

                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Tab deletion UI refresh complete");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error deleting tab: {ex.Message}");
            }
        }


        // TABS FEATURE: Switch tab using fence ID
        // FIXED: Now calls RefreshTabStripUI instead of RefreshTabStyling to force a clean redraw.
        public static void SwitchTabByFence(dynamic fence, int newTabIndex, NonActivatingWindow fenceWindow)
        {
            try
            {
                // Get fresh fence data by ID to avoid stale references
                string fenceId = fence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                // Validate tab index
                var tabs = currentFence.Tabs as JArray ?? new JArray();
                if (newTabIndex < 0 || newTabIndex >= tabs.Count) return;

                // Check current tab
                int currentTabIndex = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");
                if (currentTabIndex == newTabIndex) return; // Already there

                // Update Data
                IDictionary<string, object> fenceDict = currentFence as IDictionary<string, object> ??
                    ((JObject)currentFence).ToObject<IDictionary<string, object>>();
                fenceDict["CurrentTab"] = newTabIndex;

                // Save
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (fenceIndex >= 0)
                {
                    FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                    FenceDataManager.SaveFenceData();
                }

                var freshFence = FenceDataManager.FenceData[fenceIndex];

                // 1. Refresh Icons (Content)
                RefreshFenceContentSimple(fenceWindow, freshFence, newTabIndex);

                // 2. Refresh Tabs (UI Redraw) - THE FIX
                // We use RefreshTabStripUI instead of RefreshTabStyling.
                // This destroys the old buttons and creates new ones with the correct "Active" state baked in.
                RefreshTabStripUI(fenceWindow, freshFence);

                string tabName = tabs[newTabIndex]["TabName"]?.ToString() ?? $"Tab {newTabIndex}";
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Switched to tab '{tabName}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error switching tab: {ex.Message}");
            }
        }







        // TABS FEATURE: Move Tab (Left/Right)
        public static void MoveTab(dynamic fence, int fromIndex, int direction, NonActivatingWindow fenceWindow)
        {
            try
            {
                // Direction: -1 for Left, +1 for Right
                int toIndex = fromIndex + direction;

                // 1. Get Fresh Data
                string fenceId = fence.Id?.ToString();
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                var tabs = currentFence.Tabs as JArray;
                if (tabs == null) return;

                // 2. Validate Bounds
                if (toIndex < 0 || toIndex >= tabs.Count) return;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Moving tab {fromIndex} to {toIndex} for fence '{currentFence.Title}'");

                // 3. Swap Tabs
                var tempTab = tabs[fromIndex];
                tabs[fromIndex] = tabs[toIndex];
                tabs[toIndex] = tempTab;

                // 4. Update CurrentTab Index if necessary
                // If we moved the active tab, follow it. 
                // If we moved a tab into the active slot, update the index to stay on the same content.
                int currentTabIdx = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");

                if (currentTabIdx == fromIndex)
                {
                    currentFence.CurrentTab = toIndex; // Follow the moved tab
                }
                else if (currentTabIdx == toIndex)
                {
                    currentFence.CurrentTab = fromIndex; // The other tab swapped into our slot
                }

                // 5. Save & Refresh
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (fenceIndex >= 0)
                {
                    IDictionary<string, object> fenceDict = currentFence is IDictionary<string, object> dict ?
                        dict : ((JObject)currentFence).ToObject<IDictionary<string, object>>();

                    fenceDict["Tabs"] = tabs;
                    fenceDict["CurrentTab"] = currentFence.CurrentTab; // Updated index

                    FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                    FenceDataManager.SaveFenceData();

                    // Refresh Strip (Buttons)
                    RefreshTabStripUI(fenceWindow, FenceDataManager.FenceData[fenceIndex]);

                    // Refresh Content (Icons) - just in case the active tab index changed logic
                    RefreshFenceContentSimple(fenceWindow, FenceDataManager.FenceData[fenceIndex], (int)fenceDict["CurrentTab"]);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error moving tab: {ex.Message}");
            }
        }






        // TABS FEATURE: Tab0-Fence Content Synchronization Manager
        private static bool _isSynchronizing = false; // Prevent circular sync operations

        /// <summary>
        /// Synchronizes content between Tab0 and main Items to ensure they remain identical
        /// Called whenever items are added/removed from either location
        /// </summary>
        /// <param name="fenceId">The fence ID to synchronize</param>
        /// <param name="sourceLocation">Where the change originated: "tab0" or "main"</param>
        /// <param name="operationType">Type of operation: "add", "remove", "full"</param>
        private static void SynchronizeTab0Content(string fenceId, string sourceLocation, string operationType = "full")
        {
            if (_isSynchronizing)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    "Sync already in progress, skipping to prevent circular operation");
                return;
            }

            try
            {
                _isSynchronizing = true;

                var fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (fence == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                        $"Cannot sync: fence {fenceId} not found");
                    return;
                }

                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                var tabs = fence.Tabs as JArray ?? new JArray();
                var mainItems = fence.Items as JArray ?? new JArray();

                // Only sync if tabs are enabled and Tab0 exists
                if (!tabsEnabled || tabs.Count == 0) return;

                var tab0 = tabs[0] as JObject;
                if (tab0 == null) return;

                var tab0Items = tab0["Items"] as JArray ?? new JArray();

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Synchronizing {operationType} from {sourceLocation} for fence '{fence.Title}' - Tab0: {tab0Items.Count} items, Main: {mainItems.Count} items");

                bool syncPerformed = false;

                // Determine sync direction and perform synchronization
                if (sourceLocation == "tab0")
                {
                    // Tab0 changed - sync to main Items
                    if (!AreItemArraysEqual(tab0Items, mainItems))
                    {
                        fence.Items = JArray.FromObject(tab0Items.ToArray());
                        syncPerformed = true;
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Synced {tab0Items.Count} items from Tab0 to main Items for fence '{fence.Title}'");
                    }
                }
                else if (sourceLocation == "main")
                {
                    // Main Items changed - sync to Tab0
                    if (!AreItemArraysEqual(mainItems, tab0Items))
                    {
                        tab0["Items"] = JArray.FromObject(mainItems.ToArray());
                        syncPerformed = true;
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Synced {mainItems.Count} items from main Items to Tab0 for fence '{fence.Title}'");
                    }
                }

                // Save changes if synchronization was performed
                if (syncPerformed)
                {
                    int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                    if (fenceIndex >= 0)
                    {
                        FenceDataManager.FenceData[fenceIndex] = fence;
                        FenceDataManager.SaveFenceData();
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in Tab0 synchronization: {ex.Message}");
            }
            finally
            {
                _isSynchronizing = false;
            }
        }

        /// <summary>
        /// Helper method to compare two JArrays for equality
        /// </summary>
        private static bool AreItemArraysEqual(JArray array1, JArray array2)
        {
            if (array1.Count != array2.Count) return false;

            for (int i = 0; i < array1.Count; i++)
            {
                var item1 = array1[i] as JObject;
                var item2 = array2[i] as JObject;

                if (item1 == null || item2 == null) return false;

                // Compare essential properties (Filename is the key identifier)
                string filename1 = item1["Filename"]?.ToString();
                string filename2 = item2["Filename"]?.ToString();

                if (filename1 != filename2) return false;
            }

            return true;
        }

        // TABS FEATURE: Simplified content refresh that reuses existing infrastructure

        public static void RefreshFenceContentSimple(NonActivatingWindow fenceWindow, dynamic fence, int tabIndex)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"RefreshFenceContentSimple: switching to tab {tabIndex}");

                // 1. Find the WrapPanel
                var border = fenceWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                var scrollViewer = dockPanel?.Children.OfType<ScrollViewer>().FirstOrDefault();
                var wrapPanel = scrollViewer?.Content as WrapPanel;

                if (wrapPanel == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot find WrapPanel for content refresh");
                    return;
                }

                // 2. Clear existing content
                wrapPanel.Children.Clear();

                // 3. Get items from the specified tab
                var tabs = fence.Tabs as JArray ?? new JArray();
                if (tabIndex >= 0 && tabIndex < tabs.Count)
                {
                    var activeTab = tabs[tabIndex] as JObject;
                    if (activeTab != null)
                    {
                        var items = activeTab["Items"] as JArray ?? new JArray();
                        string tabName = activeTab["TabName"]?.ToString() ?? $"Tab {tabIndex}";

                        // 4. Sort items by DisplayOrder
                        var sortedItems = items
                            .OfType<JObject>()
                            .OrderBy(item => item["DisplayOrder"]?.Type == JTokenType.Integer ? item["DisplayOrder"].Value<int>() : 0)
                            .ToList();

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Loading {sortedItems.Count} items from tab '{tabName}'");

                        // 5. Iterate and Add Icons (Using Unified Logic)
                        foreach (dynamic icon in sortedItems)
                        {
                            // FIX A: Use the main FenceManager.AddIcon method!
                            // This ensures network indicators, clean icons, and custom sizes are applied.
                            AddIcon(icon, wrapPanel, fence);

                            StackPanel sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as StackPanel;
                            if (sp != null)
                            {
                                // FIX B: Extract Properties safely
                                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict
                                    ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();

                                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];

                                // FIX C: Extract Arguments
                                string arguments = null;
                                if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk")
                                {
                                    try
                                    {
                                        WshShell shell = new WshShell();
                                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                                        arguments = shortcut.Arguments;
                                    }
                                    catch { }
                                }

                                // FIX D: Attach Events
                                ClickEventAdder(sp, filePath, isFolder, arguments);

                                // FIX E: Attach Centralized Context Menu (Crucial for Right-Click)
                                AttachIconContextMenu(sp, icon, fence, fenceWindow);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in RefreshFenceContentSimple: {ex.Message}");
            }
        }






        // TABS FEATURE: Remove item from current tab
        private static void RemoveItemFromTab(dynamic item, dynamic fence, int tabIndex, WrapPanel wrapPanel)
        {
            try
            {
                var tabs = fence.Tabs as JArray ?? new JArray();
                if (tabIndex >= 0 && tabIndex < tabs.Count)
                {
                    var activeTab = tabs[tabIndex] as JObject;
                    if (activeTab != null)
                    {
                        var items = activeTab["Items"] as JArray ?? new JArray();
                        string itemFilename = item.Filename?.ToString();

                        // Find and remove the item
                        for (int i = items.Count - 1; i >= 0; i--)
                        {
                            var currentItem = items[i] as JObject;
                            if (currentItem != null && currentItem["Filename"]?.ToString() == itemFilename)
                            {
                                items.RemoveAt(i);
                                break;
                            }
                        }

                        // Update the tab
                        activeTab["Items"] = items;

                        // Save changes
                        string fenceId = fence.Id?.ToString();
                        int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                        if (fenceIndex >= 0)
                        {
                            FenceDataManager.FenceData[fenceIndex] = fence;
                            FenceDataManager.SaveFenceData();
                            // ENHANCED: Add Tab0 synchronization after removal
                            if (tabIndex == 0)
                            {
                               // string fenceId = fence.Id?.ToString();
                                if (!string.IsNullOrEmpty(fenceId))
                                {
                                    SynchronizeTab0Content(fenceId, "tab0", "remove");
                                }
                            }

                        }

                        // Refresh the display
                        var fenceWindow = FindVisualParent<NonActivatingWindow>(wrapPanel);
                        if (fenceWindow != null)
                        {
                            RefreshFenceContentSimple(fenceWindow, fence, tabIndex);
                        }

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Removed item '{itemFilename}' from tab {tabIndex}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error removing item from tab: {ex.Message}");
            }
        }

        // TABS FEATURE: Update tab button styling
        // Updated to support the new ScrollViewer structure
        public static void RefreshTabStyling(NonActivatingWindow fenceWindow, int activeTabIndex)
        {
            try
            {
                var border = fenceWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                // 1. Find the Container (Grid)
                var containerGrid = dockPanel.Children.OfType<Grid>()
                    .FirstOrDefault(g => g.Tag?.ToString() == "TAB_STRIP_CONTAINER");

                if (containerGrid == null) return;

                // 2. Find the Tab StackPanel (Inside ScrollViewer)
                var scrollViewer = containerGrid.Children.OfType<ScrollViewer>().FirstOrDefault();
                var tabStack = scrollViewer?.Content as StackPanel;

                // 3. Find the [+] Button (Direct child of Grid)
                var plusButton = containerGrid.Children.OfType<Button>()
                    .FirstOrDefault(b => b.Tag?.ToString() == "ADD_TAB");

                string fenceId = fenceWindow.Tag?.ToString();
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                string fenceColorName = currentFence.CustomColor?.ToString();
                if (string.IsNullOrEmpty(fenceColorName)) fenceColorName = SettingsManager.SelectedColor;

                // 4. Update Tab Buttons
                if (tabStack != null)
                {
                    foreach (var child in tabStack.Children)
                    {
                        if (child is Button tabButton && tabButton.Tag is int idx)
                        {
                            bool isActive = (idx == activeTabIndex);
                            ApplyTabStyle(tabButton, isActive, fenceColorName);
                        }
                    }
                }

                // 5. Update [+] Button
                if (plusButton != null)
                {
                    ApplyTabStyle(plusButton, false, fenceColorName, true);
                }
            }
            catch { }
        }





        // TABS FEATURE: Create custom button template with rounded top corners only
        private static ControlTemplate CreateTabButtonTemplate()
        {
            var template = new ControlTemplate(typeof(Button));

            // Create the border with rounded top corners
            var border = new FrameworkElementFactory(typeof(Border));
            border.Name = "border";
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(4, 4, 0, 0)); // Top corners rounded
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            border.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            border.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            border.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));

            // Create the content presenter
            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.ContentProperty, new TemplateBindingExtension(Button.ContentProperty));

            border.AppendChild(contentPresenter);
            template.VisualTree = border;

            // Add hover trigger for better interaction feedback
            var trigger = new Trigger();
            trigger.Property = Button.IsMouseOverProperty;
            trigger.Value = true;
            trigger.Setters.Add(new Setter(Button.OpacityProperty, 0.8));

            template.Triggers.Add(trigger);

            return template;
        }


        public static void RefreshTabColors(NonActivatingWindow fenceWindow, string newColorName)
        {
            try
            {
                var border = fenceWindow.Content as Border;
                if (border == null) return;
                var dockPanel = border.Child as DockPanel;
                if (dockPanel == null) return;
                var tabStrip = dockPanel.Children.OfType<StackPanel>().FirstOrDefault(sp => sp.Orientation == Orientation.Horizontal && sp.Height == 20);
                if (tabStrip == null) return;

                string fenceId = fenceWindow.Tag?.ToString();
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                int currentTab = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");

                foreach (var child in tabStrip.Children)
                {
                    if (child is Button tabButton)
                    {
                        if (tabButton.Tag is int idx)
                        {
                            bool isActive = (idx == currentTab);
                            ApplyTabStyle(tabButton, isActive, newColorName);
                        }
                        else if (tabButton.Tag?.ToString() == "ADD_TAB")
                        {
                            ApplyTabStyle(tabButton, false, newColorName, true);
                        }
                    }
                }
            }
            catch { }
        }



        // Helper event handlers to avoid lambda capture issues
        private static void TabButton_MouseEnter(object sender, RoutedEventArgs e) { }

        private static void TabButton_MouseLeave(object sender, RoutedEventArgs e) { }

        private static void MigrateLegacyJson()
        {
            try
            {
                bool jsonModified = false;
                var validColors = new HashSet<string> { "Red", "Green", "Teal", "Blue", "Bismark", "White", "Beige", "Gray", "Black", "Purple", "Fuchsia", "Yellow", "Orange" };
                var validEffects = Enum.GetNames(typeof(LaunchEffectsManager.LaunchEffect)).ToHashSet();

                for (int i = 0; i < FenceDataManager.FenceData.Count; i++)
                {
                    dynamic fence = FenceDataManager.FenceData[i];
                    IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict
                        ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();

                    // --- 1. CORE VALIDATION (VITAL) ---
                    // Add GUID if missing
                    if (!fenceDict.ContainsKey("Id"))
                    {
                        fenceDict["Id"] = Guid.NewGuid().ToString();
                        jsonModified = true;
                    }

                    // --- 2. PORTAL FENCE MIGRATION ---
                    if (fence.ItemsType?.ToString() == "Portal")
                    {
                        string portalPath = fence.Items?.ToString();
                        if (!string.IsNullOrEmpty(portalPath) && !System.IO.Directory.Exists(portalPath))
                        {
                            fenceDict["IsFolder"] = true;
                            jsonModified = true;
                        }
                    }
                    else
                    {
                        var items = fence.Items as JArray ?? new JArray();
                        bool itemsModified = false;
                        foreach (var item in items)
                        {
                            var itemDict = item as JObject;
                            if (itemDict != null)
                            {
                                if (itemDict["IsFolder"] == null)
                                {
                                    string path = itemDict["Filename"]?.ToString();
                                    itemDict["IsFolder"] = System.IO.Directory.Exists(path);
                                    itemsModified = true;
                                }
                                // Ensure IsLink exists
                                if (!itemDict.ContainsKey("IsLink"))
                                {
                                    itemDict["IsLink"] = false;
                                    itemsModified = true;
                                }
                                // Ensure IsNetwork exists
                                if (!itemDict.ContainsKey("IsNetwork"))
                                {
                                    string fname = itemDict["Filename"]?.ToString() ?? "";
                                    itemDict["IsNetwork"] = IsNetworkPath(fname);
                                    itemsModified = true;
                                }
                                // Ensure AlwaysRunAsAdmin exists
                                if (!itemDict.ContainsKey("AlwaysRunAsAdmin"))
                                {
                                    itemDict["AlwaysRunAsAdmin"] = false;
                                    itemsModified = true;
                                }
                            }
                        }
                        if (itemsModified)
                        {
                            fenceDict["Items"] = items;
                            jsonModified = true;
                        }
                    }

                    // --- 3. PROPERTY INITIALIZATION (DEFAULTS) ---
                    // Visuals
                    if (!fenceDict.ContainsKey("IsLocked")) { fenceDict["IsLocked"] = "false"; jsonModified = true; }
                    if (!fenceDict.ContainsKey("IsRolled")) { fenceDict["IsRolled"] = "false"; jsonModified = true; }
                    if (!fenceDict.ContainsKey("UnrolledHeight"))
                    {
                        double height = fenceDict.ContainsKey("Height") ? Convert.ToDouble(fenceDict["Height"]) : 130;
                        fenceDict["UnrolledHeight"] = height;
                        jsonModified = true;
                    }

                    if (!fenceDict.ContainsKey("IconSize")) { fenceDict["IconSize"] = "Medium"; jsonModified = true; }
                    if (!fenceDict.ContainsKey("IconSpacing")) { fenceDict["IconSpacing"] = 5; jsonModified = true; }
                    if (!fenceDict.ContainsKey("CustomColor")) { fenceDict["CustomColor"] = null; jsonModified = true; }
                    if (!fenceDict.ContainsKey("CustomLaunchEffect")) { fenceDict["CustomLaunchEffect"] = null; jsonModified = true; }
                    if (!fenceDict.ContainsKey("IsHidden")) { fenceDict["IsHidden"] = "false"; jsonModified = true; }

                    // Text
                    if (!fenceDict.ContainsKey("TextColor")) { fenceDict["TextColor"] = null; jsonModified = true; }
                    if (!fenceDict.ContainsKey("TitleTextColor")) { fenceDict["TitleTextColor"] = null; jsonModified = true; }
                    if (!fenceDict.ContainsKey("TitleTextSize")) { fenceDict["TitleTextSize"] = "Medium"; jsonModified = true; }
                    if (!fenceDict.ContainsKey("BoldTitleText")) { fenceDict["BoldTitleText"] = "false"; jsonModified = true; }
                    if (!fenceDict.ContainsKey("DisableTextShadow")) { fenceDict["DisableTextShadow"] = "false"; jsonModified = true; }
                    if (!fenceDict.ContainsKey("GrayscaleIcons")) { fenceDict["GrayscaleIcons"] = "false"; jsonModified = true; }

                    // Border
                    if (!fenceDict.ContainsKey("FenceBorderColor")) { fenceDict["FenceBorderColor"] = null; jsonModified = true; }
                    if (!fenceDict.ContainsKey("FenceBorderThickness")) { fenceDict["FenceBorderThickness"] = 0; jsonModified = true; }

                    // --- 4. TABS FEATURE (Structure Only, No Merge) ---
                    if (!fenceDict.ContainsKey("TabsEnabled"))
                    {
                        fenceDict["TabsEnabled"] = "false";
                        jsonModified = true;
                    }
                    if (!fenceDict.ContainsKey("CurrentTab"))
                    {
                        fenceDict["CurrentTab"] = 0;
                        jsonModified = true;
                    }
                    if (!fenceDict.ContainsKey("Tabs"))
                    {
                        fenceDict["Tabs"] = new JArray();
                        jsonModified = true;
                    }

                    // --- 5. TAB VALIDATION (Structure Repair) ---
                    // Ensure existing tabs have valid structure (Name, Items Array)
                    // but DO NOT move items around or merge lists.
                    var tabs = fenceDict["Tabs"] as JArray ?? new JArray();
                    if (tabs.Count > 0)
                    {
                        bool tabsModified = false;
                        for (int t = 0; t < tabs.Count; t++)
                        {
                            var tab = tabs[t] as JObject;
                            if (tab == null) continue;

                            if (tab["TabName"] == null)
                            {
                                tab["TabName"] = $"Tab {t + 1}";
                                tabsModified = true;
                            }
                            if (tab["Items"] == null || !(tab["Items"] is JArray))
                            {
                                tab["Items"] = new JArray();
                                tabsModified = true;
                            }
                        }
                        if (tabsModified)
                        {
                            fenceDict["Tabs"] = tabs;
                            jsonModified = true;
                        }
                    }

                    // Update Master List
                    FenceDataManager.FenceData[i] = JObject.FromObject(fenceDict);
                }

                if (jsonModified)
                {
                    FenceDataManager.SaveFenceData();
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Migrated fences.json with updated fields (Validations Only)");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error migrating fences.json: {ex.Message}");
            }
        }


        public static void LoadAndCreateFences(TargetChecker targetChecker)


        {

            // Get current program version from assembly
            string currentVersion = Assembly.GetEntryAssembly()?.GetName().Version?.ToString() ?? "0.0.0"; // This is current version for registry tracking


           
            RegistryHelper.SetProgramManagementValues(currentVersion);



            // === SINGLE INSTANCE CHECK (WITH DISABLE OPTION) ===
            try
            {
                // Check if single instance enforcement is disabled
                if (SettingsManager.DisableSingleInstance)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "SingleInstance: Single instance enforcement disabled. Multiple instances allowed.");
                }
                else
                {
                    // Small delay to ensure process visibility
                    System.Threading.Thread.Sleep(100);

                    Process currentProcess = Process.GetCurrentProcess();
                    string processName = Path.GetFileNameWithoutExtension(currentProcess.ProcessName);
                    Process[] allInstances = Process.GetProcessesByName(processName);

                    // Check if this is a duplicate instance
                    if (allInstances.Length > 1)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                            $"SingleInstance: Duplicate instance detected. Found {allInstances.Length} instances. Writing trigger and exiting.");

                        // Write registry trigger for the original instance
                        bool registryWritten = RegistryHelper.WriteTrigger();

                        if (registryWritten)
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                "SingleInstance: Registry trigger written successfully. Exiting duplicate instance.");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                                "SingleInstance: Failed to write registry trigger. Still exiting duplicate instance.");
                        }

                        Environment.Exit(0);
                        return;
                    }

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "SingleInstance: This is the first instance. Continuing with normal startup.");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"SingleInstance: Error in single instance check: {ex.Message}. Continuing startup.");
            }





            // Initialize FenceDataManager
            FenceDataManager.Initialize();
            string exePath = Assembly.GetEntryAssembly().Location;
            string exeDir = System.IO.Path.GetDirectoryName(exePath);
            FenceDataManager.JsonFilePath = System.IO.Path.Combine(exeDir, "fences.json");
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
                EnableSounds = SettingsManager.EnableSounds,
                TintValue = SettingsManager.TintValue,
                MenuTintValue = SettingsManager.MenuTintValue,
                MenuIcon = SettingsManager.MenuIcon,
                LockIcon = SettingsManager.LockIcon,
                SelectedColor = SettingsManager.SelectedColor,
                IsLogEnabled = SettingsManager.IsLogEnabled,
                singleClickToLaunch = SettingsManager.SingleClickToLaunch,
                LaunchEffect = SettingsManager.LaunchEffect,
                CheckNetworkPaths = false 
            };

            bool jsonLoadSuccessful = false;

            if (System.IO.File.Exists(FenceDataManager.JsonFilePath))
            {
                jsonLoadSuccessful = LoadFenceDataFromJson();
            }

            // If JSON loading failed or file doesn't exist, initialize with defaults
            if (!jsonLoadSuccessful || FenceDataManager.FenceData == null || FenceDataManager.FenceData.Count == 0)
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
            foreach (dynamic fence in FenceDataManager.FenceData.ToList()) // Use ToList to avoid collection modification issues
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
                    FenceDataManager.FenceData.Remove(fence);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Removed Portal Fence '{fence.Title}' from FenceDataManager.FenceData");
                }
                FenceDataManager.SaveFenceData();
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

            foreach (dynamic fence in FenceDataManager.FenceData.ToList())
            {
                CreateFence(fence, targetChecker);
            }



            if (!SettingsManager.DisableSingleInstance)
            {
                StartRegistryMonitor();
            }
            else
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    "RegistryMonitor: Skipped - single instance enforcement disabled");
            }




            // Initialize Auto-Backup Timer
            BackupManager.InitializeAutoBackup();





        }

        private static bool LoadFenceDataFromJson()
        {
            try
            {
                string jsonContent = System.IO.File.ReadAllText(FenceDataManager.JsonFilePath);

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
                    FenceDataManager.FenceData = JsonConvert.DeserializeObject<List<dynamic>>(jsonContent);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                        $"Successfully loaded {FenceDataManager.FenceData?.Count ?? 0} fences from JSON array.");
                    return FenceDataManager.FenceData != null;
                }
                catch (JsonSerializationException)
                {
                    // If that fails, try to parse as a single fence object
                    try
                    {
                        var singleFence = JsonConvert.DeserializeObject<dynamic>(jsonContent);
                        if (singleFence != null)
                        {
                            FenceDataManager.FenceData = new List<dynamic> { singleFence };
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
                string backupPath = FenceDataManager.JsonFilePath + ".corrupted." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                System.IO.File.Copy(FenceDataManager.JsonFilePath, backupPath, true);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    $"Created backup of corrupted JSON file: {backupPath}");
            }
            catch (Exception backupEx)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Failed to create backup of corrupted JSON file: {backupEx.Message}");
            }
        }


        //Replacement on v 2.5 .3. 141


        private static void InitializeDefaultFence()
        {
            try
            {
                // 1. Create the Standard Data Fence (For user icons)
                var dataFence = new
                {
                    Id = Guid.NewGuid().ToString(),
                    Title = "New Fence - Drop your shortcuts here",
                    X = 20.0,
                    Y = 20.0,
                    Width = 360.0,
                    Height = 180.0,
                    ItemsType = "Data",
                    Items = new JArray(),

                    // Defaults
                    IsLocked = "false",
                    IsHidden = "false",
                    IsRolled = "false",
                    UnrolledHeight = 130.0,
                    TabsEnabled = "false",
                    CurrentTab = 0,
                    Tabs = new JArray(),

                    // Visual Defaults
                    CustomColor = (string)null,
                    FenceBorderThickness = 2
                };

                // 2. Create the "Startup Tips" Note Fence
                var noteFence = new
                {
                    Id = Guid.NewGuid().ToString(), // Unique ID
                    Title = "Desktop Fences+ Startup Tips", // Explicit Name
                    X = 20.0,   // Positioned below the data fence
                    Y = 200.0,  // Data fence ends then this fence begins
                    Width = 555.0,
                    Height = 318.0,
                    ItemsType = "Note",
                    Items = new JArray(),

                    // Visuals from your spec
                    CustomColor = (string)null, // Default
                    TextColor = "Teal",         // Teal text

                    // Note Settings
                    NoteContent = "WELCOME TO DESKTOP FENCES +\r\n" +
                                  "---------------------------\r\n" +
                                  "• Roll Up/Down: Double-click the fence title bar.\r\n" +
                                  "• Rename: Ctrl + Click the title bar (Enter to save).\r\n" +
                                  "• Search (SpotSearch): Press Ctrl + ` (Tilde) to find any icon instantly.\r\n" +
                                  "• Options: Click the '♥' menu icon (top-left).\r\n" +
                                  "• Reorder Icons on a Fence: Ctrl + Drag icon to new position.\r\n" +
                                  "• Context Menu: Right-click icons or Fences for more options.\r\n" +
                                  " \r\n" +
                                  "TIP: Ctrl + Click or Ctrl + Right-click, gives even more options.\r\n\r\n" +
                                  "Try customizing this fence! Right-click the title bar -> Customize...",

                    NoteFontSize = "Medium",
                    NoteFontFamily = "Segoe UI",
                    WordWrap = "true",
                    SpellCheck = "false",

                    // Standard Properties
                    IsHidden = "false",
                    IsLocked = "false",
                    IsRolled = "false",
                    UnrolledHeight = 318.0,
                    BoldTitleText = "false",
                    DisableTextShadow = "false",
                    IconSize = "Medium",
                    IconSpacing = 5,
                    FenceBorderThickness = 2
                };

                // 3. Combine and Save
                var fences = new List<object> { dataFence, noteFence };
                string defaultJson = JsonConvert.SerializeObject(fences, Formatting.Indented);

                System.IO.File.WriteAllText(FenceDataManager.JsonFilePath, defaultJson);

                // Load into memory
                FenceDataManager.FenceData = JsonConvert.DeserializeObject<List<dynamic>>(defaultJson);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    "Initialized default configuration with Data Fence and Startup Tips");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error initializing default fences: {ex.Message}");

                // Fallback to empty list if something explodes
                FenceDataManager.FenceData = new List<dynamic>();
            }
        }


        public static void CreateFence(dynamic fence, TargetChecker targetChecker)
        {

            // --- FIX: Declare Title TextBox EARLY ---
            TextBox titletb = new TextBox
            {
                HorizontalContentAlignment = HorizontalAlignment.Center,
                Visibility = Visibility.Collapsed,
                Background = System.Windows.Media.Brushes.White,
                Foreground = System.Windows.Media.Brushes.Black,
                Padding = new Thickness(2)
            };

            // --- NEW: Declare Commit Action for robust saving ---
            Action CommitRename = null;
            // ---------------------------------------------------

            // Check for valid Portal Fence target folder
            if (fence.ItemsType?.ToString() == "Portal")
            {
                string targetPath = fence.Path?.ToString();
                if (string.IsNullOrEmpty(targetPath) || !System.IO.Directory.Exists(targetPath))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation, $"Skipping creation of Portal Fence '{fence.Title}' due to missing target folder: {targetPath ?? "null"}");
                    FenceDataManager.FenceData.Remove(fence);
                    FenceDataManager.SaveFenceData();
                    return;
                }
            }
            DockPanel dp = new DockPanel();
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
            // Add Double Click Handler
            cborder.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ClickCount == 2)
                {
                    SearchFormManager.ToggleSearch();
                    e.Handled = true;
                }
            };


            //  Add heart symbol in top-left corner
            string MenuSymbol = "♥";

            if (SettingsManager.MenuIcon == 0)
            {
                MenuSymbol = "♥";
              
            }
            else if (SettingsManager.MenuIcon == 1)
            {
                MenuSymbol = "☰";
            }
            else if (SettingsManager.MenuIcon == 2)
            {
                MenuSymbol = "≣";
            }
            else if (SettingsManager.MenuIcon == 3)
            {
                MenuSymbol = "𓃑";
            }

            TextBlock heart = new TextBlock
            {
                // Text = "♥",

                Name = "FenceMenuIcon", // New! Name
                Text = MenuSymbol,
                FontSize = 22,
                Foreground = System.Windows.Media.Brushes.White, // Match title and icon text color
                Margin = new Thickness(5, -3, 0, 0), // Position top-left, aligned with title
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                Cursor = Cursors.Hand,
                Opacity = (double)SettingsManager.MenuTintValue / 100 // 0.3 // Lower tint by default

            };
  
            _heartTextBlocks[fence] = heart;



            heart.MouseEnter += (s, e) =>
            {
                // Remove previous animation 
                heart.BeginAnimation(UIElement.OpacityProperty, null);

                heart.Opacity = 1.0;
            };

            heart.MouseLeave += (s, e) =>
            {
                double targetOpacity = (double)SettingsManager.MenuTintValue / 100;

                DoubleAnimation fadeBack = new DoubleAnimation
                {
                    From = 1.0,
                    To = targetOpacity,
                    Duration = TimeSpan.FromMilliseconds(300),
                    BeginTime = TimeSpan.FromMilliseconds(800)
                };

                heart.BeginAnimation(UIElement.OpacityProperty, fadeBack);
            };


            dp.Children.Add(heart);
            Panel.SetZIndex(heart, 100); // Ensure heart is above titleGrid to receive clicks

            // Store heart TextBlock reference for this fence
            _heartTextBlocks[fence] = heart;
            // Create and assign heart ContextMenu using centralized builder
            // heart.ContextMenu = BuildHeartContextMenu(fence);
            heart.ContextMenu = BuildHeartContextMenu(fence, false); // Normal menu by default
                                                                     //// Handle left-click to open heart context menu
                                                                     //heart.MouseLeftButtonDown += (s, e) =>
                                                                     //{
                                                                     // if (e.ChangedButton == MouseButton.Left && heart.ContextMenu != null)
                                                                     // {
                                                                     // heart.ContextMenu.IsOpen = true;
                                                                     // e.Handled = true;
                                                                     // }
                                                                     //};
                                                                     // Handle left-click to open heart context menu
            heart.MouseLeftButtonDown += (s, e) =>
            {


                // FIX: Directly call CommitRename logic
                if (titletb.IsVisible)
                {
                    CommitRename?.Invoke();
                    return;
                }


                if (e.ChangedButton == MouseButton.Left && heart.ContextMenu != null)
                {
                    // Check if Ctrl is pressed for extended menu
                    bool isCtrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                    // DEBUG: Log the Ctrl state
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Heart clicked - Ctrl pressed: {isCtrlPressed}");
                    // Rebuild context menu based on Ctrl state
                    heart.ContextMenu = BuildHeartContextMenu(fence, isCtrlPressed);
                    heart.ContextMenu.IsOpen = true;
                    e.Handled = true;
                }
            };
            // Add a protection symbol in top-right corner


            string LockSymbol = "🛡️";

            if (SettingsManager.LockIcon == 0)
            {
                LockSymbol = "🛡️";
            }
            else if (SettingsManager.LockIcon == 1)
            {
                LockSymbol = "🔑";
            }
            else if (SettingsManager.LockIcon == 2)
            {
                LockSymbol = "🔐";
            }
            else if (SettingsManager.LockIcon == 3)
            {
                LockSymbol = "🔒";
            }

            //MessageBox.Show(SettingsManager.LockIcon +" " + LockSymbol.ToString());


            TextBlock lockIcon = new TextBlock
           
            {

                //     Text = "🔐",
                //     Text = "🔑",
                //     Text = "🔒",
                //     Text = "🔓",
                Name = "FenceLockIcon", // New! Name
                Text = LockSymbol,//"🛡️",
                                FontSize = 14,
                Foreground = fence.IsLocked?.ToString().ToLower() == "true" ? System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.White,
                Margin = new Thickness(0, 3, 2, 0), // Adjusted for top-right positioning
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Cursor = Cursors.Hand,
                ToolTip = fence.IsLocked?.ToString().ToLower() == "true" ? "Fence is locked (click to unlock)" : "Fence is unlocked (click to lock)",
                Opacity = (double)SettingsManager.MenuTintValue / 100 // 0.3 // Lower tint by default
            };

  
            lockIcon.MouseEnter += (s, e) =>
            {
                // Remove previous animation
                lockIcon.BeginAnimation(UIElement.OpacityProperty, null);

                lockIcon.Opacity = 1.0;
            };

            lockIcon.MouseLeave += (s, e) =>
            {
                double targetOpacity = (double)SettingsManager.MenuTintValue / 100;

                DoubleAnimation fadeBack = new DoubleAnimation
                {
                    From = 1.0,
                    To = targetOpacity,
                    Duration = TimeSpan.FromMilliseconds(300),
                    BeginTime = TimeSpan.FromMilliseconds(800)
                };

                lockIcon.BeginAnimation(UIElement.OpacityProperty, fadeBack);
            };




            // Set initial state without saving to JSON
            UpdateLockState(lockIcon, fence, null, saveToJson: false);
            // Lock icon click handler
            lockIcon.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ChangedButton == MouseButton.Left)
                {

                    // FIX: Directly call CommitRename logic
                    if (titletb.IsVisible)
                    {
                        CommitRename?.Invoke();
                        return;
                    }

                    // Get the NonActivatingWindow to find the fence by Id
                    NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(lockIcon);
                    if (win == null)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Could not find NonActivatingWindow for lock icon click in fence '{fence.Title}'");
                        return;
                    }
                    // Find the fence in FenceDataManager.FenceData using the window's Tag (Id)
                    string fenceId = win.Tag?.ToString();
                    if (string.IsNullOrEmpty(fenceId))
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Fence Id is missing for window '{win.Title}'");
                        return;
                    }
                    dynamic currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    if (currentFence == null)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Fence with Id '{fenceId}' not found in FenceDataManager.FenceData");
                        return;
                    }
                    // Toggle the lock state
                    bool currentState = currentFence.IsLocked?.ToString().ToLower() == "true";
                    bool newState = !currentState;
                    // Update UI and JSON on the main thread
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
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
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0, GridUnitType.Pixel) }); // Col 0: Spacer
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Col 1: Title
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });                      // Col 2: Filter Icon (Auto width)
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30, GridUnitType.Pixel) }); // Col 3: Lock Icon
                                                                                                                      // End of ctrl+click handler
            ContextMenu CnMnFenceManager = new ContextMenu();
            // MenuItem miNewFence = new MenuItem { Header = "New Fence" };
            // MenuItem miNewPortalFence = new MenuItem { Header = "New Portal Fence" };
            MenuItem miNewNoteFence = new MenuItem { Header = "New Note Fence" };
            // MenuItem miRemoveFence = new MenuItem { Header = "Delete Fence" };
            // MenuItem miXT = new MenuItem { Header = "Exit" };
            MenuItem miHide = new MenuItem { Header = "Hide Fence" }; // New Hide Fence item
                                                                      // CnMnFenceManager.Items.Add(miRemoveFence);
            CnMnFenceManager.Items.Add(miHide); // Add Hide Fence
                                                // Add Note fence specific context menu items if this is a Note fence
            if (fence.ItemsType?.ToString() == "Note")
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                    $"Adding Note fence context menu items for '{fence.Title}'");
                // We'll add the TextBox reference after the window is created
                // For now, just mark that this will need Note menu items
            }
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
            // Add Note fence specific context menu items after window creation
            if (fence.ItemsType?.ToString() == "Note")
            {
                // The TextBox will be created in InitContent(), so we need to add menu items after that
                // We'll modify the context menu after InitContent() is called
            }
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
                    Top = win.Top + (win.Height - 40) / 2 // Center vertically
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



            // --- DYNAMIC MENU ITEMS ---

            // 1. Clear Dead Shortcuts (Initially Hidden)
            MenuItem miClearDeadShortcuts = null;
            Separator sepClearDead = null;

            if (fence.ItemsType?.ToString() == "Data")
            {
                sepClearDead = new Separator { Visibility = Visibility.Collapsed };
                CnMnFenceManager.Items.Add(sepClearDead);

                miClearDeadShortcuts = new MenuItem { Header = "Clear Dead Shortcuts", Visibility = Visibility.Collapsed };
                CnMnFenceManager.Items.Add(miClearDeadShortcuts);

                miClearDeadShortcuts.Click += (s, e) =>
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Clear Dead Shortcuts clicked for fence '{fence.Title}'");
                    try
                    {
                        int removedCount = FilePathUtilities.ClearDeadShortcutsFromFence(fence);
                        if (removedCount > 0)
                        {
                            FenceDataManager.SaveFenceData();
                            RefreshFenceUsingFormApproach(win, fence);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate, $"Error in Clear Dead Shortcuts: {ex.Message}");
                    }
                };
            }

            // 2. Open Folder (Portal Only)
            if (fence.ItemsType?.ToString() == "Portal")
            {
                MenuItem miOpenFolder = new MenuItem { Header = "Open fence folder" };
                miOpenFolder.Click += (s, e) =>
                {
                    try
                    {
                        string folderPath = fence.Path?.ToString();
                        if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                            Process.Start("explorer.exe", folderPath);
                    }
                    catch { }
                };
                CnMnFenceManager.Items.Add(miOpenFolder);
            }

            CnMnFenceManager.Items.Add(new Separator());

            // 3. Paste Item (Initially Hidden)
            MenuItem miPasteItem = new MenuItem { Header = "Paste Item", Visibility = Visibility.Collapsed };
            miPasteItem.Click += (s, e) => CopyPasteManager.PasteItem(fence);
            CnMnFenceManager.Items.Add(miPasteItem);

          










            //  CnMnFenceManager.Items.Add(new Separator());

            //CnMnFenceManager.Items.Add(miNewCustomize);
            MenuItem miCustomize = new MenuItem { Header = "Customize..." };
            miCustomize.Click += (s, e) =>
            {
                try
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Opening CustomizeFenceFormManager for fence '{fence.Title}'");
                    var customizeForm = new CustomizeFenceFormManager(fence);
                    customizeForm.ShowDialog();
                    if (customizeForm.DialogResult)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"User saved changes in CustomizeFenceFormManager for fence '{fence.Title}'");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"User cancelled CustomizeFenceFormManager for fence '{fence.Title}'");
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error opening CustomizeFenceFormManager: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error opening customize form: {ex.Message}", "Form Error");
                }
            };
            CnMnFenceManager.Items.Add(miCustomize);
            // CnMnFenceManager.Items.Add(new Separator());
            // CnMnFenceManager.Items.Add(new Separator());
            // CnMnFenceManager.Items.Add(miXT);
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
                double targetHeight = 28; // Default for rolled-up state  //rolled height
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

                    return;
                }

                // Skip updates if this fence is currently in a rollup/rolldown transition
                if (_fencesInTransition.Contains(fenceId))
                {

                    return;
                }
                // Find the current fence in FenceDataManager.FenceData using ID
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence != null)
                {

                    double newHeight = e.NewSize.Height;
                    double newWidth = e.NewSize.Width;
                    double oldHeight = Convert.ToDouble(currentFence.Height?.ToString() ?? "0");
                    double oldUnrolledHeight = Convert.ToDouble(currentFence.UnrolledHeight?.ToString() ?? "0");
                    bool isRolled = currentFence.IsRolled?.ToString().ToLower() == "true";

                    // Update Width and Height with the actual new values
                    currentFence.Width = newWidth;
                    currentFence.Height = newHeight;

                    // Handle UnrolledHeight update
                    if (!isRolled)
                    {

                        if (Math.Abs(newHeight - 28) > 5) // Only if height is significantly different from rolled-up height   //rolled height
                        {
                            double heightDifference = Math.Abs(newHeight - oldUnrolledHeight);
                            // DebugLog("LOGIC", fenceId, $"Height difference from old UnrolledHeight: {heightDifference:F1}");
                            currentFence.UnrolledHeight = newHeight;
                            // DebugLog("UPDATE", fenceId, $"UPDATED UnrolledHeight from {oldUnrolledHeight:F1} to {newHeight:F1}");
                        }
                        else
                        {

                        }
                    }
                    else
                    {

                    }
                    // Save to JSON
                 //   MessageBox.Show("Debug: SizeChanged handler called. Saving fence data.");
                    FenceDataManager.SaveFenceData();
                    var verifyFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    double verifyHeight = Convert.ToDouble(verifyFence.Height?.ToString() ?? "0");
                    double verifyUnrolled = Convert.ToDouble(verifyFence.UnrolledHeight?.ToString() ?? "0");

                }
                else
                {

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


               


            };
            // Make window focusable for key events during drag
            win.Focusable = true;
            win.Show();





            // --- FIX START ---

            // 1. Install the Message Hook (Enforces the 90% size limit)
            win.SourceInitialized += (s, e) =>
            {
                var source = HwndSource.FromHwnd(new WindowInteropHelper(win).Handle);
                source?.AddHook(WndProc);
            };
            // (Safety: if already initialized)
            if (PresentationSource.FromVisual(win) is HwndSource existingSource)
            {
                existingSource.AddHook(WndProc);
            }

            // 2. Reactive State Fix (The "Trick")
            // If Snap Assist sets state to Maximized, we immediately force it back to Normal.
            // This ensures the Resize Grip (bottom-right) remains visible and functional.
            win.StateChanged += (sender, args) =>
            {
                if (win.WindowState == WindowState.Maximized)
                {
                    // We use Dispatcher to let the OS finish its "snap" calculation first, 
                    // then we immediately override the mode back to Normal.
                    win.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        win.WindowState = WindowState.Normal;

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            "Snap Assist intercepted: Reverted Maximized state to Normal to preserve controls.");
                    }));
                }
            };


            // --- FIX END ---




            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
        $"Fence '{fence.Title}' successfully created and displayed");

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


            // --- DYNAMIC STATE UPDATES (ON MENU OPEN) ---
            MenuItem miExportAllToDesktop = null;
            MenuItem miNameAfterPath = null; // New: For Portal Renaming
            Separator sepNameAfterPath = null; // New: Separator for layout

            CnMnFenceManager.Opened += (contextSender, contextArgs) =>
            {
                // A. Update Paste Visibility
                bool hasCopiedItem = CopyPasteManager.HasCopiedItem();
                miPasteItem.Visibility = hasCopiedItem ? Visibility.Visible : Visibility.Collapsed;

                // B. Update Clear Dead Shortcuts Visibility
                if (miClearDeadShortcuts != null)
                {
                    bool hasDead = HasDeadShortcuts(fence);
                    miClearDeadShortcuts.Visibility = hasDead ? Visibility.Visible : Visibility.Collapsed;
                    if (sepClearDead != null) sepClearDead.Visibility = hasDead ? Visibility.Visible : Visibility.Collapsed;
                }

                // --- CTRL + RIGHT CLICK LOGIC ---
                bool isCtrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                bool isDataFence = fence.ItemsType?.ToString() == "Data";
                bool isPortalFence = fence.ItemsType?.ToString() == "Portal";

                // Clean up previous dynamic items to prevent duplicates
                if (miExportAllToDesktop != null && CnMnFenceManager.Items.Contains(miExportAllToDesktop))
                    CnMnFenceManager.Items.Remove(miExportAllToDesktop);

                if (miNameAfterPath != null && CnMnFenceManager.Items.Contains(miNameAfterPath))
                    CnMnFenceManager.Items.Remove(miNameAfterPath);

                if (sepNameAfterPath != null && CnMnFenceManager.Items.Contains(sepNameAfterPath))
                    CnMnFenceManager.Items.Remove(sepNameAfterPath);

                // C. Export All (Data Fence + Ctrl)
                if (isCtrlPressed && isDataFence)
                {
                    miExportAllToDesktop = new MenuItem { Header = "Export all icons to desktop" };
                    miExportAllToDesktop.Click += (s, e) => ExportAllIconsToDesktop(fence);

                    // Insert before Customize (safe lookup)
                    int insertIndex = CnMnFenceManager.Items.Count - 1;
                    var customizeItem = CnMnFenceManager.Items.OfType<MenuItem>()
                        .FirstOrDefault(m => m.Header.ToString() == "Customize...");
                    if (customizeItem != null) insertIndex = CnMnFenceManager.Items.IndexOf(customizeItem);

                    CnMnFenceManager.Items.Insert(insertIndex, miExportAllToDesktop);
                }

                // D. Name After Target (Portal Fence + Ctrl)
                if (isCtrlPressed && isPortalFence)
                {
                    miNameAfterPath = new MenuItem { Header = "Name Fence After Target Path" };
                    miNameAfterPath.Click += (s, e) =>
                    {
                        // Get Base Path (Not navigation path)
                        string targetPath = fence.Path?.ToString();

                        if (!string.IsNullOrEmpty(targetPath))
                        {
                            // 1. Update Live Object (Global List)
                            string id = fence.Id?.ToString();
                            var liveFence = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == id);

                            if (liveFence != null)
                            {
                                if (liveFence is JObject jFence) jFence["Title"] = targetPath;
                                else liveFence.Title = targetPath;

                                fence.Title = targetPath; // Update local reference
                            }

                            // 2. Update UI
                            titlelabel.Content = targetPath;
                            win.Title = targetPath;
                            titletb.Text = targetPath; // Update hidden textbox too

                            // 3. Save
                            FenceDataManager.SaveFenceData();
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Renamed portal fence to target: {targetPath}");
                        }
                    };

                    sepNameAfterPath = new Separator();

                    // Insert before Customize
                    int insertIndex = CnMnFenceManager.Items.Count - 1;
                    var customizeItem = CnMnFenceManager.Items.OfType<MenuItem>()
                        .FirstOrDefault(m => m.Header.ToString() == "Customize...");

                    if (customizeItem != null)
                    {
                        insertIndex = CnMnFenceManager.Items.IndexOf(customizeItem);
                        // Insert Order: Name -> Separator -> Customize
                        CnMnFenceManager.Items.Insert(insertIndex, miNameAfterPath);
                        CnMnFenceManager.Items.Insert(insertIndex + 1, sepNameAfterPath);
                    }
                    else
                    {
                        CnMnFenceManager.Items.Add(miNameAfterPath);
                        CnMnFenceManager.Items.Add(sepNameAfterPath);
                    }
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Updated fence menu - Ctrl:{isCtrlPressed}");
            };





            Grid.SetColumn(titletb, 1);
            titleGrid.Children.Add(titletb);





            // --- FILTER UI START ---
            // Only add Filter UI for Portal Fences
            Grid filterBar = null;
            if (fence.ItemsType?.ToString() == "Portal")
            {
                // 1. The Filter Icon (Funnel)
                TextBlock filterIcon = new TextBlock
                {
                    //Text = "🌪",
                    Name = "FenceFilterIcon", // New! Name
                    Text = "❖",
                    FontSize = 18,
                    Foreground = System.Windows.Media.Brushes.White,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Cursor = Cursors.Hand,
                    Margin = new Thickness(0, 0, 5, 0),
                    ToolTip = "Filter files (e.g. '*.jpg' or '>*.tmp' to exclude)"
                };

                Grid.SetColumn(filterIcon, 2);
                titleGrid.Children.Add(filterIcon);




                // 2. The Filter Bar (Stable Version)
                filterBar = new Grid
                {
                    Height = 26,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(60, 0, 0, 0)),
                    Visibility = Visibility.Collapsed,
                    Margin = new Thickness(0, 0, 0, 2)
                };

                filterBar.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                filterBar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                ComboBox cmbFilter = new ComboBox
                {
                    IsEditable = true,
                    Height = 24,
                    StaysOpenOnEdit = true,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Padding = new Thickness(5, 0, 5, 0),
                    Text = GetSafeProperty(fence, "FilterString")
                };

                // Helper: Only Repopulate when needed (prevents freeze loop)
                Action repopulateDropdown = () =>
                {
                    cmbFilter.Items.Clear();

                    // A. History
                    try
                    {
                        string currentId = fence.Id?.ToString();
                        var liveFence = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == currentId);
                        if (liveFence != null)
                        {
                            Newtonsoft.Json.Linq.JArray history = null;
                            if (liveFence is Newtonsoft.Json.Linq.JObject jFence)
                                history = jFence["FilterHistory"] as Newtonsoft.Json.Linq.JArray;
                            else
                                history = liveFence.GetType().GetProperty("FilterHistory")?.GetValue(liveFence) as Newtonsoft.Json.Linq.JArray;

                            if (history != null)
                            {
                                foreach (var item in history)
                                {
                                    cmbFilter.Items.Add(new ComboBoxItem
                                    {
                                        Content = item.ToString(),
                                        Tag = item.ToString(),
                                        FontWeight = FontWeights.Bold
                                    });
                                }
                            }
                        }
                    }
                    catch { }

                    // B. Separator (if needed)
                    if (cmbFilter.Items.Count > 0) cmbFilter.Items.Add(new Separator());

                    // C. Standard Presets (From static dictionary)
                    foreach (var kvp in _standardPresets)
                    {
                        cmbFilter.Items.Add(new ComboBoxItem { Content = kvp.Key, Tag = kvp.Value });
                    }
                };

                // FIX 1: Only populate on OPEN, never during selection/typing
                cmbFilter.DropDownOpened += (s, e) => repopulateDropdown();

                // Clear Button
                Button btnClearFilter = new Button
                {
                    Content = "✕",
                    Background = System.Windows.Media.Brushes.Transparent,
                    Foreground = System.Windows.Media.Brushes.Gray,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    Width = 20,
                    ToolTip = "Clear filter and close"
                };

                filterBar.Children.Add(cmbFilter);
                filterBar.Children.Add(btnClearFilter);
                Grid.SetColumn(btnClearFilter, 1);

                // --- LOGIC SETUP ---

                // Helper to execute/save
                Action<string> commitFilter = (text) =>
                {
                    filterIcon.Foreground = string.IsNullOrWhiteSpace(text)
                        ? System.Windows.Media.Brushes.White
                        : System.Windows.Media.Brushes.Orange;

                    if (_portalFences.ContainsKey(fence))
                        _portalFences[fence].ApplyFilter(text);

                    UpdateFenceProperty(fence, "FilterString", text, "Updated filter");

                    // Update history (logic now handles ignoring presets)
                    UpdateFilterHistory(fence, text);
                };

                // Typing Timer
                System.Windows.Threading.DispatcherTimer debounceTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(500)
                };
                debounceTimer.Tick += (s, e) =>
                {
                    debounceTimer.Stop();
                    // Visual update only while typing
                    if (_portalFences.ContainsKey(fence)) _portalFences[fence].ApplyFilter(cmbFilter.Text);

                    filterIcon.Foreground = string.IsNullOrWhiteSpace(cmbFilter.Text)
                        ? System.Windows.Media.Brushes.White
                        : System.Windows.Media.Brushes.Orange;
                };

                // 1. STYLE FIX (Internal TextBox)
                cmbFilter.Loaded += (s, e) =>
                {
                    var textBox = (TextBox)cmbFilter.Template.FindName("PART_EditableTextBox", cmbFilter);
                    if (textBox != null)
                    {
                        textBox.Background = System.Windows.Media.Brushes.Transparent;
                        textBox.Foreground = System.Windows.Media.Brushes.White;
                        textBox.CaretBrush = System.Windows.Media.Brushes.White;
                        textBox.BorderThickness = new Thickness(0);
                    }
                };

                // 2. TYPING LOGIC
                cmbFilter.KeyUp += (s, e) =>
                {
                    if (e.Key == Key.Up || e.Key == Key.Down) return;

                    if (e.Key == Key.Enter)
                    {
                        commitFilter(cmbFilter.Text); // Save history on Enter
                        cmbFilter.IsDropDownOpen = false;
                        filterBar.Visibility = Visibility.Collapsed;
                        win.ShowActivated = false;
                        Keyboard.ClearFocus();
                        e.Handled = true;
                        return;
                    }
                    else if (e.Key == Key.Escape)
                    {
                        filterBar.Visibility = Visibility.Collapsed;
                        win.ShowActivated = false;
                        Keyboard.ClearFocus();
                        e.Handled = true;
                        return;
                    }

                    debounceTimer.Stop();
                    debounceTimer.Start();
                };

                // 3. SELECTION LOGIC (FIXED)
                // We do NOT repopulate here. We just apply the value.
                cmbFilter.SelectionChanged += (s, e) =>
                {
                    if (cmbFilter.SelectedItem is ComboBoxItem item && item.Tag != null)
                    {
                        string selectedFilter = item.Tag.ToString();

                        // Break the event loop by running this later
                        System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            cmbFilter.IsDropDownOpen = false;

                            // Important: Don't set SelectedItem = null here, it causes flickering.
                            // Just force the text and run logic.
                            cmbFilter.Text = selectedFilter;

                            debounceTimer.Stop();
                            commitFilter(selectedFilter);
                        }));
                    }
                };

                // 4. CLEAR LOGIC
                btnClearFilter.Click += (s, e) =>
                {
                    cmbFilter.Text = "";
                    if (_portalFences.ContainsKey(fence)) _portalFences[fence].ApplyFilter("");
                    UpdateFenceProperty(fence, "FilterString", "", "Cleared filter");
                    filterIcon.Foreground = System.Windows.Media.Brushes.White;

                    filterBar.Visibility = Visibility.Collapsed;
                    win.ShowActivated = false;
                };

                // 5. TOGGLE BAR
                filterIcon.MouseLeftButtonDown += (s, e) =>
                {
                    if (filterBar.Visibility == Visibility.Visible)
                    {
                        if (!string.IsNullOrWhiteSpace(cmbFilter.Text)) commitFilter(cmbFilter.Text); // Save on close
                        filterBar.Visibility = Visibility.Collapsed;
                        win.ShowActivated = false;
                        Keyboard.ClearFocus();
                    }
                    else
                    {
                        filterBar.Visibility = Visibility.Visible;
                        win.ShowActivated = true;
                        win.Activate();

                        var textBox = (TextBox)cmbFilter.Template.FindName("PART_EditableTextBox", cmbFilter);
                        if (textBox != null) textBox.Focus();
                        else cmbFilter.Focus();
                    }
                    e.Handled = true;
                };

                // Initial Apply
                if (!string.IsNullOrEmpty(cmbFilter.Text))
                {
                    filterIcon.Foreground = System.Windows.Media.Brushes.Orange;
                    win.Loaded += (s, e) =>
                    {
                        if (_portalFences.ContainsKey(fence))
                            _portalFences[fence].ApplyFilter(cmbFilter.Text);
                    };
                }
            }
            // --- FILTER UI END ---


            // --- STEP 4 FIX: Centralized Rename Logic ---

            // 1. Define the Logic ONCE
            CommitRename = () =>
            {
                // Only run if actually editing
                if (titletb.Visibility != Visibility.Visible) return;

                string originalTitle = fence.Title.ToString();
                string newTitle = titletb.Text;
                string finalTitle = InterCore.ProcessTitleChange(fence, newTitle, originalTitle);

                // Update Data
                string id = fence.Id?.ToString();
                var liveFence = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == id);

                if (liveFence != null)
                {
                    if (liveFence is Newtonsoft.Json.Linq.JObject jFence)
                        jFence["Title"] = finalTitle;
                    else
                        liveFence.Title = finalTitle;
                    fence.Title = finalTitle;
                }

                // Update UI
                titlelabel.Content = finalTitle;
                win.Title = finalTitle;
                titletb.Visibility = Visibility.Collapsed;
                titlelabel.Visibility = Visibility.Visible;

                // Save
                FenceDataManager.SaveFenceData();
                win.ShowActivated = false;

                // Clear visual focus
                Keyboard.ClearFocus();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Rename committed: {finalTitle}");
            };

            // 2. Wire up Events to use the central logic
            titletb.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    CommitRename(); // Call shared logic
                    e.Handled = true;
                }
                else if (e.Key == Key.Escape)
                {
                    // Cancel Logic
                    titletb.Text = fence.Title.ToString();
                    titletb.Visibility = Visibility.Collapsed;
                    titlelabel.Visibility = Visibility.Visible;
                    win.ShowActivated = false;
                    Keyboard.ClearFocus();
                    e.Handled = true;
                }
            };

            titletb.LostFocus += (sender, e) =>
            {
                CommitRename(); // Call shared logic
            };



            // --- STEP 4 START: Configure and Add Events ---
            titletb.HorizontalContentAlignment = HorizontalAlignment.Center;
            titletb.Visibility = Visibility.Collapsed;

            // 1. Handle Keys (Enter = Save, Escape = Cancel)
            titletb.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    string originalTitle = fence.Title.ToString();
                    string newTitle = titletb.Text;
                    string finalTitle = InterCore.ProcessTitleChange(fence, newTitle, originalTitle);

                    // Update LIVE Data
                    string id = fence.Id?.ToString();
                    var liveFence = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == id);

                    if (liveFence != null)
                    {
                        if (liveFence is Newtonsoft.Json.Linq.JObject jFence)
                            jFence["Title"] = finalTitle;
                        else
                            liveFence.Title = finalTitle;
                        fence.Title = finalTitle;
                    }

                    titlelabel.Content = finalTitle;
                    win.Title = finalTitle;
                    titletb.Visibility = Visibility.Collapsed;
                    titlelabel.Visibility = Visibility.Visible;

                    FenceDataManager.SaveFenceData();

                    win.ShowActivated = false;
                    Keyboard.ClearFocus(); // Drop focus
                    win.Focus();
                }
      
     else if (e.Key == Key.Escape)
                    {
                        // ESCAPE: Cancel and Revert
                        titletb.Text = fence.Title.ToString();
                        titletb.Visibility = Visibility.Collapsed;
                        titlelabel.Visibility = Visibility.Visible;

                        win.ShowActivated = false;
                        win.Focus(); // Drop focus from textbox
                        e.Handled = true;

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Rename cancelled via Escape");
                    }
                }
                ;

            // 2. Handle Focus Loss (Save when clicking away)
            titletb.LostFocus += (sender, e) =>
            {
                // If invisible, we already handled it (e.g. via Escape)
                if (titletb.Visibility != Visibility.Visible) return;

                string originalTitle = fence.Title.ToString();
                string newTitle = titletb.Text;
                string finalTitle = InterCore.ProcessTitleChange(fence, newTitle, originalTitle);

                string id = fence.Id?.ToString();
                var liveFence = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == id);

                if (liveFence != null)
                {
                    if (liveFence is Newtonsoft.Json.Linq.JObject jFence)
                        jFence["Title"] = finalTitle;
                    else
                        liveFence.Title = finalTitle;
                    fence.Title = finalTitle;
                }

                titlelabel.Content = finalTitle;
                win.Title = finalTitle;
                titletb.Visibility = Visibility.Collapsed;
                titlelabel.Visibility = Visibility.Visible;

                FenceDataManager.SaveFenceData();

                win.ShowActivated = false;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Rename saved via LostFocus: {finalTitle}");
            };
            // --- STEP 4 END ---


            // Move lockIcon to the Grid
            Grid.SetColumn(lockIcon, 3); // Moved to Column 3
            Grid.SetRow(lockIcon, 0);
            titleGrid.Children.Add(lockIcon);
            // Add the titleGrid to the DockPanel
            DockPanel.SetDock(titleGrid, Dock.Top);
            dp.Children.Add(titleGrid);
            if (filterBar != null)
            {
                DockPanel.SetDock(filterBar, Dock.Top);
                dp.Children.Add(filterBar);
            }


            // TABS FEATURE: Add TabStrip
            // We delegate to the shared method to ensure consistent "Scrollable" layout on startup.
            RefreshTabStripUI(win, fence);

       
            string fenceId = win.Tag?.ToString();
            if (string.IsNullOrEmpty(fenceId))
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Fence Id is missing for window '{win.Title}'");
                return;
            }
            dynamic currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
            if (currentFence == null)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Fence with Id '{fenceId}' not found in FenceDataManager.FenceData");
                return;
            }
            bool isLocked = currentFence.IsLocked?.ToString().ToLower() == "true";
            titlelabel.MouseDown += (sender, e) =>
            {
                // FIX: Directly call CommitRename logic
                if (titletb.IsVisible)
                {
                    CommitRename?.Invoke();
                    return;
                }

                if (e.ClickCount == 2)
                {
                    // Roll-up/roll-down logic (swapped from Ctrl+Click)
                    NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(titlelabel);
                    string fenceId = win?.Tag?.ToString();
                    // DebugLog("IMMEDIATE", fenceId ?? "UNKNOWN", $"Ctrl+Click FIRST LINE - win.Height:{win?.Height ?? -1:F1}");
                    if (string.IsNullOrEmpty(fenceId) || win == null)
                    {
                        // DebugLog("ERROR", fenceId ?? "UNKNOWN", "Missing window or fenceId in Ctrl+Click");
                        return;
                    }
                    // DebugLog("EVENT", fenceId, "Ctrl+Click TRIGGERED");
                    // DebugLogFenceState("CTRL_CLICK_START", fenceId, win);
                    dynamic currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    if (currentFence == null)
                    {
                        // DebugLog("ERROR", fenceId, "Fence not found in FenceDataManager.FenceData for Ctrl+Click");
                        return;
                    }
                    bool isRolled = currentFence.IsRolled?.ToString().ToLower() == "true";
                    // Get fence data height (always accurate from SizeChanged handler)
                    double fenceHeight = Convert.ToDouble(currentFence.Height?.ToString() ?? "130");
                    double windowHeight = win.Height;
                    // DebugLog("SYNC_CHECK", fenceId, $"Before sync - FenceHeight:{fenceHeight:F1} | WindowHeight:{windowHeight:F1} | IsRolled:{isRolled}");
                    if (!isRolled)
                    {
                        // ROLLUP: Use fence data height (always current)
                        // DebugLog("ACTION", fenceId, "Starting ROLLUP");
                        _fencesInTransition.Add(fenceId);
                        // DebugLog("TRANSITION", fenceId, "Added to transition state");
                        // FINAL FIX: Use fence data height which is always accurate from SizeChanged handler
                        double currentHeight = fenceHeight;
                        // DebugLog("ROLLUP_HEIGHT_SOURCE", fenceId, $"Using fence.Height:{fenceHeight:F1} (win.Height was stale:{win.Height:F1})");
                        // Save current fence height as UnrolledHeight
                        IDictionary<string, object> fenceDict = currentFence as IDictionary<string, object> ?? ((JObject)currentFence).ToObject<IDictionary<string, object>>();
                        fenceDict["UnrolledHeight"] = currentHeight;
                        fenceDict["IsRolled"] = "true";
                        int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                        if (fenceIndex >= 0)
                        {
                            FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                        }
                        FenceDataManager.SaveFenceData();
                        // DebugLog("SAVE", fenceId, $"Saved ROLLUP state | UnrolledHeight:{currentHeight:F1} | IsRolled:true");
                        // Roll up animation - starts from current height
                        double targetHeight = 28;   //rolled height
                        var heightAnimation = new DoubleAnimation(currentHeight, targetHeight, TimeSpan.FromSeconds(0.3))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };
                        heightAnimation.Completed += (animSender, animArgs) =>
                        {
                            // DebugLog("ANIMATION", fenceId, "ROLLUP animation completed");
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
                                            // DebugLog("UI", fenceId, "Set WrapPanel visibility to Collapsed");
                                        }
                                    }
                                }
                            }
                            _fencesInTransition.Remove(fenceId);
                            // DebugLog("TRANSITION", fenceId, "Removed from transition state");
                            // DebugLogFenceState("ROLLUP_COMPLETE", fenceId, win);
                        };
                        win.BeginAnimation(Window.HeightProperty, heightAnimation);
                        // DebugLog("ANIMATION", fenceId, $"Started ROLLUP animation from {currentHeight:F1} to height {targetHeight:F1}");
                    }
                    else
                    {
                        // ROLLDOWN: Roll down to UnrolledHeight
                        double unrolledHeight = Convert.ToDouble(currentFence.UnrolledHeight?.ToString() ?? "130");
                        // // DebugLog("ACTION", fenceId, $"Starting ROLLDOWN to {unrolledHeight:F1}");
                        _fencesInTransition.Add(fenceId);
                        // DebugLog("TRANSITION", fenceId, "Added to transition state");
                        IDictionary<string, object> fenceDict = currentFence as IDictionary<string, object> ?? ((JObject)currentFence).ToObject<IDictionary<string, object>>();
                        fenceDict["IsRolled"] = "false";
                        int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                        if (fenceIndex >= 0)
                        {
                            FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                        }
                        FenceDataManager.SaveFenceData();
                        // DebugLog("SAVE", fenceId, $"Saved ROLLDOWN state | IsRolled:false | TargetHeight:{unrolledHeight:F1}");
                        var heightAnimation = new DoubleAnimation(win.Height, unrolledHeight, TimeSpan.FromSeconds(0.3))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };
                        heightAnimation.Completed += (animSender, animArgs) =>
                        {
                            // DebugLog("ANIMATION", fenceId, "ROLLDOWN animation completed");
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
                                            // DebugLog("UI", fenceId, "Set WrapPanel visibility to Visible");
                                        }
                                    }
                                }
                            }
                            _fencesInTransition.Remove(fenceId);
                            // DebugLog("TRANSITION", fenceId, "Removed from transition state");
                            // DebugLogFenceState("ROLLDOWN_COMPLETE", fenceId, win);
                        };
                        win.BeginAnimation(Window.HeightProperty, heightAnimation);
                        // DebugLog("ANIMATION", fenceId, $"Started ROLLDOWN animation to height {unrolledHeight:F1}");
                    }
                    e.Handled = true;
                }
                else if (e.LeftButton == MouseButtonState.Pressed)
                {
                    if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
                    {
                        // Rename fence (swapped from double-click)
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Entering edit mode for fence: {fence.Title}");
                        titletb.Text = titlelabel.Content.ToString();
                        titlelabel.Visibility = Visibility.Collapsed;
                        titletb.Visibility = Visibility.Visible;
                        win.ShowActivated = true;
                        win.Activate();
                        Keyboard.Focus(titletb);
                        titletb.SelectAll();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Focus set to title textbox for fence: {fence.Title}");
                        e.Handled = true;
                    }
                    else
                    {
                        string fenceId = win.Tag?.ToString();
                        if (string.IsNullOrEmpty(fenceId))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Fence Id is missing for window '{win.Title}' during MouseDown");
                            return;
                        }
                        dynamic currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                        if (currentFence == null)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Fence with Id '{fenceId}' not found in FenceDataManager.FenceData during MouseDown");
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
                }
            };
     


            titletb.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    string originalTitle = fence.Title.ToString();
                    string newTitle = titletb.Text;

                    // Process through InterCore for special triggers
                    string finalTitle = InterCore.ProcessTitleChange(fence, newTitle, originalTitle);

                    // --- FIX START: Update LIVE Data ---
                    // Get the ID to find the fresh object in the global list
                    string id = fence.Id?.ToString();
                    var liveFence = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == id);

                    if (liveFence != null)
                    {
                        // Update the live object in the list
                        // Handle both JObject (JSON) and ExpandoObject (New Fence)
                        if (liveFence is Newtonsoft.Json.Linq.JObject jFence)
                            jFence["Title"] = finalTitle;
                        else
                            liveFence.Title = finalTitle;

                        // Also update the local reference just in case
                        fence.Title = finalTitle;
                    }
                    // --- FIX END ---

                    // Update UI
                    titlelabel.Content = finalTitle;
                    win.Title = finalTitle;
                    titletb.Visibility = Visibility.Collapsed;
                    titlelabel.Visibility = Visibility.Visible;

                    // Save the global list which now contains the updated title
                    FenceDataManager.SaveFenceData();

                    win.ShowActivated = false;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Exited edit mode via Enter, final title for fence: {finalTitle}");
                    win.Focus();
                }
                else if (e.Key == Key.Escape)
                {
                    // FIX: ESCAPE Logic (Cancel)
                    titletb.Text = fence.Title.ToString(); // Revert text
                    titletb.Visibility = Visibility.Collapsed;
                    titlelabel.Visibility = Visibility.Visible;

                    Keyboard.ClearFocus(); // Drop focus
                    win.ShowActivated = false;
                    e.Handled = true;

                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Rename cancelled via Escape");
                }
            };

            // 2. Handle Focus Loss (Auto-Save)
            titletb.LostFocus += (sender, e) =>
            {
                // Don't save if we are cancelling (Escape key handles UI)
                if (titletb.Visibility != Visibility.Visible) return;

                string originalTitle = fence.Title.ToString();
                string newTitle = titletb.Text;
                string finalTitle = InterCore.ProcessTitleChange(fence, newTitle, originalTitle);

                string id = fence.Id?.ToString();
                var liveFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == id);

                if (liveFence != null)
                {
                    IDictionary<string, object> fenceDict = liveFence as IDictionary<string, object> ??
                        ((JObject)liveFence).ToObject<IDictionary<string, object>>();
                    fenceDict["Title"] = finalTitle;

                    int index = FenceDataManager.FenceData.IndexOf(liveFence);
                    if (index >= 0) FenceDataManager.FenceData[index] = JObject.FromObject(fenceDict);

                    fence.Title = finalTitle;
                }

                titlelabel.Content = finalTitle;
                win.Title = finalTitle;
                titletb.Visibility = Visibility.Collapsed;
                titlelabel.Visibility = Visibility.Visible;

                FenceDataManager.SaveFenceData();

                win.ShowActivated = false;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Rename saved via LostFocus: {finalTitle}");
            };


            // WrapPanel wpcont = new WrapPanel();

            // --- FIX START ---
            WrapPanel wpcont = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                ItemWidth = double.NaN,
                ItemHeight = double.NaN,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0)
            };

            // CRITICAL: Tag the panel with the Fence ID so AddIcon can find settings
            // even before the panel is attached to the window.
            wpcont.Tag = fence.Id?.ToString();
            // --- FIX END ---

            ScrollViewer wpcontscr = new ScrollViewer
            {
                Content = wpcont,
                VerticalScrollBarVisibility = SettingsManager.DisableFenceScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto,
                // HorizontalScrollBarVisibility = SettingsManager.DisableFenceScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto
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
                        // Opacity = 0.2,
                        Opacity = opacity,
                        Stretch = Stretch.UniformToFill
                    };
                }
            }
            dp.Children.Add(wpcontscr);

            void InitContent()
            {
                // 1. Handle Note fences - they don't use WrapPanel
                if (fence.ItemsType?.ToString() == "Note")
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Creating Note fence content for '{fence.Title}'");
                    dp.Children.Remove(wpcontscr); // Remove ScrollViewer
                    TextBox noteTextBox = NoteFenceManager.CreateNoteContent(fence, dp);

                    bool isNoteRolled = fence.IsRolled?.ToString().ToLower() == "true";
                    noteTextBox.Visibility = isNoteRolled ? Visibility.Collapsed : Visibility.Visible;
                    return;
                }

                // 2. Handle Data/Portal Fences
                bool isRolled = fence.IsRolled?.ToString().ToLower() == "true";
                wpcont.Visibility = isRolled ? Visibility.Collapsed : Visibility.Visible;
                wpcont.Children.Clear();

                // 2a. Data Fences (The complex logic)
                if (fence.ItemsType?.ToString() == "Data")
                {
                    JArray items = null;
                    bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

                    // Load Items (Tabs vs Main)
                    if (tabsEnabled)
                    {
                        try
                        {
                            var tabs = fence.Tabs as JArray ?? new JArray();
                            int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");
                            if (currentTab >= 0 && currentTab < tabs.Count)
                            {
                                var activeTab = tabs[currentTab] as JObject;
                                items = activeTab?["Items"] as JArray ?? new JArray();
                            }
                            else
                            {
                                items = fence.Items as JArray ?? new JArray();
                            }
                        }
                        catch { items = fence.Items as JArray ?? new JArray(); }
                    }
                    else
                    {
                        items = fence.Items as JArray ?? new JArray();
                    }

                    if (items != null)
                    {
                        // Sort
                        var sortedItems = items.OfType<JObject>()
                            .OrderBy(item => item["DisplayOrder"]?.Type == JTokenType.Integer ? item["DisplayOrder"].Value<int>() : 0)
                            .ToList();

                        foreach (dynamic icon in sortedItems)
                        {
                            // FIX: Pass 'fence' context for customization
                            AddIcon(icon, wpcont, fence);

                            StackPanel sp = wpcont.Children[wpcont.Children.Count - 1] as StackPanel;
                            if (sp != null)
                            {
                                // --- 1. DEFINE VARIABLES (Scope Fix) ---
                                // We extract these immediately so they are available for ALL blocks below
                                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict
                                    ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();

                                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";

                                // CRITICAL: Define these here to avoid CS0103 errors
                                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
                                bool isLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];

                                // --- 2. Extract Arguments ---
                                string arguments = null;
                                if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk")
                                {
                                    try
                                    {
                                        WshShell shell = new WshShell();
                                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                                        arguments = shortcut.Arguments;
                                    }
                                    catch { }
                                }

                                // --- 3. Attach Click Event ---
                                ClickEventAdder(sp, filePath, isFolder, arguments);

                                // --- 4. Target Checker Logic (Robust & Safe) ---
                                // Fail-Safe: Check network status live if flag is missing
                                bool isNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"];
                                if (!isNetwork && !string.IsNullOrEmpty(filePath))
                                {
                                    isNetwork = IsNetworkPath(filePath);
                                }

                                bool allowNetworkChecking = _options.CheckNetworkPaths ?? false;

                                // Logic: Add to checker ONLY IF:
                                // A. It is NOT a web link (using 'isLink')
                                // B. AND (It is NOT a network path OR the user explicitly allowed checks)
                                if (!isLink && (!isNetwork || allowNetworkChecking))
                                {
                                    // Safe to add to background checker (using 'isFolder')
                                    targetChecker.AddCheckAction(filePath, () => UpdateIcon(sp, filePath, isFolder), isFolder);

                                    if (isNetwork && allowNetworkChecking)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"User enabled checks: Monitoring network path {filePath}");
                                    }
                                }
                                else
                                {
                                    // Logging for exclusion
                                    if (isLink)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Excluded web link {filePath} from target checking");
                                    }

                                    if (isNetwork && !allowNetworkChecking)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Safety Exclusion: Skipped background check for network path {filePath}");
                                    }
                                }

                                // --- 5. Attach Context Menu ---
                                AttachIconContextMenu(sp, icon, fence, win);
                            }
                        }
                    }
                }
                // 2b. Portal Fences
                else if (fence.ItemsType?.ToString() == "Portal")
                {
                    try
                    {
                        _portalFences[fence] = new PortalFenceManager(fence, wpcont);
                    }
                    catch (Exception ex)
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to initialize Portal Fence: {ex.Message}", "Error");
                        FenceDataManager.FenceData.Remove(fence);
                        FenceDataManager.SaveFenceData();
                        win.Close();
                    }
                }
            }
       

            win.Drop += (sender, e) =>
            {
                e.Handled = true;
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
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
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                            $"Processing dropped item: '{droppedFile}'");
                        try
                        {
                            // Path Validation (Existence / Folder Check)
                            bool fileExists = false;
                            bool directoryExists = false;
                            bool isFolderFlag = false;

                            try
                            {
                                FileAttributes attrs = System.IO.File.GetAttributes(droppedFile);
                                isFolderFlag = attrs.HasFlag(FileAttributes.Directory);
                                directoryExists = isFolderFlag;
                                fileExists = !isFolderFlag;
                            }
                            catch (Exception pathEx)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation, $"Path validation error: {pathEx.Message}");
                                continue;
                            }

                            if (!fileExists && !directoryExists)
                            {
                                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Invalid file or directory: {droppedFile}", "Error");
                                continue;
                            }

                            // --- DATA FENCE LOGIC ---


                            // --- DATA FENCE LOGIC ---
                            if (fence.ItemsType?.ToString() == "Data")
                            {
                                if (!System.IO.Directory.Exists("Shortcuts")) System.IO.Directory.CreateDirectory("Shortcuts");
                                string baseShortcutName = System.IO.Path.Combine("Shortcuts", System.IO.Path.GetFileName(droppedFile));
                                string shortcutName = baseShortcutName;
                                int counter = 1;

                                bool isDroppedShortcut = System.IO.Path.GetExtension(droppedFile).ToLower() == ".lnk";
                                bool isDroppedUrlFile = System.IO.Path.GetExtension(droppedFile).ToLower() == ".url";

                                // FIX: Trust the extension. If it ends in .url, it IS a link.
                                // This bypasses content checks that fail on files with custom headers.
                                bool isWebLink = isDroppedUrlFile || CoreUtilities.IsWebLinkShortcut(droppedFile);

                                string targetPath;
                                bool isFolder = false;
                                string webUrl = null;

                                if (isWebLink)
                                {
                                    // Try to extract clean URL, but if it fails (weird header), fallback to file path
                                    try { webUrl = CoreUtilities.ExtractWebUrlFromFile(droppedFile); } catch { }

                                    targetPath = !string.IsNullOrEmpty(webUrl) ? webUrl : droppedFile;
                                    isFolder = false;
                                }
                                else
                                {
                                    if (isDroppedShortcut)
                                    {
                                        targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(droppedFile);
                                        if (string.IsNullOrEmpty(targetPath))
                                        {
                                            isFolder = false;
                                        }
                                        else
                                        {
                                            isFolder = System.IO.Directory.Exists(targetPath);
                                            if (!isFolder && string.IsNullOrEmpty(System.IO.Path.GetExtension(targetPath)))
                                            {
                                                isFolder = System.IO.Directory.Exists(targetPath);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        targetPath = droppedFile;
                                        isFolder = System.IO.Directory.Exists(targetPath);
                                    }
                                }

                                if (!isDroppedShortcut && !isDroppedUrlFile)
                                {
                                    // CASE A: Creating new shortcut from raw file/folder
                                    shortcutName = baseShortcutName + ".lnk";
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{System.IO.Path.GetFileNameWithoutExtension(droppedFile)} ({counter++}).lnk");
                                    }

                                    try
                                    {
                                        WshShell shell = new WshShell();
                                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutName);
                                        shortcut.TargetPath = droppedFile;
                                        if (isFolder) shortcut.WorkingDirectory = droppedFile;
                                        shortcut.Save();
                                    }
                                    catch { continue; }
                                }
                                else
                                {
                                    // CASE B: Copying existing shortcut (LNK or URL)

                                    // FIX: Determine correct extension based on type
                                    // If it's a Web Link, MUST remain .url. If it's a Shortcut, MUST remain .lnk.
                                    string ext = isWebLink ? ".url" : ".lnk";

                                    // Ensure base name has correct extension
                                    if (!baseShortcutName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                                        baseShortcutName = System.IO.Path.ChangeExtension(baseShortcutName, ext);

                                    shortcutName = baseShortcutName;

                                    // Handle Duplicates while preserving extension
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        string nameNoExt = System.IO.Path.GetFileNameWithoutExtension(droppedFile);
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{nameNoExt} ({counter++}){ext}");
                                    }

                                    if (isWebLink)
                                    {
                                        // If we got a clean URL, create a fresh file (strips bad headers)
                                        if (!string.IsNullOrEmpty(webUrl))
                                            CreateWebLinkShortcut(webUrl, shortcutName, System.IO.Path.GetFileNameWithoutExtension(droppedFile));
                                        else
                                            // Fallback: Copy original file exactly as-is (preserves functionality even if we can't parse it)
                                            System.IO.File.Copy(droppedFile, shortcutName, true);
                                    }
                                    else
                                    {
                                        System.IO.File.Copy(droppedFile, shortcutName, false);
                                    }
                                }

                                dynamic newItem = new System.Dynamic.ExpandoObject();
                                IDictionary<string, object> newItemDict = newItem;

                                newItemDict["Filename"] = shortcutName;
                                newItemDict["IsFolder"] = isFolder;
                                newItemDict["IsLink"] = isWebLink;
                                newItemDict["IsNetwork"] = IsNetworkPath(shortcutName);

                                string displayFileName = isFolder
                                    ? System.IO.Path.GetFileName(droppedFile)
                                    : System.IO.Path.GetFileNameWithoutExtension(droppedFile);
                                newItemDict["DisplayName"] = displayFileName;
                                newItemDict["AlwaysRunAsAdmin"] = false;

                                // TABS FEATURE: Get fresh fence data
                                string fenceId = fence.Id?.ToString();
                                var freshFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                                JArray items = null;
                                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                                bool addedToTab0 = false;

                                if (tabsEnabled && freshFence != null)
                                {
                                    int currentTab = Convert.ToInt32(freshFence.CurrentTab?.ToString() ?? "0");
                                    var tabs = freshFence.Tabs as JArray ?? new JArray();
                                    if (currentTab >= 0 && currentTab < tabs.Count)
                                    {
                                        var activeTab = tabs[currentTab] as JObject;
                                        if (activeTab != null)
                                        {
                                            items = activeTab["Items"] as JArray ?? new JArray();
                                            addedToTab0 = (currentTab == 0);
                                        }
                                    }
                                    if (items == null) items = freshFence.Items as JArray ?? new JArray();
                                }
                                else
                                {
                                    items = (freshFence ?? fence).Items as JArray ?? new JArray();
                                }

                                int nextDisplayOrder = items.Count;
                                newItemDict["DisplayOrder"] = nextDisplayOrder;
                                items.Add(JObject.FromObject(newItem));

                                if (!string.IsNullOrEmpty(fenceId))
                                {
                                    if (addedToTab0) SynchronizeTab0Content(fenceId, "tab0", "add");
                                    else if (!tabsEnabled) SynchronizeTab0Content(fenceId, "main", "add");
                                }

                                if (freshFence != null)
                                {
                                    int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                                    if (fenceIndex >= 0)
                                    {
                                        FenceDataManager.FenceData[fenceIndex] = freshFence;
                                        FenceDataManager.SaveFenceData();
                                    }
                                }

                                AddIcon(newItem, wpcont);
                                StackPanel sp = wpcont.Children[wpcont.Children.Count - 1] as StackPanel;
                                if (sp != null)
                                {
                                    // FIX: Extract arguments for the newly created shortcut
                                    string args = Utility.GetShortcutArguments(shortcutName);

                                    ClickEventAdder(sp, shortcutName, isFolder, args);

                                    // Target Checker Logic
                                    bool isRealNetwork = IsNetworkPath(shortcutName);
                                    bool allowNetworkChecking = _options.CheckNetworkPaths ?? false;

                                    if (!isRealNetwork || allowNetworkChecking)
                                    {
                                        targetChecker.AddCheckAction(shortcutName, () => UpdateIcon(sp, shortcutName, isFolder), isFolder);
                                        if (isRealNetwork) LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Monitoring network path: {shortcutName}");
                                    }
                                    else
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Skipped background check for: {shortcutName}");
                                    }

                                    // Attach Context Menu
                                    AttachIconContextMenu(sp, newItem, fence, win);

                                    // --- HIDDEN TWEAK: Delete Original Shortcut On Drop ---
                                    if (SettingsManager.DeleteOriginalShortcutsOnDrop && isDroppedShortcut && !isFolder && !isWebLink)
                                    {
                                        try
                                        {
                                            string fileDir = System.IO.Path.GetDirectoryName(droppedFile);
                                            string userDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                                            string commonDesktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);

                                            bool isFromDesktop = string.Equals(fileDir, userDesktop, StringComparison.OrdinalIgnoreCase) ||
                                                                 string.Equals(fileDir, commonDesktop, StringComparison.OrdinalIgnoreCase);

                                            if (isFromDesktop) System.IO.File.Delete(droppedFile);
                                        }
                                        catch { }
                                    }
                                }
                            }






                            // --- PORTAL FENCE LOGIC (RESTORED) ---
                            else if (fence.ItemsType?.ToString() == "Portal")
                            {
                                IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
                                string destinationFolder = fenceDict.ContainsKey("Path") ? fenceDict["Path"]?.ToString() : null;

                                if (string.IsNullOrEmpty(destinationFolder))
                                {
                                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"No destination folder defined for this Portal Fence.", "Error");
                                    continue;
                                }
                                if (!System.IO.Directory.Exists(destinationFolder))
                                {
                                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"The destination folder '{destinationFolder}' no longer exists.", "Error");
                                    continue;
                                }

                                string destinationPath = System.IO.Path.Combine(destinationFolder, System.IO.Path.GetFileName(droppedFile));
                                int counter = 1;
                                string baseName = System.IO.Path.GetFileNameWithoutExtension(droppedFile);
                                string extension = System.IO.Path.GetExtension(droppedFile);

                                while (System.IO.File.Exists(destinationPath) || System.IO.Directory.Exists(destinationPath))
                                {
                                    destinationPath = System.IO.Path.Combine(destinationFolder, $"{baseName} ({counter++}){extension}");
                                }

                                if (System.IO.File.Exists(droppedFile))
                                {
                                    System.IO.File.Copy(droppedFile, destinationPath, false);
                                }
                                else if (System.IO.Directory.Exists(droppedFile))
                                {
                                    BackupManager.CopyDirectory(droppedFile, destinationPath);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to add {droppedFile}: {ex.Message}", "Error");
                        }
                    }
                    FenceDataManager.SaveFenceData();
                }
                // --- URL DROP LOGIC (RESTORED) ---
                else if (e.Data.GetDataPresent(DataFormats.Text) ||
                         e.Data.GetDataPresent(DataFormats.Html) ||
                         e.Data.GetDataPresent("UniformResourceLocator"))
                {
                    try
                    {
                        string droppedUrl = ExtractUrlFromDropData(e.Data);
                        if (!string.IsNullOrEmpty(droppedUrl) && IsValidWebUrl(droppedUrl))
                        {
                            if (fence.ItemsType?.ToString() == "Data")
                            {
                                AddUrlShortcutToFence(droppedUrl, fence, wpcont);
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation, "URL drops not supported for Portal fences");
                            }
                        }
                    }
                    catch (Exception urlEx)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation, $"Error processing URL drop: {urlEx.Message}");
                    }
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
                // Find the current fence in FenceDataManager.FenceData using ID
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence != null)
                {
                    // Update position and save immediately
                    currentFence.X = win.Left;
                    currentFence.Y = win.Top;
                    FenceDataManager.SaveFenceData();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate, $"Position updated for fence '{currentFence.Title}' to X={win.Left}, Y={win.Top}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate, $"Fence with Id '{fenceId}' not found during position change");
                }
            };
            InitContent();
            // Add Note fence specific context menu items after content is initialized
            if (fence.ItemsType?.ToString() == "Note")
            {
                // Use a small delay to ensure the TextBox is fully created and added to the visual tree
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        // Find the TextBox that was created in InitContent()
                        var border = win.Content as Border;
                        var dockPanel = border?.Child as DockPanel;
                        var noteTextBox = dockPanel?.Children.OfType<TextBox>().FirstOrDefault();
                        if (noteTextBox != null)
                        {
                            NoteFenceManager.AddNoteContextMenuItems(CnMnFenceManager, fence, noteTextBox);
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                $"Added Note context menu items for fence '{fence.Title}'");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                $"CRITICAL: Could not find TextBox for Note fence '{fence.Title}' - checking DockPanel children");
                            // Debug: Log what children actually exist
                            if (dockPanel != null)
                            {
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                    $"DockPanel has {dockPanel.Children.Count} children:");
                                for (int i = 0; i < dockPanel.Children.Count; i++)
                                {
                                    var child = dockPanel.Children[i];
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                        $" Child {i}: {child.GetType().Name}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                            $"Error adding Note context menu: {ex.Message}");
                    }
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
            win.Show();







            // Check for persistent Legendary Mode (Nikos or >:)
            string fTitle = fence.Title?.ToString() ?? "";
            if (fTitle == "Nikos" || fTitle == "Nikos Georgousis" || fTitle.Contains(">:"))
            {
                // Defer slightly to ensure window is loaded
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    // We call ProcessTitleChange to trigger the visual effect
                    InterCore.ProcessTitleChange(fence, fTitle, "");
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
            {
                // Defer slightly to ensure window is loaded
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    // We call ProcessTitleChange with same names just to trigger the logic
                    InterCore.ProcessTitleChange(fence, fence.Title.ToString(), "");
                }), System.Windows.Threading.DispatcherPriority.Background);
            }







            IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
            SnapManager.AddSnapping(win, fenceDict);
            // Apply custom color if present, otherwise use global
            string customColor = fence.CustomColor?.ToString();
            Utility.ApplyTintAndColorToFence(win, string.IsNullOrEmpty(customColor) ? SettingsManager.SelectedColor : customColor);
            targetChecker.Start();
        }

        public static void AddIcon(dynamic icon, WrapPanel wpcont, dynamic fenceContext = null)
        {
            // 1. EXTRACT DATA
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict
                ? dict
                : ((JObject)icon).ToObject<IDictionary<string, object>>();

            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
            bool isLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];

            // Fail-Safe Network Detection
            bool isNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"];
            if (!isNetwork && !string.IsNullOrEmpty(filePath))
            {
                isNetwork = IsNetworkPath(filePath);
            }

            // --- STEP 1: Determine Settings Context ---
            dynamic settings = fenceContext;
            if (settings == null)
            {
                try
                {
                    if (wpcont.Tag != null)
                    {
                        string fenceId = wpcont.Tag.ToString();
                        settings = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    }
                    if (settings == null)
                    {
                        NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(wpcont);
                        string fenceId = win?.Tag?.ToString();
                        if (!string.IsNullOrEmpty(fenceId))
                        {
                            settings = GetFenceData().FirstOrDefault(f => f.Id?.ToString() == fenceId);
                        }
                    }
                }
                catch { }
            }

            // --- NEW: Identify Fence Type (Portal vs Data) ---
            bool isPortal = false;
            if (settings != null)
            {
                try
                {
                    // Handle both JObject (JSON) and ExpandoObject (Runtime)
                    if (settings is JObject jObj)
                        isPortal = jObj["ItemsType"]?.ToString() == "Portal";
                    else
                        isPortal = settings.ItemsType?.ToString() == "Portal";
                }
                catch { }
            }

            // --- STEP 2: Read Customization Settings ---
            int iconSpacing = 5;
            double iconWidth = 40, iconHeight = 40;
            bool grayscale = false;
            string textColorName = null;
            bool disableShadow = false;

            if (settings != null)
            {
                try { iconSpacing = Convert.ToInt32(settings.IconSpacing?.ToString() ?? "5"); } catch { }

                string sizeVal = settings.IconSize?.ToString() ?? "Medium";
                switch (sizeVal)
                {
                    case "Tiny": iconWidth = iconHeight = 24; break;
                    case "Small": iconWidth = iconHeight = 32; break;
                    case "Large": iconWidth = iconHeight = 48; break;
                    case "Huge": iconWidth = iconHeight = 64; break;
                    default: iconWidth = iconHeight = 40; break;
                }

                try { grayscale = settings.GrayscaleIcons?.ToString().ToLower() == "true"; } catch { }
                try { textColorName = settings.TextColor?.ToString(); } catch { }
                try { disableShadow = settings.DisableTextShadow?.ToString().ToLower() == "true"; } catch { }
            }

            // --- STEP 3: Create UI Elements ---
            StackPanel sp = new StackPanel
            {
                Margin = new Thickness(iconSpacing),
                Width = 60 + (iconSpacing * 2)
            };

            System.Windows.Controls.Image ico = new System.Windows.Controls.Image
            {
                Width = iconWidth,
                Height = iconHeight,
                Margin = new Thickness(5)
            };

            if (SettingsManager.IconVisibilityEffect != IconVisibilityEffect.None)
            {
                ico.Effect = Utility.CreateIconEffect(SettingsManager.IconVisibilityEffect);
            }





            // --- ICON EXTRACTION LOGIC ---
            ImageSource shortcutIcon = null;
            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
            bool isUrlFile = System.IO.Path.GetExtension(filePath).ToLower() == ".url";

            // Variables to track target state
            string targetPath = null;
            bool targetIsUncRoot = false;

            // FIX: Unified Custom Icon Extraction for Links AND Shortcuts
            // We check for a custom icon FIRST. If one exists, we use it immediately.
            if (isShortcut || isLink || isUrlFile)
            {
                // NEW: Special handling for .url files using manual parser
                if (isUrlFile || isLink)
                {
                    var urlIcon = GetUrlCustomIcon(filePath);
                    if (urlIcon.Path != null)
                    {
                        shortcutIcon = IconManager.ExtractIconFromFile(urlIcon.Path, urlIcon.Index);
                    }
                }

                // Existing .lnk logic
                if (shortcutIcon == null)
                {


                    try
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(System.IO.Path.GetFullPath(filePath));
                        targetPath = shortcut.TargetPath;

                        // 1. Try Custom Icon (Properties -> Change Icon)
                        // Check if IconLocation is valid and NOT ",0" (which implies default)
                        if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
                        {
                            string[] iconParts = shortcut.IconLocation.Split(',');
                            string iconPath = iconParts[0];
                            int iconIndex = 0;
                            if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex))
                                iconIndex = parsedIndex;

                            if (System.IO.File.Exists(iconPath))
                            {
                                shortcutIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                            }
                        }
                    }


                    catch { }
                }
                }

            // 2. Fallback: Determine Icon if no custom icon found
            if (shortcutIcon == null)
            {
                if (isLink || isUrlFile)
                {
                    // Default Link Icon
                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));
                }
                else if (isShortcut)
                {
                    try
                    {
                        // FIX: Use the robust Unicode-Safe reader (IShellLink)
                        // This handles tricky paths better than WshShell
                        if (string.IsNullOrEmpty(targetPath))
                            targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);

                        // Fallback to WshShell only if utility failed and we haven't tried yet
                        if (string.IsNullOrEmpty(targetPath))
                        {
                            try
                            {
                                WshShell shell = new WshShell();
                                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(System.IO.Path.GetFullPath(filePath));
                                targetPath = shortcut.TargetPath;
                            }
                            catch { }
                        }

                        // FIX 1: RUNTIME CORRECTION (Server Root)
                        // If target is \\192.168.1.10, force isFolder=true immediately.
                        if (IsUncRoot(targetPath))
                        {
                            targetIsUncRoot = true;
                            isFolder = true;   // Force folder treatment!
                            isNetwork = true;  // Force network flag
                        }
                    }
                    catch { }

                    // A. Standard Checks
                    bool targetIsFolder = false;
                    bool targetExists = false;

                    if (!string.IsNullOrEmpty(targetPath))
                    {
                        targetIsFolder = System.IO.Directory.Exists(targetPath) || targetIsUncRoot;
                        targetExists = targetIsFolder || System.IO.File.Exists(targetPath);
                    }

                    // B. SCRAPING FALLBACK
                    // If standard checks failed, scan the file content for \\Server pattern
                    if (!targetExists)
                    {
                        if (ScrapeLnkForNetworkRoot(filePath))
                        {
                            targetIsFolder = true;
                            isNetwork = true;      // Network Symbol
                            isFolder = true;       // Folder Shape
                            targetIsUncRoot = true; // CRITICAL: Force Validity (No White X)
                        }
                    }

                    // C. Icon Selection Logic
                    if (!isPortal && (isFolder || targetIsFolder))
                    {
                        shortcutIcon = null; // Force fall-through to White Theme
                    }
                    else if (targetExists)
                    {
                        shortcutIcon = Utility.GetShellIcon(targetPath, targetIsFolder);
                    }
                    else
                    {
                        shortcutIcon = Utility.GetShellIcon(filePath, isFolder);
                    }
                }
                else
                {
                    // Standard File/Folder (Not a shortcut/link)
                    if (!isPortal && isFolder)
                    {
                        shortcutIcon = null;
                    }
                    else
                    {
                        shortcutIcon = Utility.GetShellIcon(filePath, isFolder);
                    }
                }
            }

            // Final Fallback (The White Theme Logic)
            if (shortcutIcon == null)
            {
                if (isFolder)
                {
                    // FIX 2: VALIDATION OVERRIDE
                    // If it is a UNC Root (either detected via shortcut target or direct path), consider it "Valid"
                    bool isUncRoot = targetIsUncRoot || IsUncRoot(filePath);
                    bool valid = isUncRoot || FilePathUtilities.DoesFolderExist(filePath, isFolder);

                    // Double check target if shortcut
                    if (!valid && isShortcut && !string.IsNullOrEmpty(targetPath))
                    {
                        valid = IsUncRoot(targetPath) || System.IO.Directory.Exists(targetPath);
                    }

                    if (valid)
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    else
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                }
                else
                {
                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                }
            }

            ico.Source = shortcutIcon;




            // Apply Grayscale
            if (grayscale && shortcutIcon is BitmapSource bmp)
            {
                ico.Source = new FormatConvertedBitmap(bmp, PixelFormats.Gray8, BitmapPalettes.Gray256, 0);
            }
            else if (grayscale)
            {
                ico.Opacity = 0.6;
                ico.Effect = new DropShadowEffect { Color = Colors.Gray, BlurRadius = 0, ShadowDepth = 0 };
            }

            // --- STEP 4: Network Overlay ---
            if (isNetwork)
            {
                Grid iconGrid = new Grid { Width = iconWidth + 8, Height = iconHeight + 8 };
                iconGrid.Children.Add(ico);

                TextBlock networkIndicator = new TextBlock
                {
                    Text = "🔗",
                    FontSize = 14,
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(65, 135, 225)),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(2, 2, 0, 0),
                    Effect = new DropShadowEffect { Color = Colors.Black, BlurRadius = 2, Opacity = 0.7 }
                };

                iconGrid.Children.Add(networkIndicator);
                sp.Children.Add(iconGrid);
            }
            else
            {
                sp.Children.Add(ico);
            }

            // --- STEP 5: Text Label ---
            string displayName = (!iconDict.ContainsKey("DisplayName") || iconDict["DisplayName"] == null)
                ? System.IO.Path.GetFileNameWithoutExtension(filePath)
                : (string)iconDict["DisplayName"];

            if (displayName.Length > SettingsManager.MaxDisplayNameLength)
                displayName = displayName.Substring(0, SettingsManager.MaxDisplayNameLength) + "...";

            System.Windows.Media.Brush textBrush = System.Windows.Media.Brushes.White;
            if (!string.IsNullOrEmpty(textColorName))
            {
                try { textBrush = new SolidColorBrush(Utility.GetColorFromName(textColorName)); } catch { }
            }

            TextBlock lbl = new TextBlock
            {
                Text = displayName,
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Foreground = textBrush,
                MaxWidth = 70
            };

            if (!disableShadow)
            {
                lbl.Effect = new DropShadowEffect { Color = Colors.Black, Direction = 315, ShadowDepth = 2, BlurRadius = 3, Opacity = 0.8 };
            }

            sp.Children.Add(lbl);
            sp.Tag = new { FilePath = filePath, IsFolder = isFolder, Arguments = (string)(iconDict.ContainsKey("Arguments") ? iconDict["Arguments"] : null) };

            string toolTipText = $"{System.IO.Path.GetFileName(filePath)}";
            sp.ToolTip = new ToolTip { Content = toolTipText };

            wpcont.Children.Add(sp);
        }

  

        private static void CreateNewFence(string title, string itemsType, double x = 20, double y = 20, string customColor = null, string customLaunchEffect = null)
        {
            // Generate random name instead of using the passed title
            string fenceName = CoreUtilities.GenerateRandomName();

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
            // TABS FEATURE: Initialize tab properties for new fences
            newFenceDict["TabsEnabled"] = "false";  // Default to no tabs
            newFenceDict["CurrentTab"] = 0;         // Default to first tab
            newFenceDict["Tabs"] = new JArray();    // Empty tabs array

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
            FenceDataManager.FenceData.Add(newFence);
            FenceDataManager.SaveFenceData();
            CreateFence(newFence, new TargetChecker(1000));
        }


        /// <summary>
        /// HELPER: Extracts a Network Root path (e.g. \\Server) from a .lnk file by reading raw content.
        /// Used when Windows APIs fail to resolve the target.
        /// </summary>
        private static string GetScrapedNetworkPath(string lnkPath)
        {
            try
            {
                if (string.IsNullOrEmpty(lnkPath) || !System.IO.File.Exists(lnkPath)) return null;

                string fileContent;
                // Read with Share.ReadWrite to avoid locking issues
                using (var fs = new System.IO.FileStream(lnkPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                using (var sr = new System.IO.StreamReader(fs, System.Text.Encoding.Default))
                {
                    fileContent = sr.ReadToEnd();
                }

                // Regex to capture \\Servername or \\1.2.3.4
                // Matches start of line or null-preceded, followed by \\, then valid host chars, ending with null or newline
                var regex = new System.Text.RegularExpressions.Regex(@"(^|\0)\\\\([a-zA-Z0-9\.\-_]+)(\x00|$)", System.Text.RegularExpressions.RegexOptions.Multiline);
                var match = regex.Match(fileContent);

                if (match.Success)
                {
                    // Clean up the match (remove leading nulls if any) to get just "\\Server"
                    string raw = match.Value.Trim('\0', '\r', '\n');
                    if (raw.StartsWith(@"\\")) return raw;
                }

                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Scrape path failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// SAFE FALLBACK: Scans the raw binary content of a .lnk file for Network Root patterns (e.g. \\192.168.1.10).
        /// This bypasses Windows API validation failures for server roots or offline locations.
        /// </summary>
        private static bool ScrapeLnkForNetworkRoot(string lnkPath)
        {
            try
            {
                if (string.IsNullOrEmpty(lnkPath) || !System.IO.File.Exists(lnkPath)) return false;

                // 1. Read file safely with a timeout mechanism (handled by standard stream opening)
                // Use default encoding to capture ANSI/ASCII strings embedded in the binary
                string fileContent;
                using (var fs = new System.IO.FileStream(lnkPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                using (var sr = new System.IO.StreamReader(fs, System.Text.Encoding.Default))
                {
                    fileContent = sr.ReadToEnd();
                }

                // 2. Define Regex for UNC Root: Matches \\Server or \\192.168.1.1
                // Explanation:
                // \\\\       -> Literal "\\"
                // [^\\]+     -> Any character that is NOT a backslash (The server name/IP)
                // (\x00|$)   -> Ends with a null byte or end of string (typical in binary formats)
                var networkRootRegex = new System.Text.RegularExpressions.Regex(@"^\\\\([a-zA-Z0-9\.\-_]+)(\x00|$)", System.Text.RegularExpressions.RegexOptions.Multiline);

                // 3. Scan the content
                // We scan for specific "unc" prefixes often found in LNK structures or just the raw path
                if (networkRootRegex.IsMatch(fileContent)) return true;

                // Fallback simple string check for the user's specific case (\\IP) if regex is too strict on binary noise
                // We look for "\\" followed by a digit (IP) that appears in the file
                int index = fileContent.IndexOf(@"\\");
                if (index >= 0 && index + 3 < fileContent.Length)
                {
                    char nextChar = fileContent[index + 2];
                    // If it looks like \\1... or \\a... it's likely a path. 
                    // To be safe, we check if it DOESN'T look like a file path (no ":\")
                    bool hasColon = fileContent.IndexOf(@":\", index) == index + 2; // e.g. C:\ check
                    if (!hasColon) return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                // SANITIZATION: Never let a scraping error break the UI loading
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Lnk scraping suppressed error: {ex.Message}");
                return false;
            }
        }



        private static void UpdateLockState(TextBlock lockIcon, dynamic fence, bool? forceState = null, bool saveToJson = true)
        {
            // Get the actual fence from FenceDataManager.FenceData using Id to ensure correct reference
            string fenceId = fence.Id?.ToString();
            if (string.IsNullOrEmpty(fenceId))
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Fence '{fence.Title}' has no Id, cannot update lock state");
                return;
            }

            int index = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
            if (index < 0)
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate, $"Fence '{fence.Title}' not found in FenceDataManager.FenceData, cannot update lock state");
                return;
            }

            dynamic actualFence = FenceDataManager.FenceData[index];
            bool isLocked = forceState ?? (actualFence.IsLocked?.ToString().ToLower() == "true");

            // Only update JSON if explicitly requested (e.g., during toggle, not initialization)
            if (saveToJson)
            {
                UpdateFenceProperty(actualFence, "IsLocked", isLocked.ToString().ToLower(), $"Fence {(isLocked ? "locked" : "unlocked")}");
            }

            // Update UI on the main thread
            System.Windows.Application.Current.Dispatcher.Invoke(() =>


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



        // --- HELPER: Manual Icon Extraction for .url Files ---
        private static (string Path, int Index) GetUrlCustomIcon(string urlPath)
        {
            try
            {
                if (!System.IO.File.Exists(urlPath)) return (null, 0);
                var lines = System.IO.File.ReadAllLines(urlPath);
                string iconFile = null;
                int iconIndex = 0;

                foreach (var line in lines)
                {
                    string trimmed = line.Trim();
                    if (trimmed.StartsWith("IconFile=", StringComparison.OrdinalIgnoreCase))
                        iconFile = trimmed.Substring(9).Trim();
                    else if (trimmed.StartsWith("IconIndex=", StringComparison.OrdinalIgnoreCase))
                        int.TryParse(trimmed.Substring(10).Trim(), out iconIndex);
                }

                if (!string.IsNullOrEmpty(iconFile)) return (iconFile, iconIndex);
            }
            catch { }
            return (null, 0);
        }




        public static void EditItem(dynamic icon, dynamic fence, NonActivatingWindow win)
        {
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            string displayName = iconDict.ContainsKey("DisplayName") ? (string)iconDict["DisplayName"] : System.IO.Path.GetFileNameWithoutExtension(filePath);

            // Allow .lnk and .url
            string ext = System.IO.Path.GetExtension(filePath).ToLower();
            bool isEditable = ext == ".lnk" || ext == ".url";

            if (!isEditable)
            {
                MessageBoxesManager.ShowOKOnlyMessageBoxForm("Edit is not available for this item type.", "Info");
                return;
            }

            var editWindow = new EditShortcutWindow(filePath, displayName);

            // Note: The EditWindow calls 'UpdateFenceDataForIcon' internally when Save is clicked.
            // We don't need to duplicate the saving logic here.
            if (editWindow.ShowDialog() == true)
            {
                string newDisplayName = editWindow.NewDisplayName;

                // Update local memory reference immediately for responsiveness
                iconDict["DisplayName"] = newDisplayName;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"EditItem: Local display name updated to {newDisplayName}");

                // UI Refresh Logic
                if (win == null) return;

                WrapPanel wpcont = null;
                var border = win.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                var scrollViewer = dockPanel?.Children.OfType<ScrollViewer>().FirstOrDefault();
                if (scrollViewer != null) wpcont = scrollViewer.Content as WrapPanel;

                if (wpcont == null) wpcont = FenceUtilities.FindWrapPanel(win);

                if (wpcont != null)
                {
                    var sp = wpcont.Children.OfType<StackPanel>()
                        .FirstOrDefault(s => s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == filePath);

                    if (sp != null)
                    {
                        RefreshSingleIconComplete(sp, filePath, newDisplayName, win);
                    }
                    else
                    {
                        // Fallback: Rebuild if not found (e.g. Tab switch might have hidden it)
                        // In most cases RefreshSingleIconComplete handles it, or a full fence refresh will catch it later.
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, "StackPanel not found for immediate refresh, data is saved though.");
                    }
                }
            }
        }



        // --- PROFESSIONAL UX FINAL v13: The "Sorcery" Fix ---
        // 1. Reverts Padding to 'TemplateBinding' so the [+] button isn't crushed.
        // 2. Ensures the [+] button (Width 25) gets its 0 padding, while Text tabs get 10.
        private static void ApplyTabStyle(Button btn, bool isActive, string colorName, bool isPlusButton = false)
        {
            try
            {
                // 1. RESET
                btn.Style = null;
                btn.FocusVisualStyle = null;
                btn.Focusable = false;

                btn.ClearValue(Button.BackgroundProperty);
                btn.ClearValue(Button.ForegroundProperty);
                btn.ClearValue(Button.BorderBrushProperty);
                btn.ClearValue(Button.FontWeightProperty);
                btn.ClearValue(Button.PaddingProperty); // Clear local padding to be safe

                btn.MouseEnter -= Tab_MouseEnter_Lambda;
                btn.MouseLeave -= Tab_MouseLeave_Lambda;

                // 2. COLOR CALCULATION
                string effectiveColor = !string.IsNullOrEmpty(colorName) ? colorName : SettingsManager.SelectedColor;
                System.Windows.Media.Color baseColor = System.Windows.Media.Colors.Gray;
                try
                {
                    var drawingColor = Utility.GetColorFromName(effectiveColor);
                    baseColor = System.Windows.Media.Color.FromArgb(255, drawingColor.R, drawingColor.G, drawingColor.B);
                }
                catch { }

                string c = effectiveColor?.ToLower() ?? "";
                bool isExplicitDark = c.Contains("blue") || c.Contains("teal") || c.Contains("black") ||
                                      c.Contains("red") || c.Contains("green") || c.Contains("purple") ||
                                      c.Contains("bismark") || c.Contains("fuchsia") || c.Contains("default");

                double brightness = Math.Sqrt(
                    (0.299 * baseColor.R * baseColor.R) +
                    (0.587 * baseColor.G * baseColor.G) +
                    (0.114 * baseColor.B * baseColor.B)
                );

                bool isDarkTheme = isExplicitDark || brightness < 160;

                // 3. PALETTE DEFINITION
                SolidColorBrush bgActive, bgInactive, bgHover;
                SolidColorBrush textActive, textInactive;
                SolidColorBrush borderActive, borderInactive;

                if (isDarkTheme)
                {
                    bgActive = new SolidColorBrush(baseColor);
                    bgInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 255, 255, 255));
                    bgHover = new SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 255, 255, 255));

                    textActive = System.Windows.Media.Brushes.White;
                    textInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(180, 255, 255, 255));

                    borderActive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(120, 255, 255, 255));
                    borderInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 255, 255, 255));
                }
                else
                {
                    bgActive = System.Windows.Media.Brushes.White;
                    bgInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(20, 0, 0, 0));
                    bgHover = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 0, 0, 0));

                    textActive = new SolidColorBrush(System.Windows.Media.Color.FromRgb(20, 20, 20));
                    textInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(160, 0, 0, 0));

                    borderActive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 0, 0, 0));
                    borderInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 0, 0, 0));
                }

                // 4. BUTTON CONFIGURATION
                if (isPlusButton)
                {
                    // Special Case: [+] Button needs 0 padding to center the text in 25px width
                    btn.Padding = new Thickness(0);

                    if (!isActive)
                    {
                        bgInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(
                            (byte)(bgInactive.Color.A / 2),
                            bgInactive.Color.R,
                            bgInactive.Color.G,
                            bgInactive.Color.B));
                    }
                }
                else
                {
                    // Standard Tab: Needs padding for breathing room
                    btn.Padding = new Thickness(10, 2, 10, 2);
                }

                // 5. TEMPLATE GENERATION
                ControlTemplate template = new ControlTemplate(typeof(Button));
                FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
                border.Name = "Border";
                border.SetValue(Border.CornerRadiusProperty, new CornerRadius(4, 4, 0, 0));

                // CRITICAL FIX: Bind Padding to the Button's property instead of hardcoding it.
                // This allows the [+] button to have 0 padding and Text tabs to have 10 padding.
                border.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));

                FrameworkElementFactory content = new FrameworkElementFactory(typeof(ContentPresenter));
                content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
                content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

                border.AppendChild(content);
                template.VisualTree = border;

                // 6. APPLY STYLE
                Style style = new Style(typeof(Button));
                style.Setters.Add(new Setter(Button.TemplateProperty, template));

                if (isActive)
                {
                    // ACTIVE
                    style.Setters.Add(new Setter(Button.BackgroundProperty, bgActive));
                    style.Setters.Add(new Setter(Button.ForegroundProperty, textActive));
                    style.Setters.Add(new Setter(Button.BorderBrushProperty, borderActive));
                    style.Setters.Add(new Setter(Button.BorderThicknessProperty, new Thickness(1, 1, 1, 0)));
                    style.Setters.Add(new Setter(Button.FontWeightProperty, FontWeights.Bold));
                    style.Setters.Add(new Setter(Button.OpacityProperty, 1.0));

                    Trigger staticTrigger = new Trigger { Property = Button.IsEnabledProperty, Value = true };
                    staticTrigger.Setters.Add(new Setter(Border.BackgroundProperty, bgActive, "Border"));
                    staticTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, borderActive, "Border"));
                    staticTrigger.Setters.Add(new Setter(Border.BorderThicknessProperty, new Thickness(1, 1, 1, 0), "Border"));
                    template.Triggers.Add(staticTrigger);

                    if (isDarkTheme)
                        style.Setters.Add(new Setter(Button.EffectProperty, new DropShadowEffect { BlurRadius = 4, ShadowDepth = 1, Direction = 270, Color = System.Windows.Media.Colors.Black, Opacity = 0.5 }));
                }
                else
                {
                    // INACTIVE
                    style.Setters.Add(new Setter(Button.BackgroundProperty, bgInactive));
                    style.Setters.Add(new Setter(Button.ForegroundProperty, textInactive));
                    style.Setters.Add(new Setter(Button.BorderBrushProperty, borderInactive));
                    style.Setters.Add(new Setter(Button.BorderThicknessProperty, new Thickness(1, 1, 1, 1)));
                    style.Setters.Add(new Setter(Button.FontWeightProperty, FontWeights.Normal));

                    Trigger baseTrigger = new Trigger { Property = Button.IsEnabledProperty, Value = true };
                    baseTrigger.Setters.Add(new Setter(Border.BackgroundProperty, bgInactive, "Border"));
                    baseTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, borderInactive, "Border"));
                    baseTrigger.Setters.Add(new Setter(Border.BorderThicknessProperty, new Thickness(1, 1, 1, 1), "Border"));
                    template.Triggers.Add(baseTrigger);

                    Trigger hoverTrigger = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
                    hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, bgHover, "Border"));
                    hoverTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, borderActive, "Border"));
                    hoverTrigger.Setters.Add(new Setter(Button.ForegroundProperty, textActive));
                    hoverTrigger.Setters.Add(new Setter(Button.CursorProperty, Cursors.Hand));
                    template.Triggers.Add(hoverTrigger);
                }

                btn.Style = style;
            }
            catch { }
        }


        // Dummy handlers to allow -= syntax (prevents compiler errors if we were using named methods)
        // Since we use lambdas above, we don't strictly need these, but good for safety if refactoring.
        private static void Tab_MouseEnter_Lambda(object sender, System.Windows.Input.MouseEventArgs e) { }
        private static void Tab_MouseLeave_Lambda(object sender, System.Windows.Input.MouseEventArgs e) { }
    
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
        /// Refresh fence using the same approach as CustomizeFenceForm
        /// This ensures consistent sizing and behavior
        /// </summary>
        public static void RefreshFenceUsingFormApproach(NonActivatingWindow win, dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Refreshing fence using form approach for '{fence.Title}'");

                // Find the WrapPanel
                var wrapPanel = FindWrapPanel(win);
                if (wrapPanel == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot find WrapPanel for fence refresh");
                    return;
                }

                // Clear existing icons
                wrapPanel.Children.Clear();

                // TABS FEATURE: Check if tabs are enabled and load from appropriate source
                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                JArray items = null;

                if (tabsEnabled)
                {
                    // Load from current tab
                    var tabs = fence.Tabs as JArray ?? new JArray();
                    int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");

                    if (currentTab >= 0 && currentTab < tabs.Count)
                    {
                        var activeTab = tabs[currentTab] as JObject;
                        if (activeTab != null)
                        {
                            items = activeTab["Items"] as JArray ?? new JArray();
                            string tabName = activeTab["TabName"]?.ToString() ?? $"Tab {currentTab}";
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                $"Refreshing icons from tab '{tabName}' for fence '{fence.Title}'");
                        }
                    }
                }
                else
                {
                    // Load from main Items array (existing behavior)
                    items = fence.Items as JArray;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                        $"Refreshing icons from main Items for fence '{fence.Title}'");
                }

                if (items != null)
                {
                    // Sort items by DisplayOrder and add them (SAME AS FORM)
                    var sortedItems = items
                        .OfType<JObject>()
                        .OrderBy(item => item["DisplayOrder"]?.Value<int>() ?? 0)
                        .ToList();

                    foreach (dynamic item in sortedItems)
                    {
                        // Use the SAME method as the working form
                        AddIcon(item, wrapPanel);

                        // Add basic event handlers (SAME AS FORM)
                        if (wrapPanel.Children.Count > 0)
                        {
                            var sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as StackPanel;
                            if (sp != null)
                            {
                                string filePath = item.Filename?.ToString() ?? "Unknown";
                                bool isFolder = item.IsFolder?.ToString().ToLower() == "true";
                                ClickEventAdder(sp, filePath, isFolder);

                                // FIX: Attach the missing Context Menu
                                AttachIconContextMenu(sp, item, fence, win);
                            }
                        }

                    }

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Successfully refreshed {sortedItems.Count} icons for fence '{fence.Title}' using form approach");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing fence using form approach: {ex.Message}");
            }
        }

        /// <summary>
        /// Helper method to find WrapPanel in fence window
        /// </summary>
        private static WrapPanel FindWrapPanel(NonActivatingWindow win)
        {
            try
            {
                var border = win.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                var scrollViewer = dockPanel?.Children.OfType<ScrollViewer>().FirstOrDefault();
                return scrollViewer?.Content as WrapPanel;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error finding WrapPanel: {ex.Message}");
                return null;
            }
        }



        private static void UpdateIcon(StackPanel sp, string filePath, bool isFolder)
        {
            if (System.Windows.Application.Current == null) return;

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                // 1. OPTIMIZATION: Cache Check
                bool fileExists = System.IO.File.Exists(filePath) || System.IO.Directory.Exists(filePath);
                if (!fileExists) return;

                DateTime currentLastWrite;
                try { currentLastWrite = System.IO.File.GetLastWriteTime(filePath); }
                catch { return; }

                bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
                bool isUrlFile = System.IO.Path.GetExtension(filePath).ToLower() == ".url";
                bool targetValid = true;

                if (isShortcut && !Utility.IsStoreAppShortcut(filePath))
                {
                    string tPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
                    targetValid = !string.IsNullOrEmpty(tPath) &&
                                 (System.IO.File.Exists(tPath) || System.IO.Directory.Exists(tPath));
                }
                else if (!isShortcut)
                {
                    targetValid = System.IO.File.Exists(filePath) || System.IO.Directory.Exists(filePath);
                }

                bool isNowBroken = !targetValid;

                if (_iconStates.TryGetValue(filePath, out var lastState))
                {
                    if (lastState.LastWrite == currentLastWrite && lastState.IsBroken == isNowBroken)
                        return; // Cache hit
                }
                _iconStates[filePath] = (currentLastWrite, isNowBroken);

                // 2. Identify Web Link (Do NOT return early anymore)
                bool isWebLink = isUrlFile;
                if (!isWebLink)
                {
                    try
                    {
                        foreach (var fence in FenceDataManager.FenceData)
                        {
                            if (fence.ItemsType?.ToString() == "Data")
                            {
                                bool CheckList(JArray list)
                                {
                                    if (list == null) return false;
                                    foreach (var item in list)
                                    {
                                        string itemPath = item["Filename"]?.ToString();
                                        if (!string.IsNullOrEmpty(itemPath) && string.Equals(itemPath, filePath, StringComparison.OrdinalIgnoreCase))
                                            return item["IsLink"]?.ToObject<bool>() ?? false;
                                    }
                                    return false;
                                }
                                if (CheckList(fence.Items as JArray)) { isWebLink = true; break; }
                                var tabs = fence.Tabs as JArray;
                                if (tabs != null)
                                {
                                    foreach (var tab in tabs) if (CheckList(tab["Items"] as JArray)) { isWebLink = true; goto FoundLink; }
                                }
                            }
                        }
                    FoundLink:;
                    }
                    catch { }
                }

                System.Windows.Controls.Image ico = sp.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                if (ico == null) return;

                // 3. MAIN ICON LOGIC
                ImageSource newIcon = null;
                string targetPath = filePath;

                if (isShortcut) targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);

                bool targetExists = System.IO.File.Exists(targetPath) || System.IO.Directory.Exists(targetPath);
                bool isTargetFolder = System.IO.Directory.Exists(targetPath);
                if (isShortcut && isTargetFolder) isFolder = true;

                if (isShortcut) BackupOrRestoreShortcut(filePath, targetExists, isFolder);

                // --- PRIORITY LOGIC START ---

                // CASE A: WEB LINKS
                if (isWebLink)
                {
                    // 1. Try Custom Icon first
                    var urlIcon = GetUrlCustomIcon(filePath); // <--- USE NEW HELPER
                    if (urlIcon.Path != null)
                    {
                        newIcon = IconManager.ExtractIconFromFile(urlIcon.Path, urlIcon.Index);
                    }
                    else
                    {
                        // Try WshShell only if manual failed (legacy fallback)
                        try { /* existing WshShell logic */ } catch { }
                    }

                    // 2. Fallback to Theme
                    if (newIcon == null)
                    {
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));
                    }
                }
                // CASE B: FOLDERS
                else if (isFolder)
                {
                    // 1. Broken Target -> White X
                    if (!FilePathUtilities.DoesFolderExist(filePath, isFolder))
                    {
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                    }

                    // 2. Custom Icon
                    if (newIcon == null && isShortcut)
                    {
                        try
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                            if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
                            {
                                string[] iconParts = shortcut.IconLocation.Split(',');
                                string iconPath = iconParts[0];
                                int iconIndex = 0;
                                if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex))
                                    iconIndex = parsedIndex;

                                if (System.IO.File.Exists(iconPath))
                                    newIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                            }
                        }
                        catch { }
                    }

                    // 3. Theme Fallback
                    if (newIcon == null)
                    {
                        NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(sp);
                        string fenceId = win?.Tag?.ToString();
                        var fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                        bool isPortal = fence != null && fence.ItemsType?.ToString() == "Portal";

                        if (isPortal) newIcon = Utility.GetShellIcon(filePath, true);
                        else newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    }
                }
                // CASE C: BROKEN FILES
                else if (!targetExists)
                {
                    if (Utility.IsStoreAppShortcut(filePath))
                        newIcon = Utility.GetShellIcon(filePath, isFolder);
                    else
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                }
                // CASE D: SHORTCUT FILES
                else if (isShortcut)
                {
                    try
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                        if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
                        {
                            string[] iconParts = shortcut.IconLocation.Split(',');
                            string iconPath = iconParts[0];
                            int iconIndex = 0;
                            if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex)) iconIndex = parsedIndex;

                            if (System.IO.File.Exists(iconPath))
                                newIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                        }
                    }
                    catch { }

                    if (newIcon == null)
                    {
                        // Clean icon fallback
                        if (targetPath != null && System.IO.File.Exists(targetPath))
                            newIcon = Utility.GetShellIcon(targetPath, false);
                        else
                            newIcon = Utility.GetShellIcon(filePath, false);
                    }
                }
                // CASE E: STANDARD FILES
                else
                {
                    try { newIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource(); }
                    catch { newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png")); }
                }

                // Apply
                if (ico.Source != newIcon && newIcon != null)
                {
                    ico.Source = newIcon;
                    if (IconManager.IconCache.ContainsKey(filePath))
                    {
                        IconManager.IconCache[filePath] = newIcon;
                    }
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
                EnableSounds = SettingsManager.EnableSounds,
                TintValue = SettingsManager.TintValue,
                MenuTintValue = SettingsManager.MenuTintValue,
                MenuIcon = SettingsManager.MenuIcon,
                LockIcon = SettingsManager.LockIcon,
                SelectedColor = SettingsManager.SelectedColor,
                IsLogEnabled = SettingsManager.IsLogEnabled,
                singleClickToLaunch = SettingsManager.SingleClickToLaunch,
                LaunchEffect = SettingsManager.LaunchEffect,
                CheckNetworkPaths = false // Keep this as is
            };


            if (System.Windows.Application.Current != null)
            {

                // Force UI update on the main thread
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                int updatedItems = 0;
                foreach (var win in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>())
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

   

        // --- LIVE EVENT ENGINE (v2.5. 4.183) ---
        // 1. Uses Static Handlers (Prevents Event Leaks).
        // 2. Reads SettingsManager directly (Fixes Runtime Updates).
        // 3. Restores critical Folder Shortcut validation.

        public static void ClickEventAdder(StackPanel sp, string path, bool isFolder, string arguments = null)
        {
            // Store only path, isFolder, and arguments in Tag
            sp.Tag = new { FilePath = path, IsFolder = isFolder, Arguments = arguments };

            // Check if path is a shortcut and correct isFolder for folder shortcuts
            // This logic is preserved from your working version
            bool isShortcut = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
            string targetPath = isShortcut ? Utility.GetShortcutTarget(path) : path;
            if (isShortcut && System.IO.Directory.Exists(targetPath))
            {
                isFolder = true;
            }

            // Remove existing handlers to be safe (though works best with static, harmless here)
            sp.MouseLeftButtonDown -= MouseDownHandler;

            void MouseDownHandler(object sender, MouseButtonEventArgs e)
            {
                if (e.ChangedButton != MouseButton.Left) return;

                // --- FIX: Runtime Correction for Extension Mismatch ---
                if (!System.IO.File.Exists(path) && !System.IO.Directory.Exists(path))
                {
                    if (path.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase))
                    {
                        string potentialUrlPath = System.IO.Path.ChangeExtension(path, ".url");
                        if (System.IO.File.Exists(potentialUrlPath))
                        {
                            path = potentialUrlPath;
                        }
                    }
                }

                // --- CTRL + CLICK LOGIC (MODIFIED FOR PORTAL) ---
                if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                {
                    // 1. Detect Context (Fence Type)
                    NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(sp);
                    string fenceId = win?.Tag?.ToString();
                    dynamic fence = null;
                    if (!string.IsNullOrEmpty(fenceId))
                        fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);

                    // 2. Portal Fence Logic
                    if (fence != null && fence.ItemsType?.ToString() == "Portal")
                    {
                        if (isFolder)
                        {
                            // Trigger Navigation
                            NavigatePortalFence(fence, path);
                            e.Handled = true;
                            return;
                        }
                        else
                        {
                            // Do nothing for files in Portal (No Drag Support)
                            e.Handled = true;
                            return;
                        }
                    }

                    // 3. Data Fence Logic (Standard Drag)
                    System.Windows.Point mousePosition = e.GetPosition(sp);
                    IconDragDropManager.StartIconDrag(sp, mousePosition);
                    e.Handled = true;
                    return;
                }
                // ------------------------------------------------


                //// Check for Ctrl+Click to start drag operation
                //if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                //{
                //    System.Windows.Point mousePosition = e.GetPosition(sp);
                //    IconDragDropManager.StartIconDrag(sp, mousePosition);
                //    e.Handled = true;
                //    return;
                //}

                // FIX: Read LIVE Single Click Setting directly
                // This guarantees the setting applies immediately after saving Options
                bool singleClickToLaunch = SettingsManager.SingleClickToLaunch;

                try
                {
                    bool isShortcutLocal = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
                    bool targetExists;
                    string resolvedPath = path;

                    if (isShortcutLocal)
                    {
                        resolvedPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(path);
                        if (string.IsNullOrEmpty(resolvedPath))
                        {
                            targetExists = false;
                        }
                        else
                        {
                            targetExists = isFolder ? System.IO.Directory.Exists(resolvedPath) : System.IO.File.Exists(resolvedPath);
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

                    bool isStoreApp = false;
                    string scrapedPath = null;

                    if (!targetExists && isShortcutLocal)
                    {
                        isStoreApp = Utility.IsStoreAppShortcut(path);
                        if (!isStoreApp)
                        {
                            scrapedPath = GetScrapedNetworkPath(path);
                            if (!string.IsNullOrEmpty(scrapedPath))
                            {
                                targetExists = true;
                            }
                        }
                    }

                    if (!targetExists && !isStoreApp)
                    {
                        return;
                    }

                    if (singleClickToLaunch && e.ClickCount == 1)
                    {
                        LaunchItem(sp, path, isFolder, arguments);
                        e.Handled = true;
                    }
                    else if (!singleClickToLaunch && e.ClickCount == 2)
                    {
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
                        System.Windows.Point screenPosition = sp.PointToScreen(e.GetPosition(sp));
                        IconDragDropManager.HandleDragMove(screenPosition);
                    }
                    catch { }
                }
            };

            sp.MouseLeftButtonUp += (sender, e) =>
            {
                if (IconDragDropManager.IsDragging)
                {
                    try
                    {
                        WrapPanel wrapPanel = FindVisualParent<WrapPanel>(sp);
                        if (wrapPanel != null)
                        {
                            System.Windows.Point finalPosition = e.GetPosition(wrapPanel);
                            IconDragDropManager.CompleteDrag(finalPosition);
                        }
                        else IconDragDropManager.CancelDrag();
                        e.Handled = true;
                    }
                    catch
                    {
                        IconDragDropManager.CancelDrag();
                    }
                }
            };

            sp.KeyUp += (sender, e) =>
            {
                if (IconDragDropManager.IsDragging &&
                    (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl) &&
                    !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                {
                    IconDragDropManager.CancelDrag();
                    e.Handled = true;
                }
            };
        }
        // --- STATIC HANDLERS ---

        private static void Icon_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;
            var sp = sender as StackPanel;
            if (sp == null || sp.Tag == null) return;

            // 1. Extract Data safely using dynamic
            dynamic data = sp.Tag;
            string path = data.FilePath;
            bool isFolder = data.IsFolder;
            string arguments = data.Arguments;

            // 2. Extension Mismatch Fix (Legacy Logic)
            if (!System.IO.File.Exists(path) && !System.IO.Directory.Exists(path))
            {
                if (path.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase))
                {
                    string urlPath = System.IO.Path.ChangeExtension(path, ".url");
                    if (System.IO.File.Exists(urlPath))
                    {
                        path = urlPath;
                        // Update Tag for consistency
                        sp.Tag = new { FilePath = path, IsFolder = isFolder, Arguments = arguments };
                    }
                }
            }


     
            // 3. Ctrl + Click Logic (Navigation vs Drag)
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                // Detect Fence Context
                NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(sp);
                string fenceId = win?.Tag?.ToString();
                dynamic fence = null;
                if (!string.IsNullOrEmpty(fenceId))
                    fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);

                // --- NEW: PORTAL NAVIGATION LOGIC ---
                if (fence != null && fence.ItemsType?.ToString() == "Portal")
                {
                    // Only navigate if it is a Folder
                    if (isFolder)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Portal Navigation triggered for: {path}");
                        NavigatePortalFence(fence, path);
                        e.Handled = true;
                        return;
                    }
                    else
                    {
                        // It's a file in a Portal Fence. 
                        // Do NOT drag (unsupported). Do NOT navigate.
                        // Future: Open With...
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Ctrl+Click on Portal File ignored (No Drag support)");
                        e.Handled = true;
                        return;
                    }
                }
                // ------------------------------------

                // DEFAULT: Data Fence Drag
                IconDragDropManager.StartIconDrag(sp, e.GetPosition(sp));
                e.Handled = true;
                return;
            }



            // 4. LAUNCH LOGIC (Reads LIVE Settings directly)
            bool singleClick = SettingsManager.SingleClickToLaunch;

            // Target Validation Logic
            bool isShortcut = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
            bool targetExists = false;

            if (isShortcut)
            {
                string resolved = FilePathUtilities.GetShortcutTargetUnicodeSafe(path);
                if (string.IsNullOrEmpty(resolved)) resolved = GetScrapedNetworkPath(path);

                if (!string.IsNullOrEmpty(resolved))
                {
                    targetExists = isFolder ? System.IO.Directory.Exists(resolved) : System.IO.File.Exists(resolved);
                    // Correction: If we thought it was a file but it's a folder, update state
                    if (!isFolder && System.IO.Directory.Exists(resolved)) { isFolder = true; targetExists = true; }
                }
            }
            else
            {
                targetExists = isFolder ? System.IO.Directory.Exists(path) : System.IO.File.Exists(path);
            }

            // Check Special Cases
            bool isStoreApp = isShortcut && Utility.IsStoreAppShortcut(path);
            bool isScraped = !targetExists && isShortcut && !string.IsNullOrEmpty(GetScrapedNetworkPath(path));

            if (targetExists || isStoreApp || isScraped)
            {
                if (singleClick && e.ClickCount == 1)
                {
                    LaunchItem(sp, path, isFolder, arguments);
                    e.Handled = true;
                }
                else if (!singleClick && e.ClickCount == 2)
                {
                    LaunchItem(sp, path, isFolder, arguments);
                    e.Handled = true;
                }
            }
        }

        private static void Icon_MouseMove(object sender, MouseEventArgs e)
        {
            if (IconDragDropManager.IsDragging && sender is StackPanel sp)
            {
                try { IconDragDropManager.HandleDragMove(sp.PointToScreen(e.GetPosition(sp))); } catch { }
            }
        }

        private static void Icon_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (IconDragDropManager.IsDragging && sender is StackPanel sp)
            {
                try
                {
                    var wrapPanel = FindVisualParent<WrapPanel>(sp);
                    if (wrapPanel != null) IconDragDropManager.CompleteDrag(e.GetPosition(wrapPanel));
                    else IconDragDropManager.CancelDrag();
                    e.Handled = true;
                }
                catch { IconDragDropManager.CancelDrag(); }
            }
        }

        private static void Icon_KeyUp(object sender, KeyEventArgs e)
        {
            if (IconDragDropManager.IsDragging &&
               (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl) &&
               !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
            {
                IconDragDropManager.CancelDrag();
                e.Handled = true;
            }
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
                // 1. Visual Feedback
                NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(sp);

                // FIX: Use ID (Tag) lookup instead of Title (Robust)
                string fenceId = win?.Tag?.ToString();
                dynamic fence = null;

                if (!string.IsNullOrEmpty(fenceId))
                {
                    fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                }

                // Fallback to Title only if ID lookup failed
                if (fence == null)
                {
                    fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Title == win?.Title);
                }

                // FIX: Read LIVE Global Setting directly from SettingsManager
                // This bypasses the stale _options cache.
                LaunchEffectsManager.LaunchEffect effect = SettingsManager.LaunchEffect;

                // Check for Fence-Specific Override
                string customEffect = fence?.CustomLaunchEffect?.ToString();
                if (!string.IsNullOrEmpty(customEffect))
                {
                    try
                    {
                        effect = (LaunchEffectsManager.LaunchEffect)Enum.Parse(typeof(LaunchEffectsManager.LaunchEffect), customEffect, true);
                    }
                    catch { }
                }

                LaunchEffectsManager.ExecuteLaunchEffect(sp, effect);

                // 2. Path Resolution
                string fullPath = path;
                try { fullPath = System.IO.Path.GetFullPath(path); } catch { }

                string extension = System.IO.Path.GetExtension(fullPath).ToLower();
                bool isUrlFile = extension == ".url";
                bool isLnk = extension == ".lnk";

                string targetPath = fullPath;
                string workingDirectory = "";
                string finalArguments = arguments ?? "";

                // 3. Resolve Target
                if (isUrlFile)
                {
                    targetPath = ExtractUrlFromUrlFile(fullPath);
                    if (string.IsNullOrEmpty(targetPath))
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Invalid .url file: {path}", "Error");
                        return;
                    }
                }
                else if (isLnk)
                {
                    string resolved = FilePathUtilities.GetShortcutTargetUnicodeSafe(fullPath);
                    if (string.IsNullOrEmpty(resolved)) resolved = GetScrapedNetworkPath(fullPath);

                    if (!string.IsNullOrEmpty(resolved))
                    {
                        targetPath = resolved;
                        try
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(fullPath);
                            string rawTarget = shortcut.TargetPath?.ToLower() ?? "";
                            if (!rawTarget.EndsWith("explorer.exe"))
                            {
                                if (string.IsNullOrEmpty(finalArguments)) finalArguments = shortcut.Arguments;
                                workingDirectory = shortcut.WorkingDirectory;
                            }
                        }
                        catch { }
                    }
                }

                // 4. Classification
                bool isWebUrl = IsWebUrl(targetPath);
                bool isStoreApp = isLnk && Utility.IsStoreAppShortcut(fullPath);
                bool isNetworkRoot = !string.IsNullOrEmpty(targetPath) && targetPath.StartsWith(@"\\") && !targetPath.Contains(@":\");
                bool isSpecialPath = IsSpecialWindowsPath(targetPath);

                bool isTargetFolder = false;
                bool targetExists = false;

                if (!isWebUrl && !isSpecialPath && !isNetworkRoot && !isStoreApp)
                {
                    isTargetFolder = System.IO.Directory.Exists(targetPath);
                    targetExists = isTargetFolder || System.IO.File.Exists(targetPath);
                }

                // 5. Validation Guard
                if (!targetExists && !isWebUrl && !isSpecialPath && !isStoreApp && !isNetworkRoot)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Target not found: {targetPath}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Target not found:\n{targetPath}", "Launch Error");
                    return;
                }

                // 6. Admin Logic
                bool alwaysRunAsAdmin = false;
                try
                {
                    if (fence != null && fence.ItemsType?.ToString() == "Data")
                    {
                        var items = fence.Items as JArray ?? new JArray();
                        bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

                        if (tabsEnabled)
                        {
                            var tabs = fence.Tabs as JArray ?? new JArray();
                            int currentTabIndex = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");
                            if (currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                            {
                                var currentTab = tabs[currentTabIndex] as JObject;
                                items = currentTab?["Items"] as JArray ?? items;
                            }
                        }

                        var matchingItem = items.FirstOrDefault(i => string.Equals(
                            System.IO.Path.GetFullPath(i["Filename"]?.ToString() ?? ""),
                            fullPath, StringComparison.OrdinalIgnoreCase));

                        if (matchingItem != null)
                            alwaysRunAsAdmin = Convert.ToBoolean(matchingItem["AlwaysRunAsAdmin"] ?? false);
                    }
                }
                catch { }

                // 7. Execution
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.UseShellExecute = true;

                if (alwaysRunAsAdmin && !isWebUrl && !isSpecialPath && !isTargetFolder && !isNetworkRoot)
                {
                    psi.Verb = "runas";
                }

                if (isStoreApp)
                {
                    psi.FileName = "explorer.exe";
                    psi.Arguments = $"\"{fullPath}\"";
                    psi.WorkingDirectory = "";
                }
                else if (isWebUrl)
                {
                    psi.FileName = targetPath;
                }
                else if (isTargetFolder || isNetworkRoot)
                {
                    psi.FileName = "explorer.exe";
                    psi.Arguments = $"\"{targetPath}\"";
                }
                else
                {
                    psi.FileName = targetPath;
                    if (!string.IsNullOrEmpty(finalArguments)) psi.Arguments = finalArguments;

                    if (!string.IsNullOrEmpty(workingDirectory) && System.IO.Directory.Exists(workingDirectory))
                        psi.WorkingDirectory = workingDirectory;
                    else
                    {
                        string dir = System.IO.Path.GetDirectoryName(targetPath);
                        if (!string.IsNullOrEmpty(dir) && System.IO.Directory.Exists(dir))
                            psi.WorkingDirectory = dir;
                    }
                }

                Process.Start(psi);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Launched: {targetPath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Launch Error: {ex.Message}");
                if (!SettingsManager.SuppressLaunchWarnings)
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error launching: {ex.Message}", "Launch Error");
            }
        }



        // Determines if a path is a web URL
        private static bool IsWebUrl(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;

            return path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ||
           path.StartsWith("steam://", StringComparison.OrdinalIgnoreCase);
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

        // Helper method for network path detection
        public static bool IsNetworkPath(string filePath)
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
                    var wrapPanel = FenceUtilities.FindWrapPanel(window);

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

                // 1. Read fresh data
                string freshTargetPath = shortcutPath;
                string freshArguments = "";
                bool isFolder = false;
                string workingDirectory = "";

                // Identify Type
                bool isUrl = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".url";
                bool isLnk = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".lnk";

                if (isLnk)
                {
                    try
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                        freshTargetPath = shortcut.TargetPath ?? shortcutPath;
                        freshArguments = shortcut.Arguments ?? "";
                        workingDirectory = shortcut.WorkingDirectory ?? "";
                        isFolder = System.IO.Directory.Exists(freshTargetPath);
                    }
                    catch { }
                }
                else if (isUrl)
                {
                    // For URLs, we can try to extract the clean URL, but IsFolder is always false
                    string url = CoreUtilities.ExtractWebUrlFromFile(shortcutPath);
                    if (!string.IsNullOrEmpty(url)) freshTargetPath = url;
                }

                // 2. Update Tag
                iconPanel.Tag = new
                {
                    FilePath = shortcutPath,
                    IsFolder = isFolder,
                    Arguments = freshArguments
                };

                // 3. Update Text
                if (!string.IsNullOrEmpty(newDisplayName))
                {
                    var textBlock = iconPanel.Children.OfType<TextBlock>().FirstOrDefault();
                    if (textBlock != null)
                    {
                        string displayText = newDisplayName.Length > SettingsManager.MaxDisplayNameLength
                            ? newDisplayName.Substring(0, SettingsManager.MaxDisplayNameLength) + "..."
                            : newDisplayName;
                        textBlock.Text = displayText;
                    }
                }

                // 4. Update ToolTip
                string toolTipText = $"{System.IO.Path.GetFileName(shortcutPath)}\nTarget: {freshTargetPath}";
                if (!string.IsNullOrEmpty(freshArguments)) toolTipText += $"\nArguments: {freshArguments}";
                iconPanel.ToolTip = new ToolTip { Content = toolTipText };

                // 5. UPDATE ICON IMAGE (The Critical Fix)
                var ico = iconPanel.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                if (ico != null)
                {
                    ImageSource newIcon = null;

                    // A. Try Manual Extraction for .url (Fixes Runtime Update)
                    if (isUrl)
                    {
                        var urlIcon = GetUrlCustomIcon(shortcutPath);
                        if (urlIcon.Path != null)
                        {
                            newIcon = IconManager.ExtractIconFromFile(urlIcon.Path, urlIcon.Index);
                        }
                    }
                    // B. Try WshShell for .lnk
                    else if (isLnk)
                    {
                        try
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                            if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
                            {
                                string[] parts = shortcut.IconLocation.Split(',');
                                if (System.IO.File.Exists(parts[0]))
                                {
                                    int idx = 0;
                                    if (parts.Length > 1) int.TryParse(parts[1], out idx);
                                    newIcon = IconManager.ExtractIconFromFile(parts[0], idx);
                                }
                            }
                        }
                        catch { }
                    }

                    // C. Fallback to Theme/Shell Icon
                    if (newIcon == null)
                    {
                        if (isUrl)
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));
                        }
                        else if (isFolder)
                        {
                            // Check for broken target
                            if (!FilePathUtilities.DoesFolderExist(shortcutPath, true))
                                newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                            else
                                newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                        }
                        else
                        {
                            // Standard file fallback
                            newIcon = Utility.GetShellIcon(freshTargetPath, false);
                            if (newIcon == null) newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        }
                    }

                    // Apply and Update Cache
                    if (newIcon != null)
                    {
                        ico.Source = newIcon;
                        IconManager.IconCache[shortcutPath] = newIcon; // Force update cache
                    }
                }

                // 6. Re-attach Events and Update Data
                ClickEventAdder(iconPanel, shortcutPath, isFolder, freshArguments);
                UpdateFenceDataForIcon(shortcutPath, newDisplayName, parentWindow);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Complete icon refresh successful for: {shortcutPath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error in complete icon refresh: {ex.Message}");
            }
        }
        // Updates fence JSON data after icon refresh


        // Updates fence JSON data after icon refresh
        private static void UpdateFenceDataForIcon(string shortcutPath, string newDisplayName, NonActivatingWindow parentWindow)
        {
            try
            {
                if (string.IsNullOrEmpty(newDisplayName)) return;

                string fenceId = parentWindow.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                var fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (fence == null) return;

                // FIX: Search for the item in BOTH Main Items and ALL Tabs
                bool found = false;

                // Helper function to search and update a list
                bool SearchList(JArray list)
                {
                    if (list == null) return false;
                    foreach (var item in list)
                    {
                        string itemFilename = item["Filename"]?.ToString();
                        if (!string.IsNullOrEmpty(itemFilename) &&
                            string.Equals(System.IO.Path.GetFullPath(itemFilename), System.IO.Path.GetFullPath(shortcutPath), StringComparison.OrdinalIgnoreCase))
                        {
                            item["DisplayName"] = newDisplayName;
                            return true; // Found and updated
                        }
                    }
                    return false;
                }

                // 1. Check Main Items
                if (SearchList(fence.Items as JArray)) found = true;

                // 2. Check Tabs (if enabled or present)
                if (!found)
                {
                    var tabs = fence.Tabs as JArray;
                    if (tabs != null)
                    {
                        foreach (var tab in tabs)
                        {
                            if (SearchList(tab["Items"] as JArray))
                            {
                                found = true;
                                break;
                            }
                        }
                    }
                }

                if (found)
                {
                    FenceDataManager.SaveFenceData();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Updated JSON DisplayName for: {shortcutPath}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Could not find item in JSON to update name: {shortcutPath}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error updating fence JSON: {ex.Message}");
            }
        }



        // Size feedback during resizing
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

                dynamic fenceData = GetFenceData().FirstOrDefault(f => f.Title == fence.Title);
                if (fenceData != null)
                {
                    fenceData.Width = snappedWidth;
                    fenceData.Height = snappedHeight;
                    FenceDataManager.SaveFenceData();
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




        // --- HELPER: Efficient Dead Shortcut Check using Cached States ---
        private static bool HasDeadShortcuts(dynamic fence)
        {
            try
            {
                // Only Data fences track dead shortcuts
                if (fence.ItemsType?.ToString() != "Data") return false;

                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                if (tabsEnabled)
                {
                    var tabs = fence.Tabs as JArray;
                    if (tabs != null)
                    {
                        foreach (var tab in tabs)
                        {
                            var items = tab["Items"] as JArray;
                            if (items != null && CheckItemsForDead(items)) return true;
                        }
                    }
                }
                else
                {
                    var items = fence.Items as JArray;
                    if (items != null && CheckItemsForDead(items)) return true;
                }
            }
            catch { }
            return false;
        }

        private static bool CheckItemsForDead(JArray items)
        {
            foreach (var item in items)
            {
                string path = item["Filename"]?.ToString();
                if (!string.IsNullOrEmpty(path))
                {
                    // Check our cache which is updated by the background TargetChecker
                    // This avoids scanning the disk on the UI thread
                    if (_iconStates.TryGetValue(path, out var state))
                    {
                        if (state.IsBroken) return true;
                    }
                }
            }
            return false;
        }










        // Add this static method inside your FenceManager class
        private static IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_GETMINMAXINFO = 0x0024;

            if (msg == WM_GETMINMAXINFO)
            {
                // Get the screen information for the current monitor
                var screen = System.Windows.Forms.Screen.FromHandle(hwnd);
                if (screen != null)
                {
                    // Define the limit: 90% of the working area
                    int maxWidth = (int)(screen.WorkingArea.Width * 0.90);
                    int maxHeight = (int)(screen.WorkingArea.Height * 0.90);

                    // Marshal the structure
                    MINMAXINFO mmi = (MINMAXINFO)Marshal.PtrToStructure(lParam, typeof(MINMAXINFO));

                    // Force the "Maximized" size to be 90%, not 100%
                    mmi.ptMaxSize.x = maxWidth;
                    mmi.ptMaxSize.y = maxHeight;

                    // Force the user-draggable limit to be 90%
                    mmi.ptMaxTrackSize.x = maxWidth;
                    mmi.ptMaxTrackSize.y = maxHeight;

                    Marshal.StructureToPtr(mmi, lParam, true);
                    handled = true;
                }
            }
            return IntPtr.Zero;
        }



        private static string GetSafeProperty(dynamic obj, string propName)
        {
            try
            {
                if (obj is JObject jObj && jObj[propName] != null) return jObj[propName].ToString();
                return obj.GetType().GetProperty(propName)?.GetValue(obj, null)?.ToString() ?? "";
            }
            catch { return ""; }
        }



        private static void UpdateSizeFeedback(object sender, SizeChangedEventArgs e)
        {
            var fence = sender as NonActivatingWindow;
            if (fence != null)
            {
                ShowSizeFeedback(fence.Width, fence.Height);
            }
        }



        #region Registry Monitor for Single Instance

        private static DispatcherTimer _registryMonitorTimer;

        /// <summary>
        /// Starts monitoring registry for single instance trigger
        /// Called once during application startup after all initialization is complete
        /// </summary>
        private static void StartRegistryMonitor()
        {
            try
            {
                _registryMonitorTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(500) // Check every 500ms
                };

                _registryMonitorTimer.Tick += OnRegistryMonitorTick;
                _registryMonitorTimer.Start();

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    "RegistryMonitor: Started monitoring for single instance triggers");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryMonitor: Error starting monitor: {ex.Message}");
            }
        }


 
        // <summary>
        /// Registry monitor tick event - checks for trigger and activates effect
        /// DEBUG VERSION with detailed logging
        /// </summary>
        private static void OnRegistryMonitorTick(object sender, EventArgs e)
        {
            try
            {
                // DEBUG: Log every few ticks to confirm monitor is running
                _registryMonitorTickCount++;
                if (_registryMonitorTickCount % 10 == 0) // Every 5 seconds (10 ticks * 500ms)
                {
                    //LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    //    $"RegistryMonitor: Heartbeat check #{_registryMonitorTickCount}");
                }

                string triggerValue = RegistryHelper.CheckForTrigger();

                if (!string.IsNullOrEmpty(triggerValue))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        $"RegistryMonitor: TRIGGER DETECTED! Value: {triggerValue}");

                    // Stop timer temporarily to prevent multiple triggers
                    _registryMonitorTimer.Stop();

                    // Delete trigger immediately to prevent re-triggering
                    bool deleted = RegistryHelper.DeleteTrigger();

                    //LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    //    $"RegistryMonitor: Trigger deleted: {deleted}");

                    //// Activate the single instance effect
                    //LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    //    "RegistryMonitor: About to activate effect...");

                    ActivateSingleInstanceEffect();

                    //LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    //    "RegistryMonitor: Effect activation completed");

                    // Restart timer after 2 seconds
                    var restartTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                    restartTimer.Tick += (s2, e2) =>
                    {
                        _registryMonitorTimer.Start();
                        restartTimer.Stop();
                        //LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                        //    "RegistryMonitor: Resumed monitoring");
                    };
                    restartTimer.Start();
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"RegistryMonitor: Error in monitor tick: {ex.Message}");
            }
        }
        // ALSO ADD DEBUG TO ActivateSingleInstanceEffect:

        /// <summary>
        /// Activates the visual effect when another instance attempts to launch
        /// DEBUG VERSION with detailed logging
        /// </summary>
        private static void ActivateSingleInstanceEffect()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    "SingleInstanceEffect: Starting effect activation...");

                // Check if InterCore is available
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    "SingleInstanceEffect: Calling InterCore.ActivateLighthouseSweep()");

                // Use InterCore's lighthouse sweep effect
                InterCore.ActivateLighthouseSweep();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    "SingleInstanceEffect: InterCore effect call completed");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"SingleInstanceEffect: Error activating effect: {ex.Message}");
            }
        }





        #endregion







    }

}