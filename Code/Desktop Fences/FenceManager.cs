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
                // Clear existing data
                FenceDataManager.FenceData?.Clear();

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
                int totalFenceCount = Application.Current.Windows.OfType<NonActivatingWindow>().Count();
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
                var allWindows = Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
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
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
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
        private static ContextMenu BuildHeartContextMenu(dynamic fence, bool showTabsOption = false)
        {
            // DEBUG: Log menu building
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                $"Building heart context menu for fence '{fence.Title}' - showTabsOption: {showTabsOption}");

            var menu = new ContextMenu();

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
                //bool result = TrayManager.Instance.ShowCustomMessageBoxForm(); // Call the method and store the result  
                bool result = MessageBoxesManager.ShowCustomMessageBoxForm();
                if (result == true)
                {
                    // NEW: Use BackupManager to handle the deletion backup instead of manual backup code
                    BackupManager.BackupDeletedFence(fence);

                    // Proceed with deletion - remove from data structures
                    FenceDataManager.FenceData.Remove(fence);
                    _heartTextBlocks.Remove(fence); // Remove from heart TextBlocks dictionary

                    // Clean up portal fence if applicable
                    if (_portalFences.ContainsKey(fence))
                    {
                        _portalFences[fence].Dispose();
                        _portalFences.Remove(fence);
                    }

                    // Save updated fence data
                    FenceDataManager.SaveFenceData();

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

            // TABS FEATURE: Hidden menu item for Ctrl+click (Data fences only)
            if (showTabsOption)
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    "Checking if tabs menu item should be added");

                // Only show tabs option for Data fences, not Portal fences
                bool isDataFence = fence.ItemsType?.ToString() == "Data";

                if (isDataFence)
                {
                    menu.Items.Add(new Separator());

                    // Check current tabs status
                    bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                    string checkmark = tabsEnabled ? "✓ " : "";

                    var enableTabsItem = new MenuItem { Header = $"{checkmark}Enable Tabs On This Fence" };
                    enableTabsItem.Click += (s, e) => ToggleFenceTabs(fence);
                    menu.Items.Add(enableTabsItem);

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Added tabs menu item for Data fence: '{checkmark}Enable Tabs On This Fence'");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Skipping tabs menu item - fence type '{fence.ItemsType}' not supported");
                }
            }


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
                   (validUri.Scheme == Uri.UriSchemeHttp || validUri.Scheme == Uri.UriSchemeHttps);
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
                newItemDict["DisplayOrder"] = GetNextDisplayOrder(fence);

                //// Add to fence data
                //var items = fence.Items as JArray ?? new JArray();
                //items.Add(JObject.FromObject(newItem));

                //// Handle tabs if enabled
                //bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                //if (tabsEnabled)
                //{
                //    var tabs = fence.Tabs as JArray ?? new JArray();
                //    int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");

                //    if (currentTab >= 0 && currentTab < tabs.Count)
                //    {
                //        var activeTab = tabs[currentTab] as JObject;
                //        if (activeTab != null)
                //        {
                //            var tabItems = activeTab["Items"] as JArray ?? new JArray();
                //            tabItems.Add(JObject.FromObject(newItem));
                //        }
                //    }
                //}

                //// Save fence data
                //FenceDataManager.SaveFenceData();

                //// Add icon to UI
                //IconManager.AddIconWithFenceContext(newItem, wrapPanel, fence);


                // Find the fence in FenceDataManager.FenceData and update it properly
                string fenceId = fence.Id?.ToString();
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);

                if (fenceIndex >= 0)
                {
                    // Get the actual fence from FenceDataManager.FenceData
                    dynamic actualFence = FenceDataManager.FenceData[fenceIndex];
                    var actualItems = actualFence.Items as JArray ?? new JArray();
                    actualItems.Add(JObject.FromObject(newItem));

                    // Handle tabs if enabled
                    bool tabsEnabled = actualFence.TabsEnabled?.ToString().ToLower() == "true";
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
                            }
                        }
                    }

                    // Save fence data BEFORE creating the icon
                    FenceDataManager.SaveFenceData();

                    // Add icon to UI - now the fence data is properly saved and available
                    // Add icon to UI using FenceManager.AddIcon for proper sizing
                    FenceManager.AddIcon(newItem, wrapPanel); // This adds icon with the proper size
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
                    CreateBasicContextMenu(sp, newItem, shortcutPath);
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
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceUpdate, $"Failed to find fence '{fence.Title}' in FenceDataManager.FenceData for {propertyName} update");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate, $"Error updating {propertyName} for fence '{fence.Title}': {ex.Message}");
            }
        }


     // TABS FEATURE: Toggle tabs for a specific fence (SAFE VERSION)
        private static void ToggleFenceTabs(dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Toggling tabs for fence '{fence.Title}'");

                // Get current state
                bool currentTabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                bool newTabsEnabled = !currentTabsEnabled;

                // Find fence in FenceDataManager.FenceData for proper update
                string fenceId = fence.Id?.ToString();
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);

                if (fenceIndex >= 0)
                {
                    // Find and store the current fence window position/size
                    var windows = Application.Current.Windows.OfType<NonActivatingWindow>();
                    var currentWindow = windows.FirstOrDefault(w => w.Tag?.ToString() == fenceId);

                    double savedLeft = fence.X ?? 100;
                    double savedTop = fence.Y ?? 100;
                    double savedWidth = fence.Width ?? 200;
                    double savedHeight = fence.Height ?? 300;

                    if (currentWindow != null)
                    {
                        savedLeft = currentWindow.Left;
                        savedTop = currentWindow.Top;
                        savedWidth = currentWindow.Width;
                        savedHeight = currentWindow.Height;
                    }

                    // Update fence data
                    IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ?
                        dict : ((JObject)fence).ToObject<IDictionary<string, object>>();

                    fenceDict["TabsEnabled"] = newTabsEnabled.ToString().ToLower();
                    fenceDict["X"] = savedLeft;
                    fenceDict["Y"] = savedTop;
                    fenceDict["Width"] = savedWidth;
                    fenceDict["Height"] = savedHeight;

                    FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                    FenceDataManager.SaveFenceData();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Tabs {(newTabsEnabled ? "enabled" : "disabled")} for fence '{fence.Title}'");

                    // Run migration to handle content structure changes
                    MigrateLegacyJson();

                    // Close the current window if it exists
                    if (currentWindow != null)
                    {
                        // Remove from tracking dictionaries first
                        _heartTextBlocks.Remove(fence);
                        if (_portalFences.ContainsKey(fence))
                        {
                            _portalFences[fence].Dispose();
                            _portalFences.Remove(fence);
                        }

                        currentWindow.Close();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            "Closed existing fence window");
                    }

                    // Get the updated fence data after migration
                    var updatedFence = FenceDataManager.FenceData[fenceIndex];

                    // Recreate the fence with the new settings
                    // Use a small delay to ensure the old window is properly closed
                    Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            CreateFence(updatedFence, new TargetChecker(1000));
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                                $"Successfully recreated fence '{updatedFence.Title}' with tabs {(newTabsEnabled ? "enabled" : "disabled")}");
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                                $"Error recreating fence: {ex.Message}");
                        }
                    }), System.Windows.Threading.DispatcherPriority.Background);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error toggling tabs for fence '{fence.Title}': {ex.Message}");
            }
        }
   

        //// TABS FEATURE: Dynamically refresh the tab strip UI to show changes
        public static void RefreshTabStripUI(NonActivatingWindow fenceWindow, dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Refreshing tab strip UI for fence '{fence.Title}'");

                // Find the tab strip
                var border = fenceWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot find DockPanel for tab strip refresh");
                    return;
                }

                var tabStrip = dockPanel.Children.OfType<StackPanel>()
                    .FirstOrDefault(sp => sp.Orientation == Orientation.Horizontal && sp.Height == 20);

                if (tabStrip == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot find tab strip for refresh");
                    return;
                }

                // Clear existing tab buttons
                tabStrip.Children.Clear();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cleared existing tab buttons");

                // Get current tab data
                var tabs = fence.Tabs as JArray ?? new JArray();
                int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");
                string fenceColorName = fence.CustomColor?.ToString() ?? SettingsManager.SelectedColor;
                var colorScheme = Utility.GenerateTabColorScheme(fenceColorName);

                // Recreate tab buttons
                for (int i = 0; i < tabs.Count; i++)
                {
                    var tab = tabs[i] as JObject;
                    if (tab == null) continue;

                    string tabName = tab["TabName"]?.ToString() ?? $"Tab {i + 1}";
                    bool isActiveTab = (i == currentTab);

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
                        FontWeight = FontWeights.Medium,
                        BorderThickness = new Thickness(1, 1, 1, 0),
                        Cursor = Cursors.Hand,
                        Template = CreateTabButtonTemplate()
                    };

                    // Apply styling
                    if (isActiveTab)
                    {
                        tabButton.Background = new SolidColorBrush(colorScheme.activeTab);
                        tabButton.Foreground = System.Windows.Media.Brushes.White;
                        tabButton.BorderBrush = new SolidColorBrush(colorScheme.borderColor);
                        tabButton.FontWeight = FontWeights.Bold;
                        tabButton.Opacity = 1.0;
                    }
                    else
                    {
                        tabButton.Background = new SolidColorBrush(colorScheme.inactiveTab);
                        tabButton.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 96, 96, 96));
                        tabButton.BorderBrush = new SolidColorBrush(colorScheme.borderColor);
                        tabButton.FontWeight = FontWeights.Normal;
                        tabButton.Opacity = 0.7;

                        // Add hover effects for inactive tabs
                        tabButton.MouseEnter += (s, e) =>
                        {
                            var btn = s as Button;
                            btn.Background = new SolidColorBrush(colorScheme.hoverTab);
                        };
                        tabButton.MouseLeave += (s, e) =>
                        {
                            var btn = s as Button;
                            btn.Background = new SolidColorBrush(colorScheme.inactiveTab);
                        };
                    }

                    // Add click handler with delay to allow double-click detection
                    int capturedIndex = i; // Capture for closure
                    System.Windows.Threading.DispatcherTimer clickTimer = null;

                    tabButton.Click += (s, e) =>
                    {
                        try
                        {
                            // Start or restart timer for delayed single-click
                            if (clickTimer != null)
                            {
                                clickTimer.Stop();
                            }

                            clickTimer = new System.Windows.Threading.DispatcherTimer
                            {
                                Interval = TimeSpan.FromMilliseconds(300) // Double-click detection window
                            };

                            clickTimer.Tick += (ts, te) =>
                            {
                                clickTimer.Stop();
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                    $"Tab button single-clicked (delayed): index {capturedIndex}");
                                SwitchTabByFence(fence, capturedIndex, fenceWindow);
                            };

                            clickTimer.Start();
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                                $"Error in tab click: {ex.Message}");
                        }
                    };
                    // Add double-click handler for quick rename
                    tabButton.MouseDoubleClick += (s, e) =>
                    {
                        try
                        {
                            // Stop the single-click timer if running
                            if (clickTimer != null)
                            {
                                clickTimer.Stop();
                            }

                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                $"Tab button double-clicked: starting rename for index {capturedIndex}");

                            e.Handled = true; // Prevent other events

                            // Start rename immediately
                            RenameTab(fence, capturedIndex, fenceWindow);
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                                $"Error in tab double-click: {ex.Message}");
                        }
                    };
                    // Add right-click context menu
                    ContextMenu tabContextMenu = new ContextMenu();
                    MenuItem miAddTab = new MenuItem { Header = "Add New Tab" };
                    MenuItem miRenameTab = new MenuItem { Header = "Rename Tab" };
                    MenuItem miDeleteTab = new MenuItem { Header = "Delete Tab" };
                    Separator miSeparator1 = new Separator();
                    MenuItem miMoveLeft = new MenuItem { Header = "Move Left", IsEnabled = false };
                    MenuItem miMoveRight = new MenuItem { Header = "Move Right", IsEnabled = false };

                    tabContextMenu.Items.Add(miAddTab);
                    tabContextMenu.Items.Add(miSeparator1);
                    tabContextMenu.Items.Add(miRenameTab);
                    tabContextMenu.Items.Add(miDeleteTab);
                    tabContextMenu.Items.Add(new Separator());
                    tabContextMenu.Items.Add(miMoveLeft);
                    tabContextMenu.Items.Add(miMoveRight);

                    // Add event handlers with captured index
                    miAddTab.Click += (s, e) => AddNewTab(fence, fenceWindow);
                    miRenameTab.Click += (s, e) => RenameTab(fence, capturedIndex, fenceWindow);
                    miDeleteTab.Click += (s, e) => DeleteTab(fence, capturedIndex, fenceWindow);

                    tabButton.ContextMenu = tabContextMenu;

                    // Add the button to tab strip
                    tabStrip.Children.Add(tabButton);
                }

                // Add the [+] button
                Button addTabButton = new Button
                {
                    Content = "+",
                    Tag = "ADD_TAB",
                    Height = 18,
                    Width = 25,
                    Margin = new Thickness(3, 0, 1, 0),
                    Padding = new Thickness(0),
                    FontSize = 12,
                    FontWeight = FontWeights.Bold,
                    FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                    BorderThickness = new Thickness(1),
                    Cursor = Cursors.Hand,
                    Template = CreateTabButtonTemplate(),
                    ToolTip = "Add new tab"
                };

                // Style the [+] button
                var addButtonColorScheme = colorScheme;
                addTabButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, addButtonColorScheme.inactiveTab.R, addButtonColorScheme.inactiveTab.G, addButtonColorScheme.inactiveTab.B));
                addTabButton.Foreground = new SolidColorBrush(addButtonColorScheme.borderColor);
                addTabButton.BorderBrush = new SolidColorBrush(addButtonColorScheme.borderColor);
                addTabButton.Opacity = 0.8;

                // Add hover effects
                addTabButton.MouseEnter += (s, e) =>
                {
                    addTabButton.Background = new SolidColorBrush(addButtonColorScheme.hoverTab);
                    addTabButton.Opacity = 1.0;
                };
                addTabButton.MouseLeave += (s, e) =>
                {
                    addTabButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, addButtonColorScheme.inactiveTab.R, addButtonColorScheme.inactiveTab.G, addButtonColorScheme.inactiveTab.B));
                    addTabButton.Opacity = 0.8;
                };

                // Add click handler with debouncing to prevent multiple rapid clicks
                bool isAddingTab = false;
                addTabButton.Click += async (s, e) =>
                {
                    if (isAddingTab) return; // Prevent rapid clicking
                    isAddingTab = true;

                    try
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Add tab button clicked");
                        AddNewTab(fence, fenceWindow);
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error adding new tab: {ex.Message}");
                    }
                    finally
                    {
                        // Reset the flag after a short delay
                        await System.Threading.Tasks.Task.Delay(500);
                        isAddingTab = false;
                    }
                };

                tabStrip.Children.Add(addTabButton);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Successfully refreshed tab strip UI: {tabs.Count} tabs + [+] button for fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error refreshing tab strip UI: {ex.Message}");
            }
        }


        // TABS FEATURE: Add new tab with random herb name
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
        public static void RenameTab(dynamic fence, int tabIndex, NonActivatingWindow fenceWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"RenameTab called for fence '{fence.Title}', tab {tabIndex}");

                // Find the specific tab button
                var border = fenceWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                var tabStrip = dockPanel.Children.OfType<StackPanel>()
                    .FirstOrDefault(sp => sp.Orientation == Orientation.Horizontal && sp.Height == 20);
                if (tabStrip == null) return;

                // Find the button for this tab index
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

                // Focus and select text with proper timing (improved from fence title editing pattern)
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

                        // Return focus to fence window (same as fence title editing)
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

                // Handle Enter key (save) and Escape (cancel) - same as fence title editing
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

                // Handle focus loss (save) - same pattern as fence title editing
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
        public static void DeleteTab(dynamic fence, int tabIndex, NonActivatingWindow fenceWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"DeleteTab called for fence '{fence.Title}', tab {tabIndex}");

                // Get fresh fence data
                string fenceId = fence.Id?.ToString();
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                var tabs = currentFence.Tabs as JArray ?? new JArray();

                // Don't allow deleting the last tab
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

                // Confirm deletion if tab has items
                if (items.Count > 0)
                {
                    //bool result = TrayManager.ShowTabDeleteConfirmationForm(tabName, items.Count);
                    bool result = MessageBoxesManager.ShowTabDeleteConfirmationForm(tabName, items.Count);
                    if (!result)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"User cancelled tab deletion for '{tabName}'");
                        return;
                    }
                }

                // Remove the tab
                tabs.RemoveAt(tabIndex);

                // Adjust CurrentTab if necessary
                int currentTab = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");
                int newCurrentTab = currentTab;

                if (currentTab >= tabIndex)
                {
                    newCurrentTab = Math.Max(0, currentTab - 1);
                }

                // Update fence data
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
                        $"Deleted tab '{tabName}' from fence '{currentFence.Title}'");

                    // Get updated fence and refresh
                    var updatedFence = FenceDataManager.FenceData[fenceIndex];
                    //     RefreshFenceContentSimple(fenceWindow, updatedFence, newCurrentTab);
                    //        RefreshTabStyling(fenceWindow, newCurrentTab);
                    // Refresh the entire tab strip to reflect deletion
                    RefreshTabStripUI(fenceWindow, FenceDataManager.FenceData[fenceIndex]);

                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Tab deleted successfully");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error deleting tab: {ex.Message}");
            }
        }

   
        // TABS FEATURE: Switch tab using fence ID (avoid stale references)
        public static void SwitchTabByFence(dynamic fence, int newTabIndex, NonActivatingWindow fenceWindow)
        {
            try
            {
                // Get fresh fence data by ID to avoid stale references
                string fenceId = fence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot switch tab: fence ID missing");
                    return;
                }

                // Find fresh fence data
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Cannot switch tab: fence with ID '{fenceId}' not found");
                    return;
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"SwitchTabByFence: fence '{currentFence.Title}', newTabIndex {newTabIndex}");

                // Validate tab index
                var tabs = currentFence.Tabs as JArray ?? new JArray();
                if (newTabIndex < 0 || newTabIndex >= tabs.Count)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                        $"Cannot switch tab: index {newTabIndex} out of bounds (0-{tabs.Count - 1})");
                    return;
                }

                // Check current tab
                int currentTabIndex = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Current tab: {currentTabIndex}, New tab: {newTabIndex}");

                if (currentTabIndex == newTabIndex)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                        $"Already on tab {newTabIndex}");
                    return;
                }

                // Update CurrentTab directly in FenceDataManager.FenceData
                IDictionary<string, object> fenceDict = currentFence as IDictionary<string, object> ??
                    ((JObject)currentFence).ToObject<IDictionary<string, object>>();
                fenceDict["CurrentTab"] = newTabIndex;

                // Find and update the fence in FenceDataManager.FenceData
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (fenceIndex >= 0)
                {
                    FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                    FenceDataManager.SaveFenceData();
                }

                // Use the fresh fence data for refreshing
                var freshFence = FenceDataManager.FenceData[fenceIndex];

                // Refresh content with fresh data
                RefreshFenceContentSimple(fenceWindow, freshFence, newTabIndex);

                // Update tab styling
                RefreshTabStyling(fenceWindow, newTabIndex);

                string tabName = tabs[newTabIndex]["TabName"]?.ToString() ?? $"Tab {newTabIndex}";
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Successfully switched to tab '{tabName}' (index {newTabIndex})");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in SwitchTabByFence: {ex.Message}");
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

                // Find the WrapPanel
                var border = fenceWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                var scrollViewer = dockPanel?.Children.OfType<ScrollViewer>().FirstOrDefault();
                var wrapPanel = scrollViewer?.Content as WrapPanel;

                if (wrapPanel == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot find WrapPanel for content refresh");
                    return;
                }

                // Clear existing content
                wrapPanel.Children.Clear();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cleared WrapPanel content");

                // Get items from the specified tab
                var tabs = fence.Tabs as JArray ?? new JArray();
                if (tabIndex >= 0 && tabIndex < tabs.Count)
                {
                    var activeTab = tabs[tabIndex] as JObject;
                    if (activeTab != null)
                    {
                        var items = activeTab["Items"] as JArray ?? new JArray();
                        string tabName = activeTab["TabName"]?.ToString() ?? $"Tab {tabIndex}";

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"Loading {items.Count} items from tab '{tabName}'");

                        // Sort and add items
                        var sortedItems = items
                            .OfType<JObject>()
                            .OrderBy(item => item["DisplayOrder"]?.Value<int>() ?? 0)
                            .ToList();

                        foreach (dynamic icon in sortedItems)
                        {
                            IconManager.AddIconWithFenceContext(icon, wrapPanel, fence);

                            // Add basic click events (simplified)
                            StackPanel sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as StackPanel;
                            if (sp != null)
                            {
                                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ?
                                    dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
                                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];

                                // Add full event handling including context menus
                                ClickEventAdder(sp, filePath, isFolder);

                                // Add context menu (replicating logic from CreateFence)
                                // Add context menu (replicating logic from CreateFence)
                                ContextMenu iconContextMenu = new ContextMenu();
                                MenuItem miEdit = new MenuItem { Header = "Edit" };
                                MenuItem miMove = new MenuItem { Header = "Move.." };
                                MenuItem miRemove = new MenuItem { Header = "Remove" };
                                MenuItem miFindTarget = new MenuItem { Header = "Find target..." };
                                MenuItem miCopyPath = new MenuItem { Header = "Copy path" };
                                MenuItem miRunAsAdmin = new MenuItem { Header = "Run as administrator" };
                                // Add Copy Item functionality - CopyPasteManager integration
                                MenuItem miCopyItem = new MenuItem { Header = "Copy Item" };

                                iconContextMenu.Items.Add(miEdit);
                                iconContextMenu.Items.Add(miMove);
                                iconContextMenu.Items.Add(miRemove);
                                iconContextMenu.Items.Add(new Separator());
                                iconContextMenu.Items.Add(miCopyItem);
                                iconContextMenu.Items.Add(new Separator());
                                iconContextMenu.Items.Add(miFindTarget);
                                iconContextMenu.Items.Add(miCopyPath);
                                iconContextMenu.Items.Add(miRunAsAdmin);

                                // Capture the fence window from the method parameter
                                var currentFenceWindow = fenceWindow; // Use the fenceWindow parameter from RefreshFenceContentSimple

                                // Add Copy Item event handler - CopyPasteManager integration
                                miCopyItem.Click += (s, e) => CopyPasteManager.CopyItem(icon, fence);

                                // Add event handlers with captured window
                                miEdit.Click += (s, e) => EditItem(icon, sp, currentFenceWindow);
                                miRemove.Click += (s, e) => RemoveItemFromTab(icon, fence, tabIndex, wrapPanel);
                                miMove.Click += (s, e) => ItemMoveDialog.ShowMoveDialog(icon, fence, wrapPanel.Dispatcher);

                                sp.ContextMenu = iconContextMenu;
                                // Add dynamic menu state updates - make Copy Item enabled and handle CTRL+Right Click
                                MenuItem miSendToDesktop = null; // Declare for conditional addition
                                iconContextMenu.Opened += (contextSender, contextArgs) =>
                                {
                                    // Update Copy Item enabled state dynamically when menu opens
                                    miCopyItem.IsEnabled = true; // Copy should always be enabled for existing items

                                    // Check if CTRL is pressed for "Send to Desktop" option
                                    bool isCtrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

                                    // Remove existing Send to Desktop if present
                                    if (miSendToDesktop != null && iconContextMenu.Items.Contains(miSendToDesktop))
                                    {
                                        iconContextMenu.Items.Remove(miSendToDesktop);
                                    }

                                    // Add Send to Desktop if CTRL is pressed
                                    if (isCtrlPressed)
                                    {
                                        miSendToDesktop = new MenuItem { Header = "Send to Desktop" };
                                        miSendToDesktop.Click += (s, e) => CopyPasteManager.SendToDesktop(icon);

                                        // Find Copy Item position and insert after it
                                        int copyItemIndex = -1;
                                        for (int i = 0; i < iconContextMenu.Items.Count; i++)
                                        {
                                            if (iconContextMenu.Items[i] is MenuItem mi && mi.Header.ToString() == "Copy Item")
                                            {
                                                copyItemIndex = i;
                                                break;
                                            }
                                        }

                                        if (copyItemIndex >= 0)
                                        {
                                            iconContextMenu.Items.Insert(copyItemIndex + 1, miSendToDesktop);
                                        }
                                    }

                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                        $"Updated icon context menu states - Copy available, CTRL pressed: {isCtrlPressed}");
                                };
                            }
                        }

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Successfully loaded {sortedItems.Count} items from tab '{tabName}'");
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

        ////// TABS FEATURE: Switch to a different tab and refresh content
        ////public static void SwitchTab(NonActivatingWindow fenceWindow, int newTabIndex)
        ////{
        ////    try
        ////    {
        ////        // Get fence data by window ID
        ////        string fenceId = fenceWindow.Tag?.ToString();
        ////        if (string.IsNullOrEmpty(fenceId))
        ////        {
        ////            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot switch tab: fence ID missing");
        ////            return;
        ////        }

        ////        // Find the fence in FenceDataManager.FenceData
        ////        var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
        ////        if (currentFence == null)
        ////        {
        ////            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Cannot switch tab: fence with ID '{fenceId}' not found");
        ////            return;
        ////        }

        ////        // Validate tab index
        ////        var tabs = currentFence.Tabs as JArray ?? new JArray();
        ////        if (newTabIndex < 0 || newTabIndex >= tabs.Count)
        ////        {
        ////            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
        ////                $"Cannot switch tab: index {newTabIndex} out of bounds (0-{tabs.Count - 1}) for fence '{currentFence.Title}'");
        ////            return;
        ////        }

        ////        // Check if already on this tab
        ////        int currentTabIndex = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");
        ////        if (currentTabIndex == newTabIndex)
        ////        {
        ////            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
        ////                $"Already on tab {newTabIndex} for fence '{currentFence.Title}'");
        ////            return;
        ////        }

        ////        // Update CurrentTab in fence data
        ////        UpdateFenceProperty(currentFence, "CurrentTab", newTabIndex.ToString(),
        ////            $"Switched to tab {newTabIndex}");

        ////        // Refresh the fence content to show new tab
        ////        RefreshFenceContent(fenceWindow, currentFence);

        ////        // Update tab button styling
        ////        RefreshTabStyling(fenceWindow, newTabIndex);

        ////        string tabName = tabs[newTabIndex]["TabName"]?.ToString() ?? $"Tab {newTabIndex}";
        ////        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
        ////            $"Successfully switched to tab '{tabName}' (index {newTabIndex}) for fence '{currentFence.Title}'");
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
        ////            $"Error switching tab: {ex.Message}");
        ////    }
        ////}

        //// TABS FEATURE: Refresh fence content to show items from current tab
        //private static void RefreshFenceContent(NonActivatingWindow fenceWindow, dynamic fence)
        //{
        //    try
        //    {
        //        // Find the WrapPanel in the fence window
        //        var border = fenceWindow.Content as Border;
        //        if (border == null) return;

        //        var dockPanel = border.Child as DockPanel;
        //        if (dockPanel == null) return;

        //        var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
        //        if (scrollViewer == null) return;

        //        var wrapPanel = scrollViewer.Content as WrapPanel;
        //        if (wrapPanel == null) return;

        //        // Clear existing content
        //        wrapPanel.Children.Clear();

        //        // Load items from current tab (using same logic as InitContent)
        //        bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
        //        JArray items = null;

        //        if (tabsEnabled)
        //        {
        //            var tabs = fence.Tabs as JArray ?? new JArray();
        //            int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");

        //            if (currentTab >= 0 && currentTab < tabs.Count)
        //            {
        //                var activeTab = tabs[currentTab] as JObject;
        //                if (activeTab != null)
        //                {
        //                    items = activeTab["Items"] as JArray ?? new JArray();
        //                    string tabName = activeTab["TabName"]?.ToString() ?? $"Tab {currentTab}";
        //                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
        //                        $"Refreshing content for tab '{tabName}' with {items.Count} items");
        //                }
        //            }
        //        }
        //        else
        //        {
        //            items = fence.Items as JArray ?? new JArray();
        //        }

        //        if (items != null)
        //        {
        //            // Sort and add items (same as InitContent logic)
        //            var sortedItems = items
        //                .OfType<JObject>()
        //                .OrderBy(item =>
        //                {
        //                    var orderToken = item["DisplayOrder"];
        //                    return orderToken?.Type == JTokenType.Integer ? orderToken.Value<int>() : 0;
        //                })
        //                .ToList();

        //            foreach (dynamic icon in sortedItems)
        //            {
        //                IconManager.AddIconWithFenceContext(icon, wrapPanel, fence);

        //                // Add existing event handlers (same as CreateFence logic)
        //                StackPanel sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as StackPanel;
        //                if (sp != null)
        //                {
        //                    IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ?
        //                        dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
        //                    string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
        //                    bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
        //                    // Use the fence's existing click event logic
        //                    ClickEventAdder(sp, filePath, isFolder);

        //                    // Don't create new TargetChecker - use the existing one from the fence
        //                    // The original fence creation already has a TargetChecker running
        //                }
        //            }

        //            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
        //                $"Refreshed fence content: loaded {sortedItems.Count} items");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
        //            $"Error refreshing fence content: {ex.Message}");
        //    }
        //}

        // TABS FEATURE: Update tab button styling to reflect active tab
        private static void RefreshTabStyling(NonActivatingWindow fenceWindow, int activeTabIndex)
        {
            try
            {
                // Find the TabStrip
                var border = fenceWindow.Content as Border;
                if (border == null) return;

                var dockPanel = border.Child as DockPanel;
                if (dockPanel == null) return;

                var tabStrip = dockPanel.Children.OfType<StackPanel>()
                    .FirstOrDefault(sp => sp.Orientation == Orientation.Horizontal && sp.Height == 20);

                if (tabStrip == null) return;

                // Get fence data for color scheme
                string fenceId = fenceWindow.Tag?.ToString();
                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                string fenceColorName = currentFence.CustomColor?.ToString() ?? SettingsManager.SelectedColor;
                var colorScheme = Utility.GenerateTabColorScheme(fenceColorName);

                // Update all tab buttons
                int tabIndex = 0;
                foreach (Button tabButton in tabStrip.Children.OfType<Button>())
                {
                    bool isActiveTab = (tabIndex == activeTabIndex);

                    if (isActiveTab)
                    {
                        // Active tab styling - more prominent
                        tabButton.Background = new SolidColorBrush(colorScheme.activeTab);
                        tabButton.BorderBrush = new SolidColorBrush(colorScheme.borderColor);
                        tabButton.FontWeight = FontWeights.Bold;
                        tabButton.Opacity = 1.0;
                        tabButton.Foreground = System.Windows.Media.Brushes.White;
                    }
                    else
                    {
                        // Inactive tab styling - more muted
                        tabButton.Background = new SolidColorBrush(colorScheme.inactiveTab);
                        tabButton.BorderBrush = new SolidColorBrush(colorScheme.borderColor);
                        tabButton.FontWeight = FontWeights.Normal;
                        tabButton.Opacity = 0.7;
                        tabButton.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 96, 96, 96));
                    }

                    tabIndex++;
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Updated tab styling: active tab is {activeTabIndex}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error refreshing tab styling: {ex.Message}");
            }
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

        // TABS FEATURE: Refresh tab colors when fence color changes
        public static void RefreshTabColors(NonActivatingWindow fenceWindow, string newColorName)
        {
            try
            {
                // Find the TabStrip in the fence window
                var border = fenceWindow.Content as Border;
                if (border == null) return;

                var dockPanel = border.Child as DockPanel;
                if (dockPanel == null) return;

                // Look for TabStrip (StackPanel with Horizontal orientation)
                var tabStrip = dockPanel.Children.OfType<StackPanel>()
                    .FirstOrDefault(sp => sp.Orientation == Orientation.Horizontal && sp.Height == 20);

                if (tabStrip == null) return; // No tabs on this fence

                // Get the fence data to determine current tab
                string fenceId = fenceWindow.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                var currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (currentFence == null) return;

                int currentTab = Convert.ToInt32(currentFence.CurrentTab?.ToString() ?? "0");

                // Generate new color scheme based on the new color
                var colorScheme = Utility.GenerateTabColorScheme(newColorName);

                // Update all tab buttons with new color scheme
                int tabIndex = 0;
                foreach (Button tabButton in tabStrip.Children.OfType<Button>())
                {
                    bool isActiveTab = (tabIndex == currentTab);

                    if (isActiveTab)
                    {
                        // Update active tab with new colors
                        tabButton.Background = new SolidColorBrush(colorScheme.activeTab);
                        tabButton.BorderBrush = new SolidColorBrush(colorScheme.borderColor);
                    }
                    else
                    {
                        // Update inactive tab with new colors
                        tabButton.Background = new SolidColorBrush(colorScheme.inactiveTab);
                        tabButton.BorderBrush = new SolidColorBrush(colorScheme.borderColor);

                        // Update hover effects with new colors
                        tabButton.MouseEnter -= TabButton_MouseEnter; // Remove old handler
                        tabButton.MouseLeave -= TabButton_MouseLeave; // Remove old handler

                        tabButton.MouseEnter += (s, e) =>
                        {
                            var btn = s as Button;
                            btn.Background = new SolidColorBrush(colorScheme.hoverTab);
                        };
                        tabButton.MouseLeave += (s, e) =>
                        {
                            var btn = s as Button;
                            btn.Background = new SolidColorBrush(colorScheme.inactiveTab);
                        };
                    }

                    tabIndex++;
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Refreshed tab colors for fence '{currentFence.Title}' with color '{newColorName}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error refreshing tab colors: {ex.Message}");
            }
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
                // var validEffects = Enum.GetNames(typeof(LaunchEffect)).ToHashSet();
                var validEffects = Enum.GetNames(typeof(LaunchEffectsManager.LaunchEffect)).ToHashSet();
                for (int i = 0; i < FenceDataManager.FenceData.Count; i++)
                {
                    dynamic fence = FenceDataManager.FenceData[i];
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
                        var validSizes = new[] { "Tiny", "Small", "Medium", "Large", "Huge" };
                        if (string.IsNullOrEmpty(iconSize) || !validSizes.Contains(iconSize))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid IconSize '{iconSize}' in {fence.Title}, resetting to \"Medium\"");
                            fenceDict["IconSize"] = "Medium";
                            jsonModified = true;
                        }
                    }


                    // TABS FEATURE: Migration for TabsEnabled property
                    if (!fenceDict.ContainsKey("TabsEnabled"))
                    {
                        fenceDict["TabsEnabled"] = "false"; // String for JSON compatibility
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added TabsEnabled=\"false\" to {fence.Title}");
                    }
                    else
                    {
                        // Validate existing TabsEnabled value
                        string tabsEnabled = fenceDict["TabsEnabled"]?.ToString();
                        if (tabsEnabled != "true" && tabsEnabled != "false")
                        {
                            fenceDict["TabsEnabled"] = "false";
                            jsonModified = true;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid TabsEnabled '{tabsEnabled}' in {fence.Title}, reset to \"false\"");
                        }
                    }

                    // TABS FEATURE: Migration for CurrentTab property
                    if (!fenceDict.ContainsKey("CurrentTab"))
                    {
                        fenceDict["CurrentTab"] = 0; // Integer default
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added CurrentTab=0 to {fence.Title}");
                    }
                    else
                    {
                        // Validate existing CurrentTab value
                        if (!int.TryParse(fenceDict["CurrentTab"]?.ToString(), out int currentTab) || currentTab < 0)
                        {
                            fenceDict["CurrentTab"] = 0;
                            jsonModified = true;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Invalid CurrentTab in {fence.Title}, reset to 0");
                        }
                    }
                    // TABS FEATURE: Migration for Tabs array property with intelligent content management
                    if (!fenceDict.ContainsKey("Tabs"))
                    {
                        fenceDict["Tabs"] = new JArray(); // Empty array default
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added empty Tabs array to {fence.Title}");
                    }

                    // TABS FEATURE: Intelligent migration based on current state
                    bool currentTabsEnabled = fenceDict["TabsEnabled"]?.ToString().ToLower() == "true";
                    var mainItems = fenceDict["Items"] as JArray ?? new JArray();
                    var tabs = fenceDict["Tabs"] as JArray ?? new JArray();

                    // Scenario 1: Tabs enabled but no tab structure exists yet
                    // This happens when user first enables tabs - migrate main Items to Tab 0
                    if (currentTabsEnabled && tabs.Count == 0 && mainItems.Count > 0)
                    {
                        // Create Tab 0 with existing fence content
                        var tab0 = new JObject();
                        tab0["TabName"] = fence.Title?.ToString() ?? "Main"; // Use fence title as default tab name
                        tab0["Items"] = mainItems; // Move existing items to Tab 0
                        tabs.Add(tab0);

                        // Clear main Items since they're now in Tab 0
                        fenceDict["Items"] = new JArray();
                        fenceDict["Tabs"] = tabs;
                        jsonModified = true;
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                            $"First-time tabs enable: Migrated {mainItems.Count} items to Tab 0 for fence '{fence.Title}'");
                    }

                    // Scenario 2: Tabs disabled but tab structure exists
                    // This happens when user disables tabs - migrate Tab 0 back to main Items
                    else if (!currentTabsEnabled && tabs.Count > 0)
                    {
                        var tab0 = tabs[0] as JObject;
                        if (tab0 != null)
                        {
                            var tab0Items = tab0["Items"] as JArray ?? new JArray();

                            // ENHANCED: Merge Tab0 items with existing main Items to prevent loss
                            if (tab0Items.Count > 0 || mainItems.Count > 0)
                            {
                                // Create a merged items array with no duplicates
                                var mergedItems = new JArray();
                                var existingFilenames = new HashSet<string>();

                                // Add Tab0 items first (they take priority)
                                foreach (var item in tab0Items)
                                {
                                    string filename = item["Filename"]?.ToString();
                                    if (!string.IsNullOrEmpty(filename) && !existingFilenames.Contains(filename))
                                    {
                                        mergedItems.Add(item);
                                        existingFilenames.Add(filename);
                                    }
                                }

                                // Add any main Items that aren't already in Tab0
                                foreach (var item in mainItems)
                                {
                                    string filename = item["Filename"]?.ToString();
                                    if (!string.IsNullOrEmpty(filename) && !existingFilenames.Contains(filename))
                                    {
                                        mergedItems.Add(item);
                                    }
                                }

                                // Update DisplayOrder for merged items
                                for (int j = 0; j < mergedItems.Count; j++)
                                {
                                    mergedItems[j]["DisplayOrder"] = j;
                                }

                                fenceDict["Items"] = mergedItems;
                                jsonModified = true;

                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                    $"Tabs disabled: Merged {tab0Items.Count} Tab0 items + {mainItems.Count} main Items = {mergedItems.Count} total items for fence '{fence.Title}'");
                            }
                        }
                    }

                    // Scenario 3: Validate existing tab structure (if tabs enabled and structure exists)
                    else if (currentTabsEnabled && tabs.Count > 0)
                    {
                        try
                        {
                            // ENHANCED: Check if main Items has content that needs to be merged into Tab0
                            if (mainItems.Count > 0)
                            {
                                var tab0 = tabs[0] as JObject;
                                if (tab0 != null)
                                {
                                    var tab0Items = tab0["Items"] as JArray ?? new JArray();

                                    // Merge main Items into Tab0 (avoiding duplicates)
                                    var existingFilenames = new HashSet<string>();

                                    // Track existing Tab0 items
                                    foreach (var item in tab0Items)
                                    {
                                        string filename = item["Filename"]?.ToString();
                                        if (!string.IsNullOrEmpty(filename))
                                        {
                                            existingFilenames.Add(filename);
                                        }
                                    }

                                    // Add new items from main Items to Tab0
                                    bool itemsAdded = false;
                                    foreach (var item in mainItems)
                                    {
                                        string filename = item["Filename"]?.ToString();
                                        if (!string.IsNullOrEmpty(filename) && !existingFilenames.Add(filename))
                                        {
                                            continue; // Skip duplicate
                                        }

                                        tab0Items.Add(item);
                                        itemsAdded = true;
                                    }

                                    if (itemsAdded)
                                    {
                                        // Update DisplayOrder for all items in Tab0
                                        for (int k = 0; k < tab0Items.Count; k++)
                                        {
                                            tab0Items[k]["DisplayOrder"] = k;
                                        }

                                        tab0["Items"] = tab0Items;
                                        fenceDict["Items"] = new JArray(); // Clear main Items after merging
                                        jsonModified = true;

                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                            $"Merged {mainItems.Count} main Items into Tab0 for fence '{fence.Title}' (tabs were re-enabled)");
                                    }
                                }
                            }

                            // Validate each tab object structure
                            for (int tabIndex = 0; tabIndex < tabs.Count; tabIndex++)
                            {
                                var tab = tabs[tabIndex] as JObject;
                                if (tab == null)
                                {
                                    // Remove invalid tab entry
                                    tabs.RemoveAt(tabIndex);
                                    tabIndex--; // Adjust index after removal
                                    jsonModified = true;
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Removed invalid tab entry at index {tabIndex} in {fence.Title}");
                                    continue;
                                }

                                // Ensure TabName exists
                                if (tab["TabName"] == null || string.IsNullOrWhiteSpace(tab["TabName"].ToString()))
                                {
                                    tab["TabName"] = tabIndex == 0 ? fence.Title?.ToString() ?? "Main" : $"Tab {tabIndex + 1}";
                                    jsonModified = true;
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added default TabName to tab {tabIndex} in {fence.Title}");
                                }

                                // Ensure Items array exists
                                if (tab["Items"] == null || !(tab["Items"] is JArray))
                                {
                                    tab["Items"] = new JArray();
                                    jsonModified = true;
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added empty Items array to tab {tabIndex} in {fence.Title}");
                                }

                                // Migrate tab items (existing code for item validation)
                                var tabItems = tab["Items"] as JArray ?? new JArray();
                                bool tabItemsModified = false;
                                int tabOrderCounter = 0;

                                foreach (var tabItem in tabItems)
                                {
                                    var tabItemDict = tabItem as JObject;
                                    if (tabItemDict != null)
                                    {
                                        // Add IsFolder if missing
                                        if (!tabItemDict.ContainsKey("IsFolder"))
                                        {
                                            string path = tabItemDict["Filename"]?.ToString();
                                            tabItemDict["IsFolder"] = !string.IsNullOrEmpty(path) && System.IO.Directory.Exists(path);
                                            tabItemsModified = true;
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added IsFolder to tab item {path} in {fence.Title}");
                                        }

                                        // Add DisplayOrder if missing
                                        if (!tabItemDict.ContainsKey("DisplayOrder"))
                                        {
                                            tabItemDict["DisplayOrder"] = tabOrderCounter;
                                            tabItemsModified = true;
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added DisplayOrder={tabOrderCounter} to tab item in {fence.Title}");
                                        }

                                        // Add IsLink if missing
                                        if (!tabItemDict.ContainsKey("IsLink"))
                                        {
                                            tabItemDict["IsLink"] = false;
                                            tabItemsModified = true;
                                        }

                                        // Add IsNetwork if missing
                                        if (!tabItemDict.ContainsKey("IsNetwork"))
                                        {
                                            string filename = tabItemDict["Filename"]?.ToString() ?? "";
                                            tabItemDict["IsNetwork"] = IsNetworkPath(filename);
                                            tabItemsModified = true;
                                        }

                                        // Add DisplayName if missing
                                        if (!tabItemDict.ContainsKey("DisplayName") && tabItemDict.ContainsKey("Filename"))
                                        {
                                            string filename = tabItemDict["Filename"]?.ToString();
                                            if (!string.IsNullOrEmpty(filename))
                                            {
                                                tabItemDict["DisplayName"] = System.IO.Path.GetFileNameWithoutExtension(filename);
                                                tabItemsModified = true;
                                            }
                                        }
                                    }
                                    tabOrderCounter++;
                                }

                                if (tabItemsModified)
                                {
                                    tab["Items"] = tabItems;
                                    jsonModified = true;
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Updated tab items in tab {tabIndex} of {fence.Title}");
                                }
                            }

                            // Validate CurrentTab is within bounds
                            int currentTabValue = Convert.ToInt32(fenceDict["CurrentTab"]);
                            if (currentTabValue >= tabs.Count && tabs.Count > 0)
                            {
                                fenceDict["CurrentTab"] = 0;
                                jsonModified = true;
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"CurrentTab out of bounds in {fence.Title}, reset to 0");
                            }

                            fenceDict["Tabs"] = tabs;
                        }
                        catch (Exception tabEx)
                        {
                            // If any error occurs, disable tabs and use main Items
                            fenceDict["TabsEnabled"] = "false";
                            fenceDict["CurrentTab"] = 0;
                            fenceDict["Tabs"] = new JArray();
                            jsonModified = true;
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Error validating tabs in {fence.Title}: {tabEx.Message}, disabled tabs");
                        }
                    }

                    // ENHANCED: Ensure Tab0 synchronization after any migration
                    if (jsonModified && currentTabsEnabled && tabs.Count > 0)
                    {
                        string fenceId = fence.Id?.ToString();
                        if (!string.IsNullOrEmpty(fenceId))
                        {
                            // Use a small delay to ensure data is saved before sync
                            Task.Delay(50).ContinueWith(_ =>
                            {
                                SynchronizeTab0Content(fenceId, "tab0", "full");
                            });
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


                    // Update the original fence in FenceDataManager.FenceData with the modified dictionary
                    FenceDataManager.FenceData[i] = JObject.FromObject(fenceDict);
                }

                if (jsonModified)
                {
                    FenceDataManager.SaveFenceData();
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Migrated fences.json with updated fields");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, "No migration needed for fences.json");
                }
                // ENHANCED: Perform Tab0 synchronization check after migration
foreach (var fence in FenceDataManager.FenceData)
{
    bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
    if (tabsEnabled)
    {
        string fenceId = fence.Id?.ToString();
        if (!string.IsNullOrEmpty(fenceId))
        {
            SynchronizeTab0Content(fenceId, "tab0", "startup");
        }
    }
}
LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Tab0 synchronization check completed");

            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error migrating fences.json: {ex.Message}\nStackTrace: {ex.StackTrace}");
            }
        }

        public static void LoadAndCreateFences(TargetChecker targetChecker)


        {

            // Get current program version from assembly
            string currentVersion = Assembly.GetEntryAssembly()?.GetName().Version?.ToString() ?? "0.0.0"; // This is current version for registry tracking


            //  string currentVersion = "2.5.2.111"; // This is current version for registry tracking
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
                TintValue = SettingsManager.TintValue,
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

        private static void InitializeDefaultFence()
        {
            string defaultJson = "[{\"Id\":\"" + Guid.NewGuid().ToString() + "\",\"Title\":\"New Fence\",\"X\":20,\"Y\":20,\"Width\":230,\"Height\":130,\"ItemsType\":\"Data\",\"Items\":[],\"IsLocked\":\"false\",\"IsHidden\":\"false\",\"CustomColor\":null,\"CustomLaunchEffect\":null,\"IsRolled\":\"false\",\"UnrolledHeight\":130,\"TabsEnabled\":\"false\",\"CurrentTab\":0,\"Tabs\":[]}]";
            System.IO.File.WriteAllText(FenceDataManager.JsonFilePath, defaultJson);
            FenceDataManager.FenceData = JsonConvert.DeserializeObject<List<dynamic>>(defaultJson);
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
            //   heart.ContextMenu = BuildHeartContextMenu(fence);
            heart.ContextMenu = BuildHeartContextMenu(fence, false); // Normal menu by default

            //// Handle left-click to open heart context menu
            //heart.MouseLeftButtonDown += (s, e) =>
            //{
            //    if (e.ChangedButton == MouseButton.Left && heart.ContextMenu != null)
            //    {
            //        heart.ContextMenu.IsOpen = true;
            //        e.Handled = true;
            //    }
            //};
            // Handle left-click to open heart context menu
            heart.MouseLeftButtonDown += (s, e) =>
            {
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

                    dynamic currentFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                    if (currentFence == null)
                    {
                        //     DebugLog("ERROR", fenceId, "Fence not found in FenceDataManager.FenceData for Ctrl+Click");
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

                        int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                        if (fenceIndex >= 0)
                        {
                            FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                        }
                        FenceDataManager.SaveFenceData();

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

                        int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                        if (fenceIndex >= 0)
                        {
                            FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                        }
                        FenceDataManager.SaveFenceData();

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







            // Clear Dead Shortcuts menu item (only for Data fences)
            if (fence.ItemsType?.ToString() == "Data")
            {
                CnMnFenceManager.Items.Add(new Separator());
                MenuItem miClearDeadShortcuts = new MenuItem { Header = "Clear Dead Shortcuts" };
                CnMnFenceManager.Items.Add(miClearDeadShortcuts);

                miClearDeadShortcuts.Click += (s, e) =>
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                        $"Clear Dead Shortcuts clicked for fence '{fence.Title}'");

                    try
                    {
                        int removedCount = FilePathUtilities.ClearDeadShortcutsFromFence(fence);

                        if (removedCount > 0)
                        {
                            // Save the updated fence data
                            FenceDataManager.SaveFenceData();

                            //                        // Refresh the fence display to show changes
                            //                        // For tabbed fences, refresh current tab; for regular fences, refresh all
                            //                        bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                            //                        if (tabsEnabled)
                            //                        {
                            //                            int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");
                            //                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                            //$"About to refresh fence - TabsEnabled: {tabsEnabled}, Method: {(tabsEnabled ? "RefreshFenceContentSimple" : "RefreshFenceContent")}");
                            //                            RefreshFenceContentSimple(win, fence, currentTab);
                            //                        }
                            //                        else
                            //                        {
                            //                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceUpdate,
                            //$"About to refresh fence - TabsEnabled: {tabsEnabled}, Method: {(tabsEnabled ? "RefreshFenceContentSimple" : "RefreshFenceContent")}");
                            //                            RefreshFenceContent(win, fence);
                            //                        }


                            // Refresh the fence display to show changes (using same approach as CustomizeFenceForm)
                            RefreshFenceUsingFormApproach(win, fence);

                            // Show completion message
                            string message = removedCount == 1 ?
                                "Removed 1 dead shortcut." :
                                $"Removed {removedCount} dead shortcuts.";
                         //   MessageBoxesManager.ShowOKOnlyMessageBoxForm(message, "Clear Dead Shortcuts");
                        }
                        else
                        {
                          //  MessageBoxesManager.ShowOKOnlyMessageBoxForm("No dead shortcuts found.", "Clear Dead Shortcuts");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                            $"Error in Clear Dead Shortcuts: {ex.Message}");
                       // MessageBoxesManager.ShowOKOnlyMessageBoxForm("An error occurred while clearing dead shortcuts.", "Error");
                    }
                };
            }











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
            //CnMnFenceManager.Items.Add(miCustomize); // Add Customize submenu


            // Add Paste Item functionality - CopyPasteManager integration  
            MenuItem miPasteItem = new MenuItem { Header = "Paste Item" };
            miPasteItem.Click += (s, e) => CopyPasteManager.PasteItem(fence);
            CnMnFenceManager.Items.Add(miPasteItem);

            // Add dynamic menu state updates - make Paste Item enabled state dynamic
            CnMnFenceManager.Opened += (contextSender, contextArgs) =>
            {
                // Update Paste Item enabled state dynamically when menu opens
                bool hasCopiedItem = CopyPasteManager.HasCopiedItem();
                miPasteItem.IsEnabled = hasCopiedItem;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Updated fence context menu - Paste available: {hasCopiedItem}");
            };
            CnMnFenceManager.Items.Add(new Separator());

            //// New menu item for CustomizeFenceForm 
            //MenuItem miNewCustomize = new MenuItem { Header = "Customize..." };
            //miNewCustomize.Click += (s, e) =>
            //{
            //    try
            //    {
            //        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Opening CustomizeFenceForm for fence '{fence.Title}'");
            //        using (var customizeForm = new CustomizeFenceForm(fence))
            //        {
            //            customizeForm.ShowDialog();
            //            if (customizeForm.DialogResult)
            //            {
            //                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"User saved changes in CustomizeFenceForm for fence '{fence.Title}'");
            //            }
            //            else
            //            {
            //                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"User cancelled CustomizeFenceForm for fence '{fence.Title}'");
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error opening CustomizeFenceForm: {ex.Message}");
            //        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error opening customize form: {ex.Message}", "Form Error");
            //    }
            //};
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
                 
                        if (Math.Abs(newHeight - 26) > 5) // Only if height is significantly different from rolled-up height
                        {
                            double heightDifference = Math.Abs(newHeight - oldUnrolledHeight);
                            //       DebugLog("LOGIC", fenceId, $"Height difference from old UnrolledHeight: {heightDifference:F1}");

                            currentFence.UnrolledHeight = newHeight;
                            //      DebugLog("UPDATE", fenceId, $"UPDATED UnrolledHeight from {oldUnrolledHeight:F1} to {newHeight:F1}");
                        }
                        else
                        {
                          
                        }
                    }
                    else
                    {
                 
                    }

                    // Save to JSON
              
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

            // TABS FEATURE: Add TabStrip conditionally for tabbed fences
            StackPanel tabStrip = null;
            bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

            if (tabsEnabled)
            {
                tabStrip = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Height = 20, // Reduced height to fit the smaller buttons elegantly
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(20, 0, 0, 0)), // Even more subtle background
                    Margin = new Thickness(0, 1, 0, 0), // Tighter margins
                    VerticalAlignment = VerticalAlignment.Bottom // Align to bottom for better visual connection
                };

                DockPanel.SetDock(tabStrip, Dock.Top);
                dp.Children.Add(tabStrip);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Added TabStrip for tabbed fence '{fence.Title}'");
                // TABS FEATURE: Create tab buttons from Tabs array
                try
                {
                    var tabs = fence.Tabs as JArray ?? new JArray();
                    int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");

                    // Ensure CurrentTab is within bounds
                    if (currentTab >= tabs.Count)
                    {
                        currentTab = 0;
                    }

                    for (int i = 0; i < tabs.Count; i++)
                    {
                        var tab = tabs[i] as JObject;
                        if (tab == null) continue;

                        string tabName = tab["TabName"]?.ToString() ?? $"Tab {i + 1}";
                        bool isActiveTab = (i == currentTab);

                        Button tabButton = new Button
                        {
                            Content = tabName,
                            Tag = i, // Store tab index for identification
                            Height = 18, // Reduced to 2/3 of original (26 * 2/3 ≈ 17-18)
                            MinWidth = 50,
                            Margin = new Thickness(1, 0, 1, 0), // Tighter margins for elegance
                            Padding = new Thickness(10, 2, 10, 2), // Slightly more horizontal padding
                            FontSize = 10, // Slightly smaller font for elegance
                            FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                            FontWeight = FontWeights.Medium, // Slightly bolder for better readability
                            BorderThickness = new Thickness(1, 1, 1, 0), // No bottom border to blend with content
                            Cursor = Cursors.Hand,
                            // Custom template to achieve rounded top corners only
                            Template = CreateTabButtonTemplate()
                        };

                        // TABS FEATURE: Get fence color scheme for theming
                        string fenceColorName = fence.CustomColor?.ToString() ?? SettingsManager.SelectedColor;
                        var colorScheme = Utility.GenerateTabColorScheme(fenceColorName);
                        // Apply themed styling with better visual distinction
                        if (isActiveTab)
                        {
                            // Active tab styling - more prominent
                            tabButton.Background = new SolidColorBrush(colorScheme.activeTab);
                            tabButton.Foreground = System.Windows.Media.Brushes.White;
                            tabButton.BorderBrush = new SolidColorBrush(colorScheme.borderColor);
                            tabButton.FontWeight = FontWeights.Bold; // Make active tab bold
                            tabButton.Opacity = 1.0; // Full opacity for active
                        }
                        else
                        {
                            // Inactive tab styling - more muted
                            tabButton.Background = new SolidColorBrush(colorScheme.inactiveTab);
                            tabButton.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 96, 96, 96)); // Lighter gray
                            tabButton.BorderBrush = new SolidColorBrush(colorScheme.borderColor);
                            tabButton.FontWeight = FontWeights.Normal; // Normal weight for inactive
                            tabButton.Opacity = 0.7; // Reduced opacity for inactive
                        }
                        // Add hover effects for inactive tabs using themed colors
                        if (!isActiveTab)
                        {
                            tabButton.MouseEnter += (s, e) =>
                            {
                                var btn = s as Button;
                                btn.Background = new SolidColorBrush(colorScheme.hoverTab);
                            };
                            tabButton.MouseLeave += (s, e) =>
                            {
                                var btn = s as Button;
                                btn.Background = new SolidColorBrush(colorScheme.inactiveTab);
                            };
                        }

                        // TABS FEATURE: Add click handler for tab switching with better reliability
                        tabButton.Click += (s, e) =>
                        {
                            try
                            {
                                var clickedButton = s as Button;
                                int newTabIndex = Convert.ToInt32(clickedButton.Tag);

                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                    $"Tab button clicked: index {newTabIndex} for fence '{fence.Title}'");

                                // Use the fence object directly instead of searching by window
                                SwitchTabByFence(fence, newTabIndex, win);
                            }
                            catch (Exception ex)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                                    $"Error handling tab click: {ex.Message}");
                            }
                        };
                        // TABS FEATURE: Add right-click context menu for tab management
                        ContextMenu tabContextMenu = new ContextMenu();

                        MenuItem miAddTab = new MenuItem { Header = "Add New Tab" };
                        MenuItem miRenameTab = new MenuItem { Header = "Rename Tab" };
                        MenuItem miDeleteTab = new MenuItem { Header = "Delete Tab" };
                        Separator miSeparator1 = new Separator();
                        MenuItem miMoveLeft = new MenuItem { Header = "Move Left", IsEnabled = false }; 
                        MenuItem miMoveRight = new MenuItem { Header = "Move Right", IsEnabled = false }; 

                        tabContextMenu.Items.Add(miAddTab);
                        tabContextMenu.Items.Add(miSeparator1);
                        tabContextMenu.Items.Add(miRenameTab);
                        tabContextMenu.Items.Add(miDeleteTab);
                        tabContextMenu.Items.Add(new Separator());
                        tabContextMenu.Items.Add(miMoveLeft);
                        tabContextMenu.Items.Add(miMoveRight);

                        // Capture current context for event handlers
                        int capturedTabIndex = i;

                        miAddTab.Click += (s, e) => AddNewTab(fence, win);
                        miRenameTab.Click += (s, e) => RenameTab(fence, capturedTabIndex, win);
                        miDeleteTab.Click += (s, e) => DeleteTab(fence, capturedTabIndex, win);

                        tabButton.ContextMenu = tabContextMenu;
                        tabStrip.Children.Add(tabButton);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Added tab button '{tabName}' (index {i}, active: {isActiveTab})");
                    }

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation, $"Created {tabs.Count} tab buttons for fence '{fence.Title}'");
                    // TABS FEATURE: Add [+] button for adding new tabs
                    Button addTabButton = new Button
                    {
                        Content = "+",
                        Tag = "ADD_TAB", // Special tag to identify add button
                        Height = 18,
                        Width = 25, // Slightly wider for the + symbol
                        Margin = new Thickness(3, 0, 1, 0), // Bit more space before +
                        Padding = new Thickness(0),
                        FontSize = 12,
                        FontWeight = FontWeights.Bold,
                        FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                        BorderThickness = new Thickness(1),
                        Cursor = Cursors.Hand,
                        Template = CreateTabButtonTemplate(),
                        ToolTip = "Add new tab"
                    };

                    // Style the [+] button with muted colors using existing color scheme
                    string addButtonFenceColor = fence.CustomColor?.ToString() ?? SettingsManager.SelectedColor;
                    var addButtonColorScheme = Utility.GenerateTabColorScheme(addButtonFenceColor);
                    addTabButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, addButtonColorScheme.inactiveTab.R, addButtonColorScheme.inactiveTab.G, addButtonColorScheme.inactiveTab.B)); // Semi-transparent
                    addTabButton.Foreground = new SolidColorBrush(addButtonColorScheme.borderColor);
                    addTabButton.BorderBrush = new SolidColorBrush(addButtonColorScheme.borderColor);
                    addTabButton.Opacity = 0.8;

                    // Add hover effects
                    addTabButton.MouseEnter += (s, e) =>
                    {
                        addTabButton.Background = new SolidColorBrush(addButtonColorScheme.hoverTab);
                        addTabButton.Opacity = 1.0;
                    };
                    addTabButton.MouseLeave += (s, e) =>
                    {
                        addTabButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, addButtonColorScheme.inactiveTab.R, addButtonColorScheme.inactiveTab.G, addButtonColorScheme.inactiveTab.B));
                        addTabButton.Opacity = 0.8;
                    };
                    // Add click handler for new tab creation
                    addTabButton.Click += (s, e) =>
                    {
                        try
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Add tab button clicked for fence '{fence.Title}'");
                            AddNewTab(fence, win);
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error adding new tab: {ex.Message}");
                        }
                    };

                    tabStrip.Children.Add(addTabButton);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Added [+] button for fence '{fence.Title}'");
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation, $"Error creating tab buttons for fence '{fence.Title}': {ex.Message}");
                }

            }
            else
            {



                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Skipped TabStrip for non-tabbed fence '{fence.Title}'");
            }




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
                    FenceDataManager.SaveFenceData();
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
                FenceDataManager.SaveFenceData();
                win.ShowActivated = false;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Exited edit mode via click, final title for fence: {fence.Title}");
            };
            WrapPanel wpcont = new WrapPanel();
            ScrollViewer wpcontscr = new ScrollViewer
            {
                Content = wpcont,
                VerticalScrollBarVisibility = SettingsManager.DisableFenceScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto,
              //  HorizontalScrollBarVisibility = SettingsManager.DisableFenceScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto
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
                    // TABS FEATURE: Determine which items to load based on tabs configuration
                    JArray items = null;
                    bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

                    if (tabsEnabled)
                    {
                        try
                        {
                            var tabs = fence.Tabs as JArray ?? new JArray();
                            int currentTab = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");

                            // Validate CurrentTab is within bounds
                            if (currentTab >= 0 && currentTab < tabs.Count)
                            {
                                var activeTab = tabs[currentTab] as JObject;
                                if (activeTab != null)
                                {
                                    items = activeTab["Items"] as JArray ?? new JArray();
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                        $"Loading tab '{activeTab["TabName"]}' (index {currentTab}) for fence '{fence.Title}'");
                                }
                                else
                                {
                                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                                        $"Active tab {currentTab} is invalid, falling back to main Items for fence '{fence.Title}'");
                                    items = fence.Items as JArray ?? new JArray();
                                }
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                                    $"CurrentTab {currentTab} out of bounds (0-{tabs.Count - 1}), falling back to main Items for fence '{fence.Title}'");
                                items = fence.Items as JArray ?? new JArray();
                            }
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                $"Error loading tab items for fence '{fence.Title}': {ex.Message}, falling back to main Items");
                            items = fence.Items as JArray ?? new JArray();
                        }
                    }
                    else
                    {
                        // Tabs disabled - use main Items array (existing behavior)
                        items = fence.Items as JArray ?? new JArray();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                            $"Tabs disabled, loading main Items for fence '{fence.Title}'");
                    }

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

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                            $"Loading {sortedItems.Count} items in display order for fence '{fence.Title}' {(tabsEnabled ? "(from active tab)" : "(from main Items)")}");

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

                                // Add Copy Item functionality - CopyPasteManager integration
                                MenuItem miCopyItem = new MenuItem { Header = "Copy Item" };

                                mn.Items.Add(miEdit);
                                mn.Items.Add(miMove);
                                mn.Items.Add(miRemove);
                                mn.Items.Add(new Separator());
                                mn.Items.Add(miCopyItem);
                                mn.Items.Add(new Separator());
                                mn.Items.Add(miRunAsAdmin);
                                mn.Items.Add(new Separator());
                                mn.Items.Add(miCopyPath);
                                mn.Items.Add(miFindTarget);
                                sp.ContextMenu = mn;

                                // Add dynamic menu state updates - make Copy Item enabled and handle CTRL+Right Click
                                MenuItem miSendToDesktop = null; // Declare for conditional addition
                                mn.Opened += (contextSender, contextArgs) =>
                                {
                                    // Update Copy Item enabled state dynamically when menu opens
                                    miCopyItem.IsEnabled = true; // Copy should always be enabled for existing items

                                    // Check if CTRL is pressed for "Send to Desktop" option
                                    bool isCtrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

                                    // Remove existing Send to Desktop if present
                                    if (miSendToDesktop != null && mn.Items.Contains(miSendToDesktop))
                                    {
                                        mn.Items.Remove(miSendToDesktop);
                                    }

                                    // Add Send to Desktop if CTRL is pressed
                                    if (isCtrlPressed)
                                    {
                                        miSendToDesktop = new MenuItem { Header = "Send to Desktop" };
                                        miSendToDesktop.Click += (s, e) => CopyPasteManager.SendToDesktop(icon);

                                        // Find Copy Item position and insert after it
                                        int copyItemIndex = -1;
                                        for (int i = 0; i < mn.Items.Count; i++)
                                        {
                                            if (mn.Items[i] is MenuItem mi && mi.Header.ToString() == "Copy Item")
                                            {
                                                copyItemIndex = i;
                                                break;
                                            }
                                        }

                                        if (copyItemIndex >= 0)
                                        {
                                            mn.Items.Insert(copyItemIndex + 1, miSendToDesktop);
                                        }
                                    }

                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                        $"Updated context menu states - Copy available, CTRL pressed: {isCtrlPressed}");
                                };


                                miRunAsAdmin.IsEnabled = Utility.IsExecutableFile(filePath);

                                // Add Copy Item event handler - CopyPasteManager integration
                                miCopyItem.Click += (sender, e) => CopyPasteManager.CopyItem(icon, fence);

                                miMove.Click += (sender, e) => ItemMoveDialog.ShowMoveDialog(icon, fence, win.Dispatcher);
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
                                                FenceDataManager.SaveFenceData();

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
                                    // Use LogManager instead of direct file writing
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                        $"Copied target folder path to clipboard: {folderPath}");
                                };

                                miCopyFullPath.Click += (sender, e) =>
                                {
                                    string targetPath = Utility.GetShortcutTarget(filePath);
                                    Clipboard.SetText(targetPath);
                                    // Use LogManager instead of direct file writing
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                        $"Copied target full path to clipboard: {targetPath}");
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

                                    //LAUNCH ADD - CRITICAL FIX: Set working directory for administrator execution
                                    if (System.IO.File.Exists(targetPath))
                                    {
                                        string workingDir = System.IO.Path.GetDirectoryName(targetPath);
                                        if (!string.IsNullOrEmpty(workingDir) && System.IO.Directory.Exists(workingDir))
                                        {
                                            psi.WorkingDirectory = workingDir;
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                                $"Set admin working directory: {workingDir}");
                                        }
                                    }


                                    try
                                    {
                                        Process.Start(psi);
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Successfully launched {targetPath} as admin");
                                    }
                                    catch (Exception ex)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to launch {targetPath} as admin: {ex.Message}");
                                                   MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error running as admin: {ex.Message}", "Error");

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
                                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Invalid file or directory: {droppedFile}", "Error");
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
                                bool isWebLink = CoreUtilities.IsWebLinkShortcut(droppedFile); // Use our new helper method

                                string targetPath;
                                bool isFolder = false;
                                string webUrl = null;

                                if (isWebLink)
                                {
                                    // Handle web link shortcuts
                                    webUrl = CoreUtilities.ExtractWebUrlFromFile(droppedFile);
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
                                        targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(droppedFile);

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

                                                        // Set icon to folder icon with full path
                                                        unicodeShortcut.IconLocation = $"{Environment.GetFolderPath(Environment.SpecialFolder.System)}\\shell32.dll,3";

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
                                    if (FilePathUtilities.ValidateUnicodeFolderPath(droppedFile, out validatedPath))
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


                                // TABS FEATURE: Get fresh fence data FIRST to avoid stale CurrentTab
                                string fenceId = fence.Id?.ToString();
                                var freshFence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                                JArray items = null;
                                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";
                                bool addedToTab0 = false; // Track if we're adding to Tab0 for sync

                                if (tabsEnabled && freshFence != null)
                                {
                                    // Use FRESH fence data for CurrentTab
                                    int currentTab = Convert.ToInt32(freshFence.CurrentTab?.ToString() ?? "0");
                                    var tabs = freshFence.Tabs as JArray ?? new JArray();

                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                        $"Adding dropped item to tab {currentTab} using fresh fence data");

                                    if (currentTab >= 0 && currentTab < tabs.Count)
                                    {
                                        var activeTab = tabs[currentTab] as JObject;
                                        if (activeTab != null)
                                        {
                                            items = activeTab["Items"] as JArray ?? new JArray();
                                            string tabName = activeTab["TabName"]?.ToString() ?? $"Tab {currentTab}";
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                                $"Using items from tab '{tabName}' with {items.Count} existing items");

                                            // Track if we're adding to Tab0 for synchronization
                                            addedToTab0 = (currentTab == 0);
                                        }
                                    }

                                    // Fallback to main Items if tab structure is invalid
                                    if (items == null)
                                    {
                                        items = freshFence.Items as JArray ?? new JArray();
                                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                                            $"Invalid tab structure, using main Items as fallback");
                                    }
                                }
                                else
                                {
                                    // Tabs disabled - use main Items
                                    items = (freshFence ?? fence).Items as JArray ?? new JArray();
                                }


                                // Add DisplayOrder as the next available order number
                                int nextDisplayOrder = items.Count;
                                newItemDict["DisplayOrder"] = nextDisplayOrder;
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                    $"Assigned DisplayOrder={nextDisplayOrder} to new dropped item {shortcutName}");

                                // Add item to the correct array
                                items.Add(JObject.FromObject(newItem));

                                // ENHANCED: Synchronize Tab0 content after adding item
                                if (!string.IsNullOrEmpty(fenceId))
                                {
                                    if (addedToTab0)
                                    {
                                        // Added to Tab0 - sync to main Items
                                        SynchronizeTab0Content(fenceId, "tab0", "add");
                                    }
                                    else if (!tabsEnabled)
                                    {
                                        // Added to main Items when tabs disabled - sync to Tab0 if it exists
                                        SynchronizeTab0Content(fenceId, "main", "add");
                                    }
                                }

                                // Update FenceDataManager.FenceData and save
                                if (freshFence != null)
                                {
                                    int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                                    if (fenceIndex >= 0)
                                    {
                                        FenceDataManager.FenceData[fenceIndex] = freshFence;
                                        FenceDataManager.SaveFenceData();
                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                            $"Successfully saved dropped item '{shortcutName}' to correct location");
                                    }
                                }

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

                                    // Add Copy Item functionality - CopyPasteManager integration
                                    MenuItem miCopyItem = new MenuItem { Header = "Copy Item" };

                                    mn.Items.Add(miEdit);
                                    mn.Items.Add(miMove);
                                    mn.Items.Add(miRemove);
                                    mn.Items.Add(new Separator());
                                    mn.Items.Add(miCopyItem);
                                    mn.Items.Add(new Separator());
                                    mn.Items.Add(miRunAsAdmin);
                                    mn.Items.Add(new Separator());
                                    mn.Items.Add(miCopyPath);
                                    mn.Items.Add(miFindTarget);
                                    sp.ContextMenu = mn;



                                    // Add dynamic menu state updates - make Copy Item enabled and handle CTRL+Right Click
                                    MenuItem miSendToDesktop = null; // Declare for conditional addition
                                    mn.Opened += (contextSender, contextArgs) =>
                                    {
                                        // Update Copy Item enabled state dynamically when menu opens
                                        miCopyItem.IsEnabled = true; // Copy should always be enabled for existing items

                                        // Check if CTRL is pressed for "Send to Desktop" option
                                        bool isCtrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

                                        // Remove existing Send to Desktop if present
                                        if (miSendToDesktop != null && mn.Items.Contains(miSendToDesktop))
                                        {
                                            mn.Items.Remove(miSendToDesktop);
                                        }

                                        // Add Send to Desktop if CTRL is pressed
                                        if (isCtrlPressed)
                                        {
                                            miSendToDesktop = new MenuItem { Header = "Send to Desktop" };
                                            miSendToDesktop.Click += (s, e) => CopyPasteManager.SendToDesktop(newItem);

                                            // Find Copy Item position and insert after it
                                            int copyItemIndex = -1;
                                            for (int i = 0; i < mn.Items.Count; i++)
                                            {
                                                if (mn.Items[i] is MenuItem mi && mi.Header.ToString() == "Copy Item")
                                                {
                                                    copyItemIndex = i;
                                                    break;
                                                }
                                            }

                                            if (copyItemIndex >= 0)
                                            {
                                                mn.Items.Insert(copyItemIndex + 1, miSendToDesktop);
                                            }
                                        }

                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                            $"Updated drag & drop context menu - Copy available, CTRL pressed: {isCtrlPressed}");
                                    };



                                    miRunAsAdmin.IsEnabled = Utility.IsExecutableFile(shortcutName);
                                    // Add Copy Item event handler - CopyPasteManager integration
                                    miCopyItem.Click += (sender, e) => CopyPasteManager.CopyItem(newItem, fence);

                                    miMove.Click += (sender, e) => ItemMoveDialog.ShowMoveDialog(newItem, fence, win.Dispatcher);
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
                                                    FenceDataManager.SaveFenceData();

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

                                    //miRunAsAdmin.Click += (sender, e) =>
                                    //{
                                    //    string targetPath = Utility.GetShortcutTarget(shortcutName);
                                    //    string runArguments = null;
                                    //    if (System.IO.Path.GetExtension(shortcutName).ToLower() == ".lnk")
                                    //    {
                                    //        WshShell shell = new WshShell();
                                    //        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutName);
                                    //        runArguments = shortcut.Arguments;
                                    //    }
                                    //    ProcessStartInfo psi = new ProcessStartInfo
                                    //    {
                                    //        FileName = targetPath,
                                    //        UseShellExecute = true,
                                    //        Verb = "runas"
                                    //    };
                                    //    if (!string.IsNullOrEmpty(runArguments))
                                    //    {
                                    //        psi.Arguments = runArguments;
                                    //        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Run as admin with arguments: {runArguments} for {shortcutName}");
                                    //    }
                                    //    try
                                    //    {
                                    //        Process.Start(psi);
                                    //        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Successfully launched {targetPath} as admin");
                                    //    }
                                    //    catch (Exception ex)
                                    //    {
                                    //        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to launch {targetPath} as admin: {ex.Message}");
                                    //        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error running as admin: {ex.Message}", "Error");
                                    //    }
                                    //};
                                    miRunAsAdmin.Click += (sender, e) =>
                                    {
                                        string targetPath = Utility.GetShortcutTarget(shortcutName);
                                        string runArguments = null;
                                        if (System.IO.Path.GetExtension(shortcutName).ToLower() == ".lnk")
                                        {
                                            WshShell shell = new WshShell();
                                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutName);
                                            runArguments = shortcut.Arguments;
                                        }

                                        string fileExtension = System.IO.Path.GetExtension(targetPath).ToLower();

                                        if (fileExtension == ".ps1")
                                        {
                                            try
                                            {
                                                // Use schtasks to run PowerShell script as admin
                                                string taskName = $"PSScript_{DateTime.Now:yyyyMMddHHmmss}";
                                                string psCommand = $"powershell.exe -ExecutionPolicy Bypass -File \\\"{targetPath}\\\"";
                                                if (!string.IsNullOrEmpty(runArguments))
                                                {
                                                    psCommand += $" {runArguments}";
                                                }

                                                // Create temporary scheduled task
                                                ProcessStartInfo createTask = new ProcessStartInfo
                                                {
                                                    FileName = "schtasks.exe",
                                                    Arguments = $"/create /tn \"{taskName}\" /tr \"{psCommand}\" /sc once /st 00:00 /rl highest /f",
                                                    UseShellExecute = true,
                                                    Verb = "runas",
                                                    CreateNoWindow = true
                                                };

                                                Process createProcess = Process.Start(createTask);
                                                createProcess.WaitForExit();

                                                if (createProcess.ExitCode == 0)
                                                {
                                                    // Run the task immediately
                                                    ProcessStartInfo runTask = new ProcessStartInfo
                                                    {
                                                        FileName = "schtasks.exe",
                                                        Arguments = $"/run /tn \"{taskName}\"",
                                                        UseShellExecute = false,
                                                        CreateNoWindow = true
                                                    };
                                                    Process.Start(runTask);

                                                    // Delete the task after a delay
                                                    System.Threading.Tasks.Task.Delay(2000).ContinueWith(t =>
                                                    {
                                                        try
                                                        {
                                                            ProcessStartInfo deleteTask = new ProcessStartInfo
                                                            {
                                                                FileName = "schtasks.exe",
                                                                Arguments = $"/delete /tn \"{taskName}\" /f",
                                                                UseShellExecute = false,
                                                                CreateNoWindow = true
                                                            };
                                                            Process.Start(deleteTask);
                                                        }
                                                        catch { /* Ignore cleanup errors */ }
                                                    });

                                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                                                        $"Successfully launched PS1 via scheduled task as admin: {targetPath}");
                                                }
                                                else
                                                {
                                                    throw new Exception("Failed to create scheduled task");
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                                                    $"Failed to launch PS1 via scheduled task: {ex.Message}");
                                                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error running PowerShell as admin: {ex.Message}", "Error");
                                            }
                                        }
                                        else
                                        {
                                            // Original code for non-PowerShell files
                                            ProcessStartInfo psi = new ProcessStartInfo
                                            {
                                                FileName = targetPath,
                                                UseShellExecute = true,
                                                Verb = "runas"
                                            };
                                            if (!string.IsNullOrEmpty(runArguments))
                                            {
                                                psi.Arguments = runArguments;
                                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Run as admin with arguments: {runArguments} for {shortcutName}");
                                            }
                                            try
                                            {
                                                Process.Start(psi);
                                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Successfully launched {targetPath} as admin");
                                            }
                                            catch (Exception ex)
                                            {
                                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to launch {targetPath} as admin: {ex.Message}");
                                                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error running as admin: {ex.Message}", "Error");
                                            }
                                        }
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
                                      MessageBoxesManager.ShowOKOnlyMessageBoxForm($"No destination folder defined for this Portal Fence. Please recreate the fence.", "Error");
                                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation, $"No Path defined for Portal Fence: {fence.Title}");
                                    continue;
                                }

                                if (!System.IO.Directory.Exists(destinationFolder))
                                {
                                     MessageBoxesManager.ShowOKOnlyMessageBoxForm($"The destination folder '{destinationFolder}' no longer exists. Please update the Portal Fence settings.", "Error");
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
                                     LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation, $"Copied directory to Portal Fence: {destinationPath}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                               MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to add {droppedFile}: {ex.Message}", "Error");
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation, $"Failed to add {droppedFile}: {ex.Message}");
                        }
                    }
                    FenceDataManager.SaveFenceData(); // Αποθηκεύουμε τις αλλαγές στο JSON
                }
                else if (e.Data.GetDataPresent(DataFormats.Text) ||
                         e.Data.GetDataPresent(DataFormats.Html) ||
                         e.Data.GetDataPresent("UniformResourceLocator") ||
                         e.Data.GetDataPresent("text/x-moz-url"))
                {
                    try
                    {
                        string droppedUrl = ExtractUrlFromDropData(e.Data);

                        if (!string.IsNullOrEmpty(droppedUrl) && IsValidWebUrl(droppedUrl))
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                $"Processing dropped URL: '{droppedUrl}'");

                            // Only process for Data fences
                            if (fence.ItemsType?.ToString() == "Data")
                            {
                                AddUrlShortcutToFence(droppedUrl, fence, wpcont);
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                                    "URL drops not supported for Portal fences");
                            }
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                                $"Invalid or empty URL from drop data: '{droppedUrl}'");
                        }
                    }
                    catch (Exception urlEx)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                            $"Error processing URL drop: {urlEx.Message}");
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
            win.Show();


            IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
            SnapManager.AddSnapping(win, fenceDict);
            // Apply custom color if present, otherwise use global
            string customColor = fence.CustomColor?.ToString();
            Utility.ApplyTintAndColorToFence(win, string.IsNullOrEmpty(customColor) ? SettingsManager.SelectedColor : customColor);
            targetChecker.Start();
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
                foreach (var fenceData in FenceDataManager.FenceData)
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
            foreach (var fenceData in FenceDataManager.FenceData)
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
                    targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);

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
                            // Use IconManager to extract custom icon
                            shortcutIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                            if (shortcutIcon != null)
                            {
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling, $"Extracted custom icon at index {iconIndex} from {iconPath} for {filePath}");
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Failed to extract custom icon at index {iconIndex} from {iconPath} for {filePath}");
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
                        effectiveTargetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                            $"Re-resolved target for AddIcon: {filePath} -> {effectiveTargetPath}");
                    }

                    // Check if target is missing first (before trying to extract icons)
                    bool targetExists = !string.IsNullOrEmpty(effectiveTargetPath) &&
                                       (System.IO.Directory.Exists(effectiveTargetPath) || System.IO.File.Exists(effectiveTargetPath));

                    if (!targetExists)
                    {
                        // Target is missing - use appropriate missing icon
                        shortcutIcon = isFolder
                            ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                            : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")} (target: '{effectiveTargetPath}')");
                    }
                    else if (!string.IsNullOrEmpty(effectiveTargetPath) && System.IO.Directory.Exists(effectiveTargetPath))
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
                    // Check if this is actually a folder that shows as a file (some Unicode folders)
                    if (System.IO.Directory.Exists(filePath))
                    {
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Using folder-White.png for folder detected as file: {filePath}");
                    }
                    else
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
                foreach (var fenceData in FenceDataManager.FenceData)
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
                foreach (var fenceData in FenceDataManager.FenceData)
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
                foreach (var fenceData in FenceDataManager.FenceData)
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

        public static void EditItem(dynamic icon, dynamic fence, NonActivatingWindow win)
        {
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            string displayName = iconDict.ContainsKey("DisplayName") ? (string)iconDict["DisplayName"] : System.IO.Path.GetFileNameWithoutExtension(filePath);
            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";

            if (!isShortcut)
            {
                //   MessageBox.Show("Edit is only available for shortcuts.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);

                MessageBoxesManager.ShowOKOnlyMessageBoxForm("Edit is only available for shortcuts.", "Info");
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
                        FenceDataManager.SaveFenceData();
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
                    wpcont = FenceUtilities.FindWrapPanel(win);
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
                                        // Use IconManager to extract custom icon
                                        shortcutIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                                        if (shortcutIcon != null)
                                        {
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Extracted custom icon at index {iconIndex} from {iconPath} for {filePath}");
                                        }
                                        else
                                        {
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Failed to extract custom icon at index {iconIndex} from {iconPath} for {filePath}");
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
                            IconManager.IconCache[filePath] = shortcutIcon;
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
                    foreach (var fence in FenceDataManager.FenceData)
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
                    targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
                    if (SettingsManager.EnableBackgroundValidationLogging)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.BackgroundValidation,
                            $"UpdateIcon Unicode-safe resolution: {filePath} -> {targetPath}");
                    }
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

                // Use proper folder existence check for folder shortcuts
                if (isFolder)
                {
                    bool folderExists = FilePathUtilities.DoesFolderExist(filePath, isFolder);
                    if (folderExists)
                    {
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                        if (SettingsManager.EnableBackgroundValidationLogging)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.BackgroundValidation,
                                $"Using folder-White.png for existing folder {filePath}");
                        }
                    }
                    else
                    {
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Using folder-WhiteX.png for missing folder {filePath}");
                    }
                }
                else if (!targetExists)
                {
                    // Use missing icon for non-folder items when target is missing
                    newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                        $"Using file-WhiteX.png for missing file {filePath}");
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
                                // Use IconManager to extract custom icon
                                newIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                                if (newIcon != null)
                                {
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Extracted custom icon at index {iconIndex} from {iconPath} for {filePath}");
                                }
                                else
                                {
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Failed to extract custom icon at index {iconIndex} from {iconPath} for {filePath}");
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

                    // CHECKPOINT: Step 3 Complete!

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
                                iconTargetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                                    $"UpdateIcon using Unicode-safe resolution: {filePath} -> {iconTargetPath}");
                            }

                            if (!string.IsNullOrEmpty(iconTargetPath) && System.IO.File.Exists(iconTargetPath))
                            {
                                try
                                {
                                    newIcon = System.Drawing.Icon.ExtractAssociatedIcon(iconTargetPath).ToImageSource();
                                    if (SettingsManager.EnableBackgroundValidationLogging)
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.BackgroundValidation,
                                            $"UpdateIcon extracted target icon for Unicode shortcut {filePath}: {iconTargetPath}");
                                    }
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
                    if (IconManager.IconCache.ContainsKey(filePath))
                    {
                        IconManager.IconCache[filePath] = newIcon;
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
                        resolvedPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(path);

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
                dynamic fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Title == win?.Title);
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
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Shortcut file not found: {path}", "Error");
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
                                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Invalid .url file: {path}", "Error");
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
                            string resolvedTarget = FilePathUtilities.GetShortcutTargetUnicodeSafe(path);

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
                                        //// This is our Unicode folder shortcut - don't pass explorer arguments to launch
                                        //workingDirectory = System.IO.Path.GetDirectoryName(targetPath) ?? "";
                                        //finalArguments = arguments ?? ""; // Use provided arguments, not explorer args

                                        // LAUNCH ADD -  FIXED: Unicode folder shortcut working directory logic
                                        string targetDir = System.IO.Path.GetDirectoryName(targetPath);
                                        if (!string.IsNullOrEmpty(targetDir) && System.IO.Directory.Exists(targetDir))
                                        {
                                            workingDirectory = targetDir;
                                        }
                                        else if (System.IO.Directory.Exists(targetPath))
                                        {
                                            // If targetPath itself is a directory, use it
                                            workingDirectory = targetPath;
                                        }
                                        else
                                        {
                                            workingDirectory = "";
                                        }
                                        finalArguments = arguments ?? ""; // Use provided arguments, not explorer args
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                            $"Unicode folder shortcut working directory: '{workingDirectory}'");





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
                                    //

                                    // LAUNCH ADD - ENHANCED: Fallback for working directory with comprehensive validation
                                    if (System.IO.Directory.Exists(targetPath))
                                    {
                                        workingDirectory = targetPath; // For folders, working directory is the folder itself
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                            $"Fallback: Set working directory to folder: {workingDirectory}");
                                    }
                                    else if (System.IO.File.Exists(targetPath))
                                    {
                                        string targetDir = System.IO.Path.GetDirectoryName(targetPath);
                                        if (!string.IsNullOrEmpty(targetDir) && System.IO.Directory.Exists(targetDir))
                                        {
                                            workingDirectory = targetDir;
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                                                $"Fallback: Set working directory from executable: {workingDirectory}");
                                        }
                                        else
                                        {
                                            workingDirectory = "";
                                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                                                $"Fallback: Cannot determine valid working directory for: {targetPath}");
                                        }
                                    }
                                    else
                                    {
                                        workingDirectory = "";
                                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                                            $"Fallback: Target does not exist, no working directory set: {targetPath}");
                                    }

                                    //// Fallback for working directory - handle folders properly
                                    //if (System.IO.Directory.Exists(targetPath))
                                    //{
                                    //    workingDirectory = targetPath; // For folders, working directory is the folder itself
                                    //}
                                    //else if (System.IO.File.Exists(targetPath))
                                    //{
                                    //    workingDirectory = System.IO.Path.GetDirectoryName(targetPath) ?? "";
                                    //}
                                    //else
                                    //{
                                    //    workingDirectory = "";
                                    //}

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
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error reading shortcut: {ex.Message}", "Error");
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
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Target '{targetPath}' was not found.", "Error");
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

                    //// Set working directory if available
                    //if (!string.IsNullOrEmpty(workingDirectory) && System.IO.Directory.Exists(workingDirectory))
                    //{
                    //    psi.WorkingDirectory = workingDirectory;
                    //    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Set working directory: {workingDirectory}");
                    //}

                    // LAUNCH ADD - ENHANCED: Set working directory with comprehensive fallback logic
                    if (!string.IsNullOrEmpty(workingDirectory) && System.IO.Directory.Exists(workingDirectory))
                    {
                        psi.WorkingDirectory = workingDirectory;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Set working directory: {workingDirectory}");
                    }
                    else if (!isUrl && !isSpecialPath && !isTargetFolder && System.IO.File.Exists(targetPath))
                    {
                        // CRITICAL FALLBACK: Always set working directory for executables when missing
                        string fallbackDir = System.IO.Path.GetDirectoryName(targetPath);
                        if (!string.IsNullOrEmpty(fallbackDir) && System.IO.Directory.Exists(fallbackDir))
                        {
                            psi.WorkingDirectory = fallbackDir;
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                $"FALLBACK: Set working directory to executable location: {fallbackDir}");
                        }
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
                        if (!SettingsManager.SuppressLaunchWarnings)
                        {
                            MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to start: {targetPath}", "Launch Error");
                        }
                    }
                }
                catch (Exception ex)
              
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Exception starting process: {ex.Message}\nTarget: {targetPath}\nArguments: {finalArguments}");
                    if (!SettingsManager.SuppressLaunchWarnings)
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error launching: {ex.Message}", "Launch Error");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Critical error in LaunchItem for {path}: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error opening item: {ex.Message}", "Error");
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

       // Updates fence JSON data after icon refresh

        private static void UpdateFenceDataForIcon(string shortcutPath, string newDisplayName, NonActivatingWindow parentWindow)
        {
            try
            {
                if (string.IsNullOrEmpty(newDisplayName)) return;

                string fenceId = parentWindow.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                var fence = FenceDataManager.FenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
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

                FenceDataManager.SaveFenceData();
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

                dynamic fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == fence.Title);
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