using IWshRuntimeLibrary;
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

namespace Desktop_Fences
{
    public static class FenceManager
    {
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


        // Add near other static fields
        private static TargetChecker _currentTargetChecker;

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
                Log($"Error reloading fences: {ex.Message}");
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
                    MessageBox.Show("Invalid backup folder - missing required files", "Restore Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
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
                Log($"Restore failed: {ex.Message}");
                MessageBox.Show($"Restore failed: {ex.Message}", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error);
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
            var aboutItem = new MenuItem { Header = "About" };
            aboutItem.Click += (s, e) => TrayManager.Instance.ShowAboutForm();
            menu.Items.Add(aboutItem);

            // Options item
            var optionsItem = new MenuItem { Header = "Options" };
            optionsItem.Click += (s, e) => TrayManager.Instance.ShowOptionsForm();
            menu.Items.Add(optionsItem);

            // Separator
            menu.Items.Add(new Separator());

            // New Fence item
            var newFenceItem = new MenuItem { Header = "New Fence" };
            newFenceItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                Log($"Creating new fence at mouse position: X={mousePosition.X}, Y={mousePosition.Y}");
                //CreateNewFence("New Fence", "Data", mousePosition.X, mousePosition.Y);
                CreateNewFence("", "Data", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newFenceItem);

            // New Portal Fence item
            var newPortalFenceItem = new MenuItem { Header = "New Portal Fence" };
            newPortalFenceItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                Log($"Creating new portal fence at mouse position: X={mousePosition.X}, Y={mousePosition.Y}");
                CreateNewFence("New Portal Fence", "Portal", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newPortalFenceItem);

            // Restore Fence item
            var restoreItem = new MenuItem
            {
                Header = "Restore Last Deleted Fence",
                Visibility = _isRestoreAvailable ? Visibility.Visible : Visibility.Collapsed
            };
            restoreItem.Click += (s, e) => RestoreLastDeletedFence();
            menu.Items.Add(restoreItem);

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
                        Log($"Updated heart ContextMenu for fence '{fence.Title}'");
                    }
                    else
                    {
                        Log($"Skipped update for fence '{fence.Title}': heart TextBlock is null");
                    }
                }
            });
        }





       // public static MenuItem restoreFenceItem { get; private set; }

        public enum LaunchEffect
        {
            Zoom,        // First effect
            Bounce,      // Scale up and down a few times
            FadeOut,        // Fade out and back in
            SlideUp,     // Slide up and return
            Rotate,      // Spin 360 degrees
            Agitate, // Shake back and forth
GrowAndFly,   // New effect - grow and fly away
Pulse,
Elastic,
Flip3D,
Spiral

        }

        // add here:
        public static List<dynamic> GetFenceData()
        {
            return _fenceData;
        }
        private static void Log(string message)
        {
            bool isLogEnabled = _options?.IsLogEnabled ?? true;
            if (isLogEnabled)
            {
                string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
            }
        }
        // Update fence property, save to JSON, and apply runtime changes

        public static void UpdateFenceProperty(dynamic fence, string propertyName, string value, string logMessage)
        {
            try
            {
                // Get the actual fence object from _fenceData to ensure we're modifying the correct instance
                // int index = _fenceData.FindIndex(f => f.Title == fence.Title.ToString());

                // Find the index by reference instead of title
                // int index = _fenceData.IndexOf(fence);

                string fenceId = fence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    Log($"Fence '{fence.Title}' has no Id");
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
                    Log($"{logMessage} for fence '{fence.Title}'");

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
                            Log($"Applied color '{value ?? "Default"}' to fence '{fence.Title}' at runtime");
                        }
                        else if (propertyName == "IsHidden")
                        {
                            // Update visibility based on IsHidden
                            bool isHidden = value?.ToLower() == "true";
                            win.Visibility = isHidden ? Visibility.Hidden : Visibility.Visible;
                            Log($"Set visibility to {(isHidden ? "Hidden" : "Visible")} for fence '{fence.Title}'");
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
                        Log($"Failed to find window for fence '{fence.Title}' to apply {propertyName}");
                    }
                }
                else
                {
                    Log($"Failed to find fence '{fence.Title}' in _fenceData for {propertyName} update");
                }
            }
            catch (Exception ex)
            {
                Log($"Error updating {propertyName} for fence '{fence.Title}': {ex.Message}");
            }
        }

        public static void RestoreLastDeletedFence()
        {
            if (!_isRestoreAvailable || string.IsNullOrEmpty(_lastDeletedFolderPath) || _lastDeletedFence == null)
            {
                MessageBox.Show("No fence to restore.", "Restore", MessageBoxButton.OK, MessageBoxImage.Information);
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


            Log("Fence restored successfully.");
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
            //UpdateHeartContextMenus();
            //UpdateRestoreFenceItemVisibility(restoreFenceItem);
            //UpdateRestoreFenceItemInAllFences();

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
                TintValue = SettingsManager.TintValue,
                SelectedColor = SettingsManager.SelectedColor,
                IsLogEnabled = SettingsManager.IsLogEnabled,
                singleClickToLaunch = SettingsManager.SingleClickToLaunch,
                LaunchEffect = SettingsManager.LaunchEffect
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
                        FenceManager.Log($"Marked Portal Fence '{fence.Title}' for removal due to missing target folder: {targetPath ?? "null"}");
                    }
                }
            }

            // Remove invalid fences and save
            if (invalidFences.Any())
            {
                foreach (var fence in invalidFences)
                {
                    _fenceData.Remove(fence);
                    FenceManager.Log($"Removed Portal Fence '{fence.Title}' from _fenceData");
                }
                SaveFenceData();
                FenceManager.Log($"Saved updated fences.json after removing {invalidFences.Count} invalid Portal Fences");
            }
            foreach (dynamic fence in _fenceData)
            {
                CreateFence(fence, targetChecker);
            }
        }
       

        private static void InitializeDefaultFence()
        {
            string defaultJson = "[{\"Title\":\"New Fence\",\"X\":20,\"Y\":20,\"Width\":230,\"Height\":130,\"ItemsType\":\"Data\",\"Items\":[]}]";
            System.IO.File.WriteAllText(_jsonFilePath, defaultJson);
            _fenceData = JsonConvert.DeserializeObject<List<dynamic>>(defaultJson);
        }
        private static void MigrateLegacyJson()
        {
            try
            {
                bool jsonModified = false;
                var validColors = new HashSet<string> { "Red", "Green","Teal", "Blue", "Bismark", "White", "Beige", "Gray", "Black", "Purple","Fuchsia", "Yellow", "Orange" };
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
                        Log($"Added Id to {fence.Title}");
                    }

                    // Existing migration: Handle Portal Fence ItemsType
                    if (fence.ItemsType?.ToString() == "Portal")
                    {
                        string portalPath = fence.Items?.ToString();
                        if (!string.IsNullOrEmpty(portalPath) && !System.IO.Directory.Exists(portalPath))
                        {
                            fenceDict["IsFolder"] = true;
                            jsonModified = true;
                            Log($"Migrated legacy portal fence {fence.Title}: Added IsFolder=true");
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
                                Log($"Migrated item in {fence.Title}: Added IsFolder for {path}");
                            }
                        }
                        fenceDict["Items"] = items; // Ensure Items updates are captured
                    }

                    // Migration: Add or validate CustomColor
                    if (!fenceDict.ContainsKey("CustomColor"))
                    {
                        fenceDict["CustomColor"] = null;
                        jsonModified = true;
                        Log($"Added CustomColor=null to {fence.Title}");
                    }
                    else
                    {
                        string customColor = fenceDict["CustomColor"]?.ToString();
                        if (!string.IsNullOrEmpty(customColor) && !validColors.Contains(customColor))
                        {
                            Log($"Invalid CustomColor '{customColor}' in {fence.Title}, resetting to null");
                            fenceDict["CustomColor"] = null;
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate CustomLaunchEffect
                    if (!fenceDict.ContainsKey("CustomLaunchEffect"))
                    {
                        fenceDict["CustomLaunchEffect"] = null;
                        jsonModified = true;
                        Log($"Added CustomLaunchEffect=null to {fence.Title}");
                    }
                    else
                    {
                        string customEffect = fenceDict["CustomLaunchEffect"]?.ToString();
                        if (!string.IsNullOrEmpty(customEffect) && !validEffects.Contains(customEffect))
                        {
                            Log($"Invalid CustomLaunchEffect '{customEffect}' in {fence.Title}, resetting to null");
                            fenceDict["CustomLaunchEffect"] = null;
                            jsonModified = true;
                        }
                    }

                    // Migration: Add or validate IsHidden
                    if (!fenceDict.ContainsKey("IsHidden"))
                    {
                        fenceDict["IsHidden"] = "false"; // Default to string "false"
                        jsonModified = true;
                        Log($"Added IsHidden=\"false\" to {fence.Title}");
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
                            Log($"Invalid IsHidden value '{fenceDict["IsHidden"]}' in {fence.Title}, resetting to \"false\"");
                            isHidden = false;
                            jsonModified = true;
                        }
                        fenceDict["IsHidden"] = isHidden.ToString().ToLower(); // Store as "true" or "false"
                        Log($"Preserved IsHidden: \"{isHidden.ToString().ToLower()}\" for {fence.Title}");
                    }

                    // Update the original fence in _fenceData with the modified dictionary
                    _fenceData[i] = JObject.FromObject(fenceDict);
                }

                if (jsonModified)
                {
                    SaveFenceData();
                    Log("Migrated fences.json with updated fields");
                }
                else
                {
                    Log("No migration needed for fences.json");
                }
            }
            catch (Exception ex)
            {
                Log($"Error migrating fences.json: {ex.Message}\nStackTrace: {ex.StackTrace}");
            }
        }
        

   private static void CreateFence(dynamic fence, TargetChecker targetChecker)
        {
            // Check for valid Portal Fence target folder
            if (fence.ItemsType?.ToString() == "Portal")
            {
                string targetPath = fence.Path?.ToString();
                if (string.IsNullOrEmpty(targetPath) || !System.IO.Directory.Exists(targetPath))
                {
                    FenceManager.Log($"Skipping creation of Portal Fence '{fence.Title}' due to missing target folder: {targetPath ?? "null"}");
                    _fenceData.Remove(fence);
                    SaveFenceData();
                    return;
                }
            }
            DockPanel dp = new DockPanel();
            Border cborder = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(100, 0, 0, 0)),
                CornerRadius = new CornerRadius(6),
                Child = dp
            };

         

            // NEW: Add heart symbol in top-left corner
            TextBlock heart = new TextBlock
            {
                Text = "♥",
                FontSize = 22,
                Foreground = Brushes.White, // Match title and icon text color
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

            ContextMenu cm = new ContextMenu();
            MenuItem miNF = new MenuItem { Header = "New Fence" };
            MenuItem miNP = new MenuItem { Header = "New Portal Fence" };
            MenuItem miRF = new MenuItem { Header = "Delete Fence" };
           // MenuItem miXT = new MenuItem { Header = "Exit" };
            MenuItem miHide = new MenuItem { Header = "Hide Fence" }; // New Hide Fence item

            // Add Customize submenu
            MenuItem miCustomize = new MenuItem { Header = "Customize Fence" };
            MenuItem miColors = new MenuItem { Header = "Color" };
            MenuItem miEffects = new MenuItem { Header = "Launch Effect" };
            
            // Valid options from MigrateLegacyJson
            var validColors = new HashSet<string> { "Red", "Green", "Teal", "Blue", "Bismark", "White", "Beige", "Gray", "Black", "Purple","Fuchsia", "Yellow","Orange" };
            var validEffects = Enum.GetNames(typeof(LaunchEffect)).ToHashSet();
            string currentCustomColor = fence.CustomColor?.ToString();
            string currentCustomEffect = fence.CustomLaunchEffect?.ToString();
        
          
            // Add color options
            MenuItem miColorDefault = new MenuItem { Header = "Default", Tag = null };
            //  miColorDefault.Click += (s, e) => UpdateFenceProperty(fence, "CustomColor", null, "Color set to Default");
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

                // miEffect.Click += (s, e) => UpdateFenceProperty(fence, "CustomLaunchEffect", effect, $"Launch Effect set to {effect}");
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

           // cm.Items.Add(miNF);
          //  cm.Items.Add(miNP);
            cm.Items.Add(miRF);
            cm.Items.Add(miHide); // Add Hide Fence
            cm.Items.Add(new Separator());
            cm.Items.Add(miCustomize); // Add Customize submenu
                                       //  cm.Items.Add(new Separator());

            //   cm.Items.Add(new Separator());
            //   cm.Items.Add(miXT);

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
                Log($"Error reading IsHidden: {ex.Message}");
            }

            bool isHidden = isHiddenString == "true";
            Log($"Fence '{fence.Title}' IsHidden state: {isHidden}");

            // Check IsHidden and apply visibility
            //  bool isHidden = false;
            // Convert dynamic to JObject for proper value access

            //    JObject fenceJObject = (JObject)fence;



            //// Get IsHidden value correctly from JToken
            //string isHiddenString = fenceJObject["IsHidden"]?.ToString().ToLower();
            //bool isHidden = isHiddenString == "true";
            //// Log the actual parsed value
            //Log($"Fence '{fence.Title}' IsHidden state: {isHidden}");

            //if (fence.IsHidden != null)
            //{
            //    // Handle both string and boolean IsHidden values
            //    if (fence.IsHidden is bool boolHidden)
            //    {
            //        isHidden = boolHidden;
            //    }
            //    else if (fence.IsHidden is string stringValue)
            //    {
            //        isHidden = stringValue.ToLower() == "true";
            //    }
            //}
            NonActivatingWindow win = new NonActivatingWindow
            {
                ContextMenu = cm,
                AllowDrop = true,
                AllowsTransparency = true,
                Background = Brushes.Transparent,
                Title = fence.Title.ToString(),
                ShowInTaskbar = false,
                WindowStyle = WindowStyle.None,
                Content = cborder,
                ResizeMode = ResizeMode.CanResizeWithGrip,
                Width = (double)fence.Width,
                Height = (double)fence.Height,
                Top = (double)fence.Y,
                Left = (double)fence.X,
                //  Tag = fence  // Add this line to store the fence object
                Tag = fence.Id.ToString()  // Store fence ID in Tag
              //  Visibility = isHidden ? Visibility.Hidden : Visibility.Visible
            };

            // Log the IsHidden state for diagnostics
            //   Log($"Fence '{fence.Title}' IsHidden state: {isHidden}");
            // Show the window first
            // Defer hiding until after the window is loaded
            win.Loaded += (s, e) =>
            {
                if (isHidden)
                {
                    win.Visibility = Visibility.Hidden;
                    TrayManager.AddHiddenFence(win);
                    Log($"Hid fence '{fence.Title}' after loading at startup");
                }
            };

            //if (isHidden)
            //{
            //    win.Visibility = Visibility.Hidden;
            //  //  TrayManager.AddHiddenFence(win);
            //    Log($"Hid fence '{fence.Title}' at startup");
            //}

            if (isHidden)
            {
                TrayManager.AddHiddenFence(win);
                Log($"Added fence '{fence.Title}' to hidden list at startup");
            }

            // Hide click
            miHide.Click += (s, e) =>
            {
                UpdateFenceProperty(fence, "IsHidden", "true", $"Hid fence '{fence.Title}'");
                TrayManager.AddHiddenFence(win);
                Log($"Triggered Hide Fence for '{fence.Title}'");
            };

            win.Show();


            // hide click
            miHide.Click += (s, e) =>
            {
                UpdateFenceProperty(fence, "IsHidden", "true", $"Hid fence '{fence.Title}'");
                TrayManager.AddHiddenFence(win);
                Log($"Triggered Hide Fence for '{fence.Title}'");
            };
            miRF.Click += (s, e) =>
            {
                bool result = TrayManager.Instance.ShowCustomMessageBoxForm(); // Call the method and store the result  
              //  System.Windows.Forms.MessageBox.Show(result.ToString()); // Display the result in a MessageBox  

                //   var result = MessageBox.Show("Are you sure you want to remove this fence?", "Confirm", MessageBoxButton.YesNo);
                //   if (result == MessageBoxResult.Yes)
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
                    UpdateAllHeartContextMenus();

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
                                    Log($"Skipped backing up missing file: {itemFilePath}");
                                }
                            }
                        }
                    }

                    // Save fence info to JSON
                    string fenceJsonPath = System.IO.Path.Combine(_lastDeletedFolderPath, "fence.json");
                    System.IO.File.WriteAllText(fenceJsonPath, JsonConvert.SerializeObject(fence, Formatting.Indented));

                    // Proceed with deletion
                    _fenceData.Remove(fence);
                    _heartTextBlocks.Remove(fence);
                    if (_portalFences.ContainsKey(fence))
                    {
                        _portalFences.Remove(fence);
                    }
                    SaveFenceData();
                    win.Close();
                    Log($"Fence {fence.Title} removed successfully");
                }
            };


            miNF.Click += (s, e) =>
            {
                bool isLogEnabled = _options.IsLogEnabled ?? true;
                void Log(string message)
                {
                    if (isLogEnabled)
                    {
                        string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                        System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                    }
                }

                Point mousePosition = win.PointToScreen(new Point(0, 0));
                mousePosition = Mouse.GetPosition(win);
                Point absolutePosition = win.PointToScreen(mousePosition);
                FenceManager.Log($"Creating new fence at position: X={absolutePosition.X}, Y={absolutePosition.Y}");
                CreateNewFence("New Fence", "Data", absolutePosition.X, absolutePosition.Y);
            };

            miNP.Click += (s, e) =>
            {
                bool isLogEnabled = _options.IsLogEnabled ?? true;
                void Log(string message)
                {
                    if (isLogEnabled)
                    {
                        string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                        System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                    }
                }

                Point mousePosition = win.PointToScreen(new Point(0, 0));
                mousePosition = Mouse.GetPosition(win);
                Point absolutePosition = win.PointToScreen(mousePosition);
                FenceManager.Log($"Creating new portal fence at position: X={absolutePosition.X}, Y={absolutePosition.Y}");
                CreateNewFence("New Portal Fence", "Portal", absolutePosition.X, absolutePosition.Y);
            };

         //   miXT.Click += (s, e) => Application.Current.Shutdown();

            Label titlelabel = new Label
            {
                Content = fence.Title.ToString(),
                Background = new SolidColorBrush(Color.FromArgb(20, 0, 0, 0)),
                Foreground = Brushes.White,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                Cursor = Cursors.SizeAll
            };
            DockPanel.SetDock(titlelabel, Dock.Top);
            dp.Children.Add(titlelabel);

            TextBox titletb = new TextBox
            {
                HorizontalContentAlignment = HorizontalAlignment.Center,
                Visibility = Visibility.Collapsed
            };
            DockPanel.SetDock(titletb, Dock.Top);
            dp.Children.Add(titletb);

            //// Προσθήκη κουκίδας για Portal Fences
            //if (fence.ItemsType?.ToString() == "Portal")
            //{
            //    Ellipse portalIndicator = new Ellipse
            //    {
            //        Width = 10,
            //        Height = 10,
            //        Fill = Brushes.WhiteSmoke,
            //        Margin = new Thickness(2, 2, 0, 0), // Πάνω αριστερά, κοντά στο titlebar
            //        HorizontalAlignment = HorizontalAlignment.Left,
            //        VerticalAlignment = VerticalAlignment.Top
            //    };
            //    dp.Children.Add(portalIndicator);
            //}
        

            titlelabel.MouseDown += (sender, e) =>
            {
                if (e.ClickCount == 2)
                {
                    bool isLogEnabled = _options.IsLogEnabled ?? true;
                    void Log(string message)
                    {
                        if (isLogEnabled)
                        {
                            string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                            System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                        }
                    }

                    FenceManager.Log($"Entering edit mode for fence: {fence.Title}");

                    titletb.Text = titlelabel.Content.ToString();
                    titlelabel.Visibility = Visibility.Collapsed;
                    titletb.Visibility = Visibility.Visible;
                    win.ShowActivated = true;
                    win.Activate();
                    Keyboard.Focus(titletb);
                    titletb.SelectAll();
                    FenceManager.Log($"Focus set to title textbox for fence: {fence.Title}");
                }
                else if (e.LeftButton == MouseButtonState.Pressed)
                {
                    win.DragMove();
                }
            };

            titletb.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    bool isLogEnabled = _options.IsLogEnabled ?? true;
                    void Log(string message)
                    {
                        if (isLogEnabled)
                        {
                            string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                            System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                        }
                    }

                    fence.Title = titletb.Text;
                    titlelabel.Content = titletb.Text;
                    win.Title = titletb.Text;
                    titletb.Visibility = Visibility.Collapsed;
                    titlelabel.Visibility = Visibility.Visible;
                    SaveFenceData();
                    win.ShowActivated = false;
                    FenceManager.Log($"Exited edit mode via Enter, new title for fence: {fence.Title}");
                    win.Focus();
                }
            };

            titletb.LostFocus += (sender, e) =>
            {
                bool isLogEnabled = _options.IsLogEnabled ?? true;
                void Log(string message)
                {
                    if (isLogEnabled)
                    {
                        string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                        System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                    }
                }

                fence.Title = titletb.Text;
                titlelabel.Content = titletb.Text;
                win.Title = titletb.Text;
                titletb.Visibility = Visibility.Collapsed;
                titlelabel.Visibility = Visibility.Visible;
                SaveFenceData();
                win.ShowActivated = false;
                FenceManager.Log($"Exited edit mode via click, new title for fence: {fence.Title}");
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
                wpcontscr.Background = new ImageBrush
                {
                    ImageSource = new BitmapImage(new Uri("pack://application:,,,/Resources/portal.png")),
                    Opacity = 0.2,
                    Stretch = Stretch.UniformToFill
                };
            }

            // Προσθήκη watermark για Portal Fences
            if (fence.ItemsType?.ToString() == "Portal")
            {
                wpcontscr.Background = new ImageBrush
                {
                    ImageSource = new BitmapImage(new Uri("pack://application:,,,/Resources/portal.png")),
                    Opacity = 0.2,
                    Stretch = Stretch.UniformToFill
                };
            }

            dp.Children.Add(wpcontscr);

            void InitContent()
            {
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

                                targetChecker.AddCheckAction(filePath, () => UpdateIcon(sp, filePath, isFolder), isFolder);

                                ContextMenu mn = new ContextMenu();
                                MenuItem miRunAsAdmin = new MenuItem { Header = "Run as administrator" };
                                MenuItem miE = new MenuItem { Header = "Edit..." };
                                MenuItem miM = new MenuItem { Header = "Move..." };
                                MenuItem miRemove = new MenuItem { Header = "Remove" };
                                MenuItem miFindTarget = new MenuItem { Header = "Open target folder..." };
                                MenuItem miCopyPath = new MenuItem { Header = "Copy path" };
                                MenuItem miCopyFolder = new MenuItem { Header = "Folder path" };
                                MenuItem miCopyFullPath = new MenuItem { Header = "Full path" };
                                miCopyPath.Items.Add(miCopyFolder);
                                miCopyPath.Items.Add(miCopyFullPath);

                                mn.Items.Add(miE);
                                mn.Items.Add(miM);
                                mn.Items.Add(miRemove);
                                mn.Items.Add(new Separator());
                                mn.Items.Add(miRunAsAdmin);
                                mn.Items.Add(new Separator());
                                mn.Items.Add(miCopyPath);
                                mn.Items.Add(miFindTarget);
                                sp.ContextMenu = mn;

                                miRunAsAdmin.IsEnabled = Utility.IsExecutableFile(filePath);
                                miM.Click += (sender, e) => MoveItem(icon, fence, win.Dispatcher);
                                miE.Click += (sender, e) => EditItem(icon, fence, win);
                                miRemove.Click += (sender, e) =>
                                {
                                    var items = fence.Items as JArray;
                                    if (items != null)
                                    {
                                        var itemToRemove = items.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                                        if (itemToRemove != null)
                                        {
                                            bool isLogEnabled = _options.IsLogEnabled ?? true;
                                            void Log(string message)
                                            {
                                                if (isLogEnabled)
                                                {
                                                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                                                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                                                }
                                            }
                                            FenceManager.Log($"Removing icon for {filePath} from fence");
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
                                                        FenceManager.Log($"Deleted shortcut: {shortcutPath}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        FenceManager.Log($"Failed to delete shortcut {shortcutPath}: {ex.Message}");
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
                                        Log($"Run as admin with arguments: {runArguments} for {filePath}");
                                    }
                                    try
                                    {
                                        Process.Start(psi);
                                        Log($"Successfully launched {targetPath} as admin");
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Failed to launch {targetPath} as admin: {ex.Message}");
                                        MessageBox.Show($"Error running as admin: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                        MessageBox.Show($"Failed to initialize Portal Fence: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                            Debug.WriteLine($"Dropped file: {droppedFile}");
                            if (!System.IO.File.Exists(droppedFile) && !System.IO.Directory.Exists(droppedFile))
                            {
                                MessageBox.Show($"Invalid file or directory: {droppedFile}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                continue;
                            }

                            bool isLogEnabled = _options.IsLogEnabled ?? true;
                            void Log(string message)
                            {
                                if (isLogEnabled)
                                {
                                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                                }
                            }

                            if (fence.ItemsType?.ToString() == "Data")
                            {
                                // Λογική για Data Fences (shortcuts)
                                if (!System.IO.Directory.Exists("Shortcuts")) System.IO.Directory.CreateDirectory("Shortcuts");
                                string baseShortcutName = System.IO.Path.Combine("Shortcuts", System.IO.Path.GetFileName(droppedFile));
                                string shortcutName = baseShortcutName;
                                int counter = 1;

                                bool isDroppedShortcut = System.IO.Path.GetExtension(droppedFile).ToLower() == ".lnk";
                                string targetPath = isDroppedShortcut ? Utility.GetShortcutTarget(droppedFile) : droppedFile;
                                bool isFolder = System.IO.Directory.Exists(targetPath) || (isDroppedShortcut && string.IsNullOrEmpty(System.IO.Path.GetExtension(targetPath)));

                                if (!isDroppedShortcut)
                                {
                                    shortcutName = baseShortcutName + ".lnk";
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{System.IO.Path.GetFileNameWithoutExtension(droppedFile)} ({counter}).lnk");
                                        counter++;
                                    }
                                    WshShell shell = new WshShell();
                                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutName);
                                    shortcut.TargetPath = droppedFile;
                                    shortcut.Save();
                                    FenceManager.Log($"Created unique shortcut: {shortcutName}");
                                }
                                else
                                {
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{System.IO.Path.GetFileNameWithoutExtension(droppedFile)} ({counter}).lnk");
                                        counter++;
                                    }
                                    System.IO.File.Copy(droppedFile, shortcutName, false);
                                    FenceManager.Log($"Copied unique shortcut: {shortcutName}");
                                }

                                dynamic newItem = new System.Dynamic.ExpandoObject();
                                IDictionary<string, object> newItemDict = newItem;
                                newItemDict["Filename"] = shortcutName;
                                newItemDict["IsFolder"] = isFolder; // Ορίζουμε το isFolder με βάση τον στόχο
                                newItemDict["DisplayName"] = System.IO.Path.GetFileNameWithoutExtension(droppedFile);
                              //  newItemDict["Arguments"] = arguments; // Add this line to store arguments


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
                                    MenuItem miE = new MenuItem { Header = "Edit" };
                                    MenuItem miM = new MenuItem { Header = "Move.." };
                                    MenuItem miRemove = new MenuItem { Header = "Remove" };
                                    MenuItem miFindTarget = new MenuItem { Header = "Find target..." };
                                    MenuItem miCopyPath = new MenuItem { Header = "Copy path" };
                                    MenuItem miCopyFolder = new MenuItem { Header = "Folder" };
                                    MenuItem miCopyFullPath = new MenuItem { Header = "Full path" };
                                    miCopyPath.Items.Add(miCopyFolder);
                                    miCopyPath.Items.Add(miCopyFullPath);

                                    mn.Items.Add(miE);
                                    mn.Items.Add(miM);
                                    mn.Items.Add(miRemove);
                                    mn.Items.Add(new Separator());
                                    mn.Items.Add(miRunAsAdmin);
                                    mn.Items.Add(new Separator());
                                    mn.Items.Add(miCopyPath);
                                    mn.Items.Add(miFindTarget);
                                    sp.ContextMenu = mn;

                                    miRunAsAdmin.IsEnabled = Utility.IsExecutableFile(shortcutName);
                                    miM.Click += (sender, e) => MoveItem(newItem, fence, win.Dispatcher);
                                    miE.Click += (sender, e) => EditItem(newItem, fence, win);
                                    miRemove.Click += (sender, e) =>
                                    {
                                        var items = fence.Items as JArray;
                                        if (items != null)
                                        {
                                            var itemToRemove = items.FirstOrDefault(i => i["Filename"]?.ToString() == shortcutName);
                                            if (itemToRemove != null)
                                            {
                                                FenceManager.Log($"Removing icon for {shortcutName} from fence");
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
                                                            FenceManager.Log($"Deleted shortcut: {shortcutPath}");
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            FenceManager.Log($"Failed to delete shortcut {shortcutPath}: {ex.Message}");
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
                                        FenceManager.Log($"Copied target folder path to clipboard: {folderPath}");
                                    };

                                    miCopyFullPath.Click += (sender, e) =>
                                    {
                                        string targetPath = Utility.GetShortcutTarget(shortcutName);
                                        Clipboard.SetText(targetPath);
                                        FenceManager.Log($"Copied target full path to clipboard: {targetPath}");
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
                                FenceManager.Log($"Added shortcut to Data Fence: {shortcutName}");
                            }
                            else if (fence.ItemsType?.ToString() == "Portal")
                            {
                                // Λογική για Portal Fences (copy) - παραμένει ίδια
                                IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
                                string destinationFolder = fenceDict.ContainsKey("Path") ? fenceDict["Path"]?.ToString() : null;

                                if (string.IsNullOrEmpty(destinationFolder))
                                {
                                    MessageBox.Show($"No destination folder defined for this Portal Fence. Please recreate the fence.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    FenceManager.Log($"No Path defined for Portal Fence: {fence.Title}");
                                    continue;
                                }

                                if (!System.IO.Directory.Exists(destinationFolder))
                                {
                                    MessageBox.Show($"The destination folder '{destinationFolder}' no longer exists. Please update the Portal Fence settings.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    FenceManager.Log($"Destination folder missing for Portal Fence: {destinationFolder}");
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
                                    FenceManager.Log($"Copied file to Portal Fence: {destinationPath}");
                                }
                                else if (System.IO.Directory.Exists(droppedFile))
                                {
                                    CopyDirectory(droppedFile, destinationPath);
                                    FenceManager.Log($"Copied directory to Portal Fence: {destinationPath}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error in drop: {ex.Message}");
                            MessageBox.Show($"Failed to add {droppedFile}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    SaveFenceData(); // Αποθηκεύουμε τις αλλαγές στο JSON
                }
            };

            win.SizeChanged += (s, e) =>
            {
                fence.Width = win.Width;
                fence.Height = win.Height;
                SaveFenceData();
            };

            win.LocationChanged += (s, e) =>
            {
                fence.X = win.Left;
                fence.Y = win.Top;
                SaveFenceData();
            };

            InitContent();
            win.Show();

            //IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
            //SnapManager.AddSnapping(win, fenceDict);
            //Utility.ApplyTintAndColorToFence(win);
            //targetChecker.Start();

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

        public static void AddIcon(dynamic icon, WrapPanel wpcont)
        {
            StackPanel sp = new StackPanel { Margin = new Thickness(5), Width = 60 };
            Image ico = new Image { Width = 40, Height = 40, Margin = new Thickness(5) };
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
            string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;
            string arguments = iconDict.ContainsKey("Arguments") ? (string)iconDict["Arguments"] : null; // Add this line

            bool isLogEnabled = _options.IsLogEnabled ?? true;
            void Log(string message)
            {
                if (isLogEnabled)
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                }
            }

            ImageSource shortcutIcon = null;
           // string arguments = null;

            if (isShortcut)
            {
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                targetPath = shortcut.TargetPath;
                arguments = shortcut.Arguments;

                if (System.IO.Directory.Exists(targetPath))
                {
                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    Log($"Using folder-White.png for shortcut {filePath} targeting existing folder {targetPath}");
                }
                else if (System.IO.File.Exists(targetPath))
                {
                    try
                    {
                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                        Log($"Using target file icon for {filePath}: {targetPath}");
                    }
                    catch (Exception ex)
                    {
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        Log($"Failed to extract target file icon for {filePath}: {ex.Message}");
                    }
                }
                else
                {
                    shortcutIcon = isFolder
                        ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                        : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                    Log($"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")}");
                }

                if (!string.IsNullOrEmpty(shortcut.IconLocation))
                {
                    string[] iconParts = shortcut.IconLocation.Split(',');
                    string iconPath = iconParts[0];
                    if (System.IO.File.Exists(iconPath))
                    {
                        try
                        {
                            shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(iconPath).ToImageSource();
                            Log($"Using custom icon from IconLocation for {filePath}: {iconPath}");
                        }
                        catch (Exception ex)
                        {
                            Log($"Failed to extract custom icon for {filePath}: {ex.Message}");
                        }
                    }
                }
            }
            else
            {
                if (System.IO.Directory.Exists(filePath))
                {
                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    Log($"Using folder-White.png for {filePath}");
                }
                else if (System.IO.File.Exists(filePath))
                {
                    try
                    {
                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource();
                        Log($"Using file icon for {filePath}");
                    }
                    catch (Exception ex)
                    {
                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        Log($"Failed to extract icon for {filePath}: {ex.Message}");
                    }
                }
                else
                {
                    shortcutIcon = isFolder
                        ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                        : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                    Log($"Using missing icon for {filePath}: {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")}");
                }
            }

            ico.Source = shortcutIcon;
            sp.Children.Add(ico);

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
                Foreground = Brushes.White,
                MaxWidth = double.MaxValue,
                Width = double.NaN,
                TextAlignment = TextAlignment.Center,
                Text = displayText
            };
            // Add shadow effect for better contrast
            lbl.Effect = new DropShadowEffect
            {
                Color = Colors.Black,       // Black shadow for maximum contrast
                Direction = 315,            // Angle: slightly upper-left for natural look
                ShadowDepth = 2,            // Small offset to avoid blurriness
                BlurRadius = 3,             // Subtle blur for readability
                Opacity = 0.8               // Strong enough to stand out
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

        // Save fence data to JSON with consistent IsHidden string format
        public static void SaveFenceData()
        {
            // Pre-process _fenceData to ensure IsHidden is stored as string "true" or "false"
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
                    fenceDict["IsHidden"] = isHidden.ToString().ToLower(); // Store as "true" or "false"
                }
                serializedData.Add(JObject.FromObject(fenceDict));
            }

            // Serialize with indented formatting
            string formattedJson = JsonConvert.SerializeObject(serializedData, Formatting.Indented);
            System.IO.File.WriteAllText(_jsonFilePath, formattedJson);
            Log($"Saved fences.json with consistent IsHidden string format");
        }
        //public static void SaveFenceData()
        //{
        //    string formattedJson = JsonConvert.SerializeObject(_fenceData, Formatting.Indented);
        //    System.IO.File.WriteAllText(_jsonFilePath, formattedJson);
        //}

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
            bool isLogEnabled = _options.IsLogEnabled ?? true;
            void Log(string message)
            {
                if (isLogEnabled)
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                }
            }

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
                    //btn.Click += (s, e) =>
                    //{
                    //    var sourceItems = sourceFence.Items as JArray;
                    //    var destItems = fence.Items as JArray ?? new JArray();
                    //    if (sourceItems != null)
                    //    {
                    //        IDictionary<string, object> itemDict = item is IDictionary<string, object> dict ? dict : ((JObject)item).ToObject<IDictionary<string, object>>();
                    //        string filename = itemDict.ContainsKey("Filename") ? itemDict["Filename"].ToString() : "Unknown";

                    //        Log($"Moving item {filename} from {sourceFence.Title} to {fence.Title}");
                    //        sourceItems.Remove(item);
                    //        destItems.Add(item);
                    //        fence.Items = destItems;
                    //        SaveFenceData();
                    //        moveWindow.Close();

                    //        var waitWindow = new Window
                    //        {
                    //            Title = "Desktop Fences +",
                    //            Width = 200,
                    //            Height = 100,
                    //            WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    //            WindowStyle = WindowStyle.None,
                    //            Background = Brushes.LightGray,
                    //            Topmost = true
                    //        };
                    //        var waitStack = new StackPanel
                    //        {
                    //            HorizontalAlignment = HorizontalAlignment.Center,
                    //            VerticalAlignment = VerticalAlignment.Center
                    //        };
                    //        waitWindow.Content = waitStack;

                    //        string exePath = Assembly.GetEntryAssembly().Location;
                    //        var iconImage = new Image
                    //        {
                    //            Source = System.Drawing.Icon.ExtractAssociatedIcon(exePath).ToImageSource(),
                    //            Width = 32,
                    //            Height = 32,
                    //            Margin = new Thickness(0, 0, 0, 5)
                    //        };
                    //        waitStack.Children.Add(iconImage);

                    //        var waitLabel = new Label
                    //        {
                    //            Content = "Please wait...",
                    //            HorizontalAlignment = HorizontalAlignment.Center
                    //        };
                    //        waitStack.Children.Add(waitLabel);

                    //        waitWindow.Show();

                    //        dispatcher.InvokeAsync(() =>
                    //        {
                    //            if (Application.Current != null && !Application.Current.Dispatcher.HasShutdownStarted)
                    //            {
                    //                Application.Current.Windows.OfType<NonActivatingWindow>().ToList().ForEach(w => w.Close());
                    //                LoadAndCreateFences(new TargetChecker(1000));
                    //                waitWindow.Close();
                    //                Log($"Item moved successfully to {fence.Title}");
                    //            }
                    //            else
                    //            {
                    //                waitWindow.Close();
                    //                Log($"Skipped fence reload due to application shutdown");
                    //            }
                    //        }, DispatcherPriority.Background);
                    //    }
                    //};


                    btn.Click += (s, e) =>
                    {
                        var sourceItems = sourceFence.Items as JArray;
                        var destItems = fence.Items as JArray ?? new JArray();
                        if (sourceItems != null)
                        {
                            IDictionary<string, object> itemDict = item is IDictionary<string, object> dict ? dict : ((JObject)item).ToObject<IDictionary<string, object>>();
                            string filename = itemDict.ContainsKey("Filename") ? itemDict["Filename"].ToString() : "Unknown";

                            Log($"Moving item {filename} from {sourceFence.Title} to {fence.Title}");

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
                                Log($"Item {filename} not found in source fence '{sourceFence.Title}'");
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
                                Background = Brushes.LightGray,
                                Topmost = true
                            };
                            var waitStack = new StackPanel
                            {
                                HorizontalAlignment = HorizontalAlignment.Center,
                                VerticalAlignment = VerticalAlignment.Center
                            };
                            waitWindow.Content = waitStack;

                            string exePath = Assembly.GetEntryAssembly().Location;
                            var iconImage = new Image
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
                                    Log($"Item moved successfully to {fence.Title}");
                                }
                                else
                                {
                                    waitWindow.Close();
                                    Log($"Skipped fence reload due to application shutdown");
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
                Log($"FindWrapPanel: Reached max depth {maxDepth} or null parent at depth {depth}");
                return null;
            }

            // Check if current element is a WrapPanel
            if (parent is WrapPanel wrapPanel)
            {
                Log($"FindWrapPanel: Found WrapPanel at depth {depth}");
                return wrapPanel;
            }

            // Recurse through visual tree
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                Log($"FindWrapPanel: Checking child {i} at depth {depth}, type: {child?.GetType()?.Name ?? "null"}");
                var result = FindWrapPanel(child, depth + 1, maxDepth);
                if (result != null)
                {
                    return result;
                }
            }

            Log($"FindWrapPanel: No WrapPanel found under parent {parent?.GetType()?.Name ?? "null"} at depth {depth}");
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
                MessageBox.Show("Edit is only available for shortcuts.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var editWindow = new EditShortcutWindow(filePath, displayName);
            if (editWindow.ShowDialog() == true)
            {
                string newDisplayName = editWindow.NewDisplayName;
                iconDict["DisplayName"] = newDisplayName;
                Log($"Updated DisplayName for {filePath} to {newDisplayName}");

                // Update fence data
                var items = fence.Items as JArray;
                if (items != null)
                {
                    var itemToUpdate = items.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                    if (itemToUpdate != null)
                    {
                        itemToUpdate["DisplayName"] = newDisplayName;
                        SaveFenceData();
                        Log($"Fence data updated for {filePath}");
                    }
                    else
                    {
                        Log($"Failed to find item {filePath} in fence items for update");
                    }
                }

                // Update UI
                if (win == null)
                {
                    Log($"Window is null for fence when updating {filePath}");
                    return;
                }

                // Attempt to find WrapPanel directly
                Log($"Starting UI update for {filePath}. Window content type: {win.Content?.GetType()?.Name ?? "null"}");
                WrapPanel wpcont = null;
                var border = win.Content as Border;
                if (border != null)
                {
                    Log($"Found Border. Child type: {border.Child?.GetType()?.Name ?? "null"}");
                    var dockPanel = border.Child as DockPanel;
                    if (dockPanel != null)
                    {
                        Log($"Found DockPanel. Checking for ScrollViewer...");
                        var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                        if (scrollViewer != null)
                        {
                            Log($"Found ScrollViewer. Content type: {scrollViewer.Content?.GetType()?.Name ?? "null"}");
                            wpcont = scrollViewer.Content as WrapPanel;
                        }
                        else
                        {
                            Log($"ScrollViewer not found in DockPanel for {filePath}");
                        }
                    }
                    else
                    {
                        Log($"DockPanel not found in Border for {filePath}");
                    }
                }
                else
                {
                    Log($"Border not found in window content for {filePath}");
                }

                // Fallback to visual tree traversal if direct lookup fails
                if (wpcont == null)
                {
                    Log($"Direct WrapPanel lookup failed, attempting visual tree traversal for {filePath}");
                    wpcont = FindWrapPanel(win);
                }

                if (wpcont != null)
                {
                    Log($"Found WrapPanel. Checking for StackPanel with Tag.FilePath: {filePath}");
                    // Access Tag.FilePath to match ClickEventAdder's anonymous object
                    var sp = wpcont.Children.OfType<StackPanel>()
                        .FirstOrDefault(s => s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == filePath);
                    if (sp != null)
                    {
                        Log($"Found StackPanel for {filePath}. Tag: {sp.Tag?.ToString() ?? "null"}");
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
                            Log($"Updated TextBlock for {filePath} to '{displayText}'");
                        }
                        else
                        {
                            Log($"TextBlock not found in StackPanel for {filePath}. Children: {string.Join(", ", sp.Children.OfType<FrameworkElement>().Select(c => c.GetType().Name))}");
                        }

                        // Update icon
                        var ico = sp.Children.OfType<Image>().FirstOrDefault();
                        if (ico != null)
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                            ImageSource shortcutIcon = null;

                            if (!string.IsNullOrEmpty(shortcut.IconLocation))
                            {
                                string[] iconParts = shortcut.IconLocation.Split(',');
                                string iconPath = iconParts[0];
                                if (System.IO.File.Exists(iconPath))
                                {
                                    try
                                    {
                                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(iconPath).ToImageSource();
                                        Log($"Applied custom icon from IconLocation for {filePath}: {iconPath}");
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Failed to extract custom icon for {filePath}: {ex.Message}");
                                    }
                                }
                            }

                            if (shortcutIcon == null)
                            {
                                string targetPath = shortcut.TargetPath;
                                if (System.IO.File.Exists(targetPath))
                                {
                                    try
                                    {
                                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                                        Log($"No custom icon, using target icon for {filePath}: {targetPath}");
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Failed to extract target icon for {filePath}: {ex.Message}");
                                    }
                                }
                                else if (System.IO.Directory.Exists(targetPath))
                                {
                                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                                    Log($"No custom icon, using folder-White.png for {filePath}");
                                }
                                else
                                {
                                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                                    Log($"No custom icon or valid target, using file-WhiteX.png for {filePath}");
                                }
                            }

                            ico.Source = shortcutIcon;
                            iconCache[filePath] = shortcutIcon;
                            Log($"Updated icon for {filePath}");
                        }
                        else
                        {
                            Log($"Image not found in StackPanel for {filePath}");
                        }
                    }
                    else
                    {
                        Log($"StackPanel not found for {filePath} in WrapPanel. Available Tag.FilePaths: {string.Join(", ", wpcont.Children.OfType<StackPanel>().Select(s => s.Tag != null ? (s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() ?? "null") : "null"))}");

                        // Fallback: Rebuild single icon
                        Log($"Attempting to rebuild icon for {filePath}");
                        // Remove existing StackPanel if present (in case Tag lookup failed due to corruption)
                        var oldSp = wpcont.Children.OfType<StackPanel>()
                            .FirstOrDefault(s => s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == filePath);
                        if (oldSp != null)
                        {
                            wpcont.Children.Remove(oldSp);
                            Log($"Removed old StackPanel for {filePath}");
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
                            Log($"Rebuilt icon for {filePath} with new StackPanel");
                        }
                        else
                        {
                            Log($"Failed to rebuild icon for {filePath}: No new StackPanel created");
                        }
                    }
                }
                else
                {
                    Log($"WrapPanel not found for {filePath}. UI update skipped.");
                }
            }
        }

        private static void UpdateIcon(StackPanel sp, string filePath, bool isFolder)
        {
            bool isLogEnabled = _options.IsLogEnabled ?? true;
            void Log(string message)
            {
                if (isLogEnabled)
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                }
            }
            if (Application.Current != null)
            {
                Application.Current.Dispatcher.Invoke(() =>
            {
                Image ico = sp.Children.OfType<Image>().FirstOrDefault();
                if (ico == null) return;

                bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
                string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;
                bool targetExists = System.IO.File.Exists(targetPath) || System.IO.Directory.Exists(targetPath);
                bool isTargetFolder = System.IO.Directory.Exists(targetPath);

                ImageSource newIcon = null;

                if (isShortcut)
                {
                    WshShell shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                    string iconLocation = shortcut.IconLocation;

                    // Check if a custom icon is set and the file exists
                    if (!string.IsNullOrEmpty(iconLocation))
                    {
                        string[] iconParts = iconLocation.Split(',');
                        string iconPath = iconParts[0];
                        if (System.IO.File.Exists(iconPath))
                        {
                            try
                            {
                                newIcon = System.Drawing.Icon.ExtractAssociatedIcon(iconPath).ToImageSource();
                               //Log($"Using custom IconLocation for {filePath}: {iconPath}");
                            }
                            catch (Exception ex)
                            {
                                Log($"Failed to load custom icon for {filePath} from {iconPath}: {ex.Message}");
                            }
                        }
                    }
                }

                // If no valid custom icon, fall back to target-based or default logic
                if (newIcon == null)
                {
                    if (targetExists)
                    {
                        if (isTargetFolder)
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                         //   Log($"Target exists, updating to folder-White.png for {filePath}");
                        }
                        else
                        {
                            try
                            {
                                newIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                              //  Log($"Target exists, updating to file icon for {filePath}");
                            }
                            catch (Exception ex)
                            {
                                newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                                Log($"Failed to extract icon for {filePath}: {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        newIcon = isFolder
                            ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
                            : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        Log($"Target missing, updating to {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")} for {filePath}");
                    }
                }

                if (ico.Source != newIcon)
                {
                    ico.Source = newIcon;
                  //  Log($"Icon updated for {filePath}");
                }
            });

            }
            else
            {
                Log("Application.Current is null, cannot update icon.");
            }
        }

    

        public static void UpdateOptionsAndClickEvents()
        {
            //bool isLogEnabled = SettingsManager.IsLogEnabled;

            //void Log(string message)
            //{
            //    if (isLogEnabled)
            //    {
            //        string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
            //        System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
            //    }
            //}

            Log($"Updating options, new singleClickToLaunch={SettingsManager.SingleClickToLaunch}");

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
                Log($"Updated click events for {updatedItems} items");
            });
            }
                else
                {
                    Log("Application.Current is null, cannot update icon.");
                }
        }


        public static void ClickEventAdder(StackPanel sp, string path, bool isFolder, string arguments = null)
        {
            bool isLogEnabled = _options.IsLogEnabled ?? true;

            // Clear existing handler
            sp.MouseLeftButtonDown -= MouseDownHandler;

            // Store only path, isFolder, and arguments in Tag
            sp.Tag = new { FilePath = path, IsFolder = isFolder, Arguments = arguments };

            Log($"Attaching handler for {path}, initial singleClickToLaunch={_options.singleClickToLaunch}");

            // Check if path is a shortcut and correct isFolder for folder shortcuts
            bool isShortcut = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
            string targetPath = isShortcut ? Utility.GetShortcutTarget(path) : path;
            if (isShortcut && System.IO.Directory.Exists(targetPath))
            {
                isFolder = true;
                Log($"Corrected isFolder to true for shortcut {path} targeting folder {targetPath}");
            }

            void MouseDownHandler(object sender, MouseButtonEventArgs e)
            {
                if (e.ChangedButton != MouseButton.Left) return;

                bool singleClickToLaunch = _options.singleClickToLaunch ?? true;
                Log($"MouseDown on {path}, ClickCount={e.ClickCount}, singleClickToLaunch={singleClickToLaunch}");

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

                    Log($"Target check: Path={path}, ResolvedPath={resolvedPath}, IsShortcut={isShortcutLocal}, IsFolder={isFolder}, TargetExists={targetExists}");

                    if (!targetExists)
                    {
                        Log($"Target not found: {resolvedPath}");
                        return;
                    }

                    if (singleClickToLaunch && e.ClickCount == 1)
                    {
                        Log($"Single click launching {path}");
                        LaunchItem(sp, path, isFolder, arguments);
                        e.Handled = true;
                    }
                    else if (!singleClickToLaunch && e.ClickCount == 2)
                    {
                        Log($"Double click launching {path}");
                        LaunchItem(sp, path, isFolder, arguments);
                        e.Handled = true;
                    }
                }
                catch (Exception ex)
                {
                    Log($"Error checking target existence: {ex.Message}");
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
    bool isLogEnabled = _options.IsLogEnabled ?? true;
    void Log(string message)
    {
        if (isLogEnabled)
        {
            string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
            System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
        }
    }

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
            Log($"Failed to find fence for {path}, using default effect");
            effect = defaultEffect;
        }
        else if (string.IsNullOrEmpty(customEffect))
        {
            effect = defaultEffect;
            Log($"No CustomLaunchEffect for {path} in fence '{fence.Title}', using default: {effect}");
        }
        else
        {
            try
            {
                effect = (LaunchEffect)Enum.Parse(typeof(LaunchEffect), customEffect, true);
                Log($"Using CustomLaunchEffect for {path} in fence '{fence.Title}': {effect}");
            }
            catch (Exception ex)
            {
                effect = defaultEffect;
                Log($"Failed to parse CustomLaunchEffect '{customEffect}' for {path} in fence '{fence.Title}', falling back to {effect}: {ex.Message}");
            }
        }

        Log($"LaunchItem called for {path} with effect {effect}");

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
            sp.RenderTransformOrigin = new Point(0.5, 0.5);
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

                                // Small delay before showing the icon again
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

        Log($"Target path resolved to: {targetPath}");
        Log($"Target exists: {targetExists}, IsFolder: {isTargetFolder}");

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
                Log($"Arguments: {arguments}");
            }

            Log($"Attempting to launch {targetPath}");
            Process.Start(psi);
            Log($"Successfully launched {targetPath}");
        }
        else
        {
            Log($"Target not found: {targetPath}");
            MessageBox.Show($"Target '{targetPath}' was not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in LaunchItem for {path}: {ex.Message}");
        MessageBox.Show($"Error opening item: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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