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

namespace Desktop_Fences
{
    public static class FenceManager
    {
        private static List<dynamic> _fenceData;
        private static string _jsonFilePath;
        private static dynamic _options;
        private static readonly Dictionary<string, ImageSource> iconCache = new Dictionary<string, ImageSource>();
        private static Dictionary<dynamic, PortalFenceManager> _portalFences = new Dictionary<dynamic, PortalFenceManager>();
        public enum LaunchEffect
        {
            Zoom,        // First effect
            Bounce,      // Scale up and down a few times
            FadeOut,        // Fade out and back in
            SlideUp,     // Slide up and return
            Rotate,      // Spin 360 degrees
                         // Pulse       // Quick scale pulse
            Agitate // Shake back and forth
        }
        // Temporary variable to select effect for testing
      //  private static LaunchEffect _currentEffect = LaunchEffect.Agitate; // Default to Zoom for now


        public static void LoadAndCreateFences(TargetChecker targetChecker)
        {
            string exePath = Assembly.GetEntryAssembly().Location;
            string exeDir = System.IO.Path.GetDirectoryName(exePath);
            _jsonFilePath = System.IO.Path.Combine(exeDir, "fences.json");
           // string optionsPath = System.IO.Path.Combine(exeDir, "options.json");



            SettingsManager.LoadSettings();
            _options = new
            {
                IsSnapEnabled = SettingsManager.IsSnapEnabled,
                TintValue = SettingsManager.TintValue,
                SelectedColor = SettingsManager.SelectedColor,
                IsLogEnabled = SettingsManager.IsLogEnabled,
                singleClickToLaunch = SettingsManager.SingleClickToLaunch,
                LaunchEffect = SettingsManager.LaunchEffect // Default to Zoom
                // LaunchEffect = SettingsManager.LaunchEffect ?? LaunchEffect.Zoom // Default to Zoom if not set
            };
         

            //if (System.IO.File.Exists(optionsPath))
            //{
            //    string optionsContent = System.IO.File.ReadAllText(optionsPath);
            //    _options = JsonConvert.DeserializeObject<dynamic>(optionsContent);
            //}
            //else
            //{
            //    _options = new { IsSnapEnabled = true, TintValue = 70, SelectedColor = "Purple", IsLogEnabled = true, singleClickToLaunch = true };
            //    System.IO.File.WriteAllText(optionsPath, JsonConvert.SerializeObject(_options, Formatting.Indented));
            //}

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

            foreach (dynamic fence in _fenceData)
            {
                CreateFence(fence, targetChecker);
            }

            foreach (var fence in Application.Current.Windows.OfType<NonActivatingWindow>())
            {
                Utility.ApplyTintAndColorToFence(fence);
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
            foreach (var fence in _fenceData)
            {
                if (fence.ItemsType?.ToString() == "Portal")
                {
                    string portalPath = fence.Items?.ToString();
                    if (!string.IsNullOrEmpty(portalPath) && !System.IO.Directory.Exists(portalPath))
                    {
                        fence.IsFolder = true;
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
                        }
                    }
                    fence.Items = items;
                }
            }
            SaveFenceData();
        }

        private static void CreateFence(dynamic fence, TargetChecker targetChecker)
        {
            DockPanel dp = new DockPanel();
            Border cborder = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(100, 0, 0, 0)),
                CornerRadius = new CornerRadius(6),
                Child = dp
            };

            ContextMenu cm = new ContextMenu();
            MenuItem miNF = new MenuItem { Header = "New Fence" };
            MenuItem miNP = new MenuItem { Header = "New Portal Fence" };
            MenuItem miRF = new MenuItem { Header = "Remove Fence" };
            MenuItem miXT = new MenuItem { Header = "Exit" };
            cm.Items.Add(miNF);
            cm.Items.Add(miNP);
            cm.Items.Add(miRF);
            cm.Items.Add(new Separator());
            cm.Items.Add(miXT);

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
                Left = (double)fence.X
            };

            miRF.Click += (s, e) =>
            {
                var result = MessageBox.Show("Are you sure you want to remove this fence?", "Confirm", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
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

                    Log($"Removing fence: {fence.Title}");
                    if (fence.ItemsType?.ToString() == "Data")
                    {
                        var items = fence.Items as JArray;
                        if (items != null)
                        {
                            foreach (var item in items.ToList())
                            {
                                string itemFilePath = item["Filename"]?.ToString();
                                if (!string.IsNullOrEmpty(itemFilePath))
                                {
                                    string shortcutPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Shortcuts", System.IO.Path.GetFileName(itemFilePath));
                                    if (System.IO.File.Exists(shortcutPath))
                                    {
                                        try
                                        {
                                            System.IO.File.Delete(shortcutPath);
                                            Log($"Deleted shortcut: {shortcutPath}");
                                        }
                                        catch (Exception ex)
                                        {
                                            Log($"Failed to delete shortcut {shortcutPath}: {ex.Message}");
                                        }
                                    }
                                    targetChecker.RemoveCheckAction(itemFilePath);
                                }
                            }
                        }
                    }

                    _fenceData.Remove(fence);
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
                Log($"Creating new fence at position: X={absolutePosition.X}, Y={absolutePosition.Y}");
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
                Log($"Creating new portal fence at position: X={absolutePosition.X}, Y={absolutePosition.Y}");
                CreateNewFence("New Portal Fence", "Portal", absolutePosition.X, absolutePosition.Y);
            };

            miXT.Click += (s, e) => Application.Current.Shutdown();

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

                    Log($"Entering edit mode for fence: {fence.Title}");
                    titletb.Text = titlelabel.Content.ToString();
                    titlelabel.Visibility = Visibility.Collapsed;
                    titletb.Visibility = Visibility.Visible;
                    win.ShowActivated = true;
                    win.Activate();
                    Keyboard.Focus(titletb);
                    titletb.SelectAll();
                    Log($"Focus set to title textbox for fence: {fence.Title}");
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
                    Log($"Exited edit mode via Enter, new title for fence: {fence.Title}");
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
                Log($"Exited edit mode via click, new title for fence: {fence.Title}");
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
                                            Log($"Removing icon for {filePath} from fence");
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
                                                        Log($"Deleted shortcut: {shortcutPath}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Log($"Failed to delete shortcut {shortcutPath}: {ex.Message}");
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
                                    Process.Start(new ProcessStartInfo
                                    {
                                        FileName = Utility.GetShortcutTarget(filePath),
                                        UseShellExecute = true,
                                        Verb = "runas"
                                    });
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
                                    Log($"Created unique shortcut: {shortcutName}");
                                }
                                else
                                {
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{System.IO.Path.GetFileNameWithoutExtension(droppedFile)} ({counter}).lnk");
                                        counter++;
                                    }
                                    System.IO.File.Copy(droppedFile, shortcutName, false);
                                    Log($"Copied unique shortcut: {shortcutName}");
                                }

                                dynamic newItem = new System.Dynamic.ExpandoObject();
                                IDictionary<string, object> newItemDict = newItem;
                                newItemDict["Filename"] = shortcutName;
                                newItemDict["IsFolder"] = isFolder; // Ορίζουμε το isFolder με βάση τον στόχο
                                newItemDict["DisplayName"] = System.IO.Path.GetFileNameWithoutExtension(droppedFile);

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
                                                Log($"Removing icon for {shortcutName} from fence");
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
                                                            Log($"Deleted shortcut: {shortcutPath}");
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Log($"Failed to delete shortcut {shortcutPath}: {ex.Message}");
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
                                        Log($"Copied target folder path to clipboard: {folderPath}");
                                    };

                                    miCopyFullPath.Click += (sender, e) =>
                                    {
                                        string targetPath = Utility.GetShortcutTarget(shortcutName);
                                        Clipboard.SetText(targetPath);
                                        Log($"Copied target full path to clipboard: {targetPath}");
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
                                Log($"Added shortcut to Data Fence: {shortcutName}");
                            }
                            else if (fence.ItemsType?.ToString() == "Portal")
                            {
                                // Λογική για Portal Fences (copy) - παραμένει ίδια
                                IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
                                string destinationFolder = fenceDict.ContainsKey("Path") ? fenceDict["Path"]?.ToString() : null;

                                if (string.IsNullOrEmpty(destinationFolder))
                                {
                                    MessageBox.Show($"No destination folder defined for this Portal Fence. Please recreate the fence.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    Log($"No Path defined for Portal Fence: {fence.Title}");
                                    continue;
                                }

                                if (!System.IO.Directory.Exists(destinationFolder))
                                {
                                    MessageBox.Show($"The destination folder '{destinationFolder}' no longer exists. Please update the Portal Fence settings.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    Log($"Destination folder missing for Portal Fence: {destinationFolder}");
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
                                    Log($"Copied file to Portal Fence: {destinationPath}");
                                }
                                else if (System.IO.Directory.Exists(droppedFile))
                                {
                                    CopyDirectory(droppedFile, destinationPath);
                                    Log($"Copied directory to Portal Fence: {destinationPath}");
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

            IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
            SnapManager.AddSnapping(win, fenceDict);
            Utility.ApplyTintAndColorToFence(win);
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

        public static void AddIcon(dynamic icon, WrapPanel wpcont)
        {
            StackPanel sp = new StackPanel { Margin = new Thickness(5), Width = 60 };
            Image ico = new Image { Width = 40, Height = 40, Margin = new Thickness(5) };
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
            string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;

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
            string arguments = null;

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
            string formattedJson = JsonConvert.SerializeObject(_fenceData, Formatting.Indented);
            System.IO.File.WriteAllText(_jsonFilePath, formattedJson);
        }

        private static void CreateNewFence(string title, string itemsType, double x = 20, double y = 20)
        {
            dynamic newFence = new System.Dynamic.ExpandoObject();
            IDictionary<string, object> newFenceDict = newFence;
            newFenceDict["Title"] = title;
            newFenceDict["X"] = x;
            newFenceDict["Y"] = y;
            newFenceDict["Width"] = 230;
            newFenceDict["Height"] = 130;
            newFenceDict["ItemsType"] = itemsType;
            newFenceDict["Items"] = itemsType == "Portal" ? "" : new JArray();

            if (itemsType == "Portal")
            {
                using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    dialog.Description = "Select the folder to monitor for this Portal Fence";
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        newFenceDict["Path"] = dialog.SelectedPath; // Ορίζουμε το Path κατά τη δημιουργία
                    }
                    else
                    {
                        return; // Αν ακυρώσει, δεν δημιουργούμε το fence
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
                    btn.Click += (s, e) =>
                    {
                        var sourceItems = sourceFence.Items as JArray;
                        var destItems = fence.Items as JArray ?? new JArray();
                        if (sourceItems != null)
                        {
                            IDictionary<string, object> itemDict = item is IDictionary<string, object> dict ? dict : ((JObject)item).ToObject<IDictionary<string, object>>();
                            string filename = itemDict.ContainsKey("Filename") ? itemDict["Filename"].ToString() : "Unknown";

                            Log($"Moving item {filename} from {sourceFence.Title} to {fence.Title}");
                            sourceItems.Remove(item);
                            destItems.Add(item);
                            fence.Items = destItems;
                            SaveFenceData();
                            moveWindow.Close();

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

        private static void EditItem(dynamic icon, dynamic fence, NonActivatingWindow win)
        {
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            string displayName = iconDict.ContainsKey("DisplayName") ? (string)iconDict["DisplayName"] : System.IO.Path.GetFileNameWithoutExtension(filePath);
            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";

            bool isLogEnabled = _options.IsLogEnabled ?? true;
            void Log(string message)
            {
                if (isLogEnabled)
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                }
            }

            if (!isShortcut)
            {
                MessageBox.Show("Edit is only available for shortcuts.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var editWindow = new EditShortcutWindow(filePath, displayName);
            if (editWindow.ShowDialog() == true)
            {
                iconDict["DisplayName"] = editWindow.NewDisplayName;
                Log($"Updated DisplayName for {filePath} to {editWindow.NewDisplayName}");

                var items = fence.Items as JArray;
                if (items != null)
                {
                    var itemToUpdate = items.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                    if (itemToUpdate != null)
                    {
                        itemToUpdate["DisplayName"] = editWindow.NewDisplayName;
                        SaveFenceData();
                        Log($"Fence data updated for {filePath}");
                    }
                }
                if (win != null)
                {
                    var wpcont = ((win.Content as Border)?.Child as DockPanel)?.Children.Cast<object>()
                        .OfType<ScrollViewer>().FirstOrDefault()?.Content as WrapPanel;
                    if (wpcont != null)
                    {
                        var sp = wpcont.Children.Cast<object>()
                            .OfType<StackPanel>()
                            .FirstOrDefault(s => (s.Tag as string) == filePath);
                        if (sp != null)
                        {
                            var lbl = sp.Children.Cast<object>().OfType<TextBlock>().FirstOrDefault();
                            var ico = sp.Children.Cast<object>().OfType<Image>().FirstOrDefault();
                            if (lbl != null && ico != null)
                            {
                                string newText = editWindow.NewDisplayName.Length > 20 ? editWindow.NewDisplayName.Substring(0, 20) + "..." : editWindow.NewDisplayName;
                                lbl.Text = newText;

                                WshShell shell = new WshShell();
                                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                                ImageSource shortcutIcon = null;

                                if (!string.IsNullOrEmpty(shortcut.IconLocation))
                                {
                                    string[] iconParts = shortcut.IconLocation.Split(',');
                                    string iconPath = iconParts[0];
                                    if (System.IO.File.Exists(iconPath))
                                    {
                                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(iconPath).ToImageSource();
                                        Log($"Applied custom icon from IconLocation for {filePath}: {iconPath}");
                                    }
                                }

                                if (shortcutIcon == null)
                                {
                                    string targetPath = shortcut.TargetPath;
                                    if (System.IO.File.Exists(targetPath))
                                    {
                                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                                        Log($"No custom icon, using target icon for {filePath}: {targetPath}");
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
                            }
                        }
                    }
                }
                //if (win != null)
                //{
                //    var wpcont = ((win.Content as Border)?.Child as DockPanel)?.Children.Cast<object>()
                //        .OfType<ScrollViewer>().FirstOrDefault()?.Content as WrapPanel;
                //    if (wpcont != null)
                //    {
                //        var sp = wpcont.Children.Cast<object>()
                //            .OfType<StackPanel>()
                //            .FirstOrDefault(s => (s.Tag as string) == filePath);
                //        if (sp != null)
                //        {
                //            var lbl = sp.Children.Cast<object>().OfType<TextBlock>().FirstOrDefault();
                //            var ico = sp.Children.Cast<object>().OfType<Image>().FirstOrDefault();
                //            if (lbl != null && ico != null)
                //            {
                //                string newText = editWindow.NewDisplayName.Length > 20 ? editWindow.NewDisplayName.Substring(0, 20) + "..." : editWindow.NewDisplayName;
                //                lbl.Text = newText;

                //                WshShell shell = new WshShell();
                //                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                //                ImageSource shortcutIcon = null;

                //                if (!string.IsNullOrEmpty(shortcut.IconLocation))
                //                {
                //                    string[] iconParts = shortcut.IconLocation.Split(',');
                //                    string iconPath = iconParts[0];
                //                    if (System.IO.File.Exists(iconPath))
                //                    {
                //                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(iconPath).ToImageSource();
                //                    }
                //                }

                //                if (shortcutIcon == null)
                //                {
                //                    string targetPath = shortcut.TargetPath;
                //                    if (System.IO.File.Exists(targetPath))
                //                    {
                //                        shortcutIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                //                    }
                //                    else if (System.IO.Directory.Exists(targetPath))
                //                    {
                //                        shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                //                    }
                //                }

                //                if (shortcutIcon == null)
                //                {
                //                    shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                //                }

                //                ico.Source = shortcutIcon;
                //                iconCache[filePath] = shortcutIcon;
                //            }
                //        }
                //    }
                //}
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
                                Log($"Using custom IconLocation for {filePath}: {iconPath}");
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
                            Log($"Target exists, updating to folder-White.png for {filePath}");
                        }
                        else
                        {
                            try
                            {
                                newIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                                Log($"Target exists, updating to file icon for {filePath}");
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
                    Log($"Icon updated for {filePath}");
                }
            });

            }
            else
            {
                Log("Application.Current is null, cannot update icon.");
            }
        }

        //private static void UpdateIcon(StackPanel sp, string filePath, bool isFolder)
        //{
        //    bool isLogEnabled = _options.IsLogEnabled ?? true;
        //    void Log(string message)
        //    {
        //        if (isLogEnabled)
        //        {
        //            string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
        //            System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
        //        }
        //    }

        //    if (Application.Current != null)
        //    {
        //        Application.Current.Dispatcher.Invoke(() =>
        //        {
        //            Image ico = sp.Children.OfType<Image>().FirstOrDefault();
        //            if (ico == null) return;

        //            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
        //            string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;
        //            bool targetExists = System.IO.File.Exists(targetPath) || System.IO.Directory.Exists(targetPath);
        //            bool isTargetFolder = System.IO.Directory.Exists(targetPath);

        //            ImageSource newIcon = null;

        //            if (targetExists)
        //            {
        //                if (isTargetFolder)
        //                {
        //                    newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
        //                    Log($"Target exists, updating to folder-White.png for {filePath}");
        //                }
        //                else
        //                {
        //                    try
        //                    {
        //                        newIcon = System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
        //                        Log($"Target exists, updating to file icon for {filePath}");
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
        //                        Log($"Failed to extract icon for {filePath}: {ex.Message}");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                newIcon = isFolder
        //                    ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"))
        //                    : new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
        //                Log($"Target missing, updating to {(isFolder ? "folder-WhiteX.png" : "file-WhiteX.png")} for {filePath}");
        //            }

        //            if (ico.Source != newIcon)
        //            {
        //                ico.Source = newIcon;
        //            }
        //        });
        //    }
        //    else
        //    {
        //        Log("Application.Current is null, cannot update icon.");
        //    }
        //}


        public static void UpdateOptionsAndClickEvents()
        {
            bool isLogEnabled = SettingsManager.IsLogEnabled;

            void Log(string message)
            {
                if (isLogEnabled)
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                }
            }

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

            void Log(string message)
            {
                if (isLogEnabled)
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                }
            }

            Log($"Attaching handler for {path}, initial singleClickToLaunch={_options.singleClickToLaunch}");

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

                bool singleClickToLaunch = _options.singleClickToLaunch ?? true; // Check current value at execution
                Log($"MouseDown on {path}, ClickCount={e.ClickCount}, singleClickToLaunch={singleClickToLaunch}");
                if (singleClickToLaunch && e.ClickCount == 1)
                {
                    Log($"Single click launching {path}");
                    Log($"Single click launching {path} with arguments: {arguments ?? "none"}");

                    LaunchItem(sp, path, isFolder, arguments);
                    e.Handled = true;
                }
                else if (!singleClickToLaunch && e.ClickCount == 2)
                {
                    Log($"Double click launching {path}");
                    Log($"Double click launching {path} with arguments: {arguments ?? "none"}");
                    LaunchItem(sp, path, isFolder, arguments);
                    e.Handled = true;
                }
            }

            sp.MouseLeftButtonDown += MouseDownHandler;
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
                //  Log($"LaunchItem called for {path}");
                Log($"LaunchItem called for {path} with effect {_options.LaunchEffect}");

                //// Animation
                //if (sp.RenderTransform == null || !(sp.RenderTransform is ScaleTransform))
                //{
                //    sp.RenderTransform = new ScaleTransform(1, 1);
                //    sp.RenderTransformOrigin = new Point(0.5, 0.5);
                //}
                //var scale = new DoubleAnimation(1, 1.2, TimeSpan.FromSeconds(0.1)) { AutoReverse = true };
                //var transform = (ScaleTransform)sp.RenderTransform;
                //transform.BeginAnimation(ScaleTransform.ScaleXProperty, scale);
                //transform.BeginAnimation(ScaleTransform.ScaleYProperty, scale);

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
                switch (_options.LaunchEffect)
                {
                    case LaunchEffect.Zoom:
                        // Existing zoom effect: Scale up to 1.2 and back
                        var zoomScale = new DoubleAnimation(1, 1.2, TimeSpan.FromSeconds(0.1)) { AutoReverse = true };
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, zoomScale);
                        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, zoomScale);
                        break;

                    case LaunchEffect.Bounce:
                        // Bounce: Scale up and down 3 times
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
                        // Fade out to 0 and back to 1
                        var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.2))
                        {
                            AutoReverse = true
                        };
                        sp.BeginAnimation(UIElement.OpacityProperty, fade);
                        break;

                    case LaunchEffect.SlideUp:
                        // Slide up 20 units and back
                        var slideUp = new DoubleAnimation(0, -20, TimeSpan.FromSeconds(0.2))
                        {
                            AutoReverse = true
                        };
                        translateTransform.BeginAnimation(TranslateTransform.YProperty, slideUp);
                        break;

                    case LaunchEffect.Rotate:
                        // Spin 360 degrees
                        var rotate = new DoubleAnimation(0, 360, TimeSpan.FromSeconds(0.4))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };
                        rotateTransform.BeginAnimation(RotateTransform.AngleProperty, rotate);
                        break;
                    //Puse effect removed
                    //case LaunchEffect.Pulse:
                    //    // Quick scale pulse: Slightly larger then back
                    //    var pulseScale = new DoubleAnimation(1, 1.15, TimeSpan.FromSeconds(0.1))
                    //    {
                    //        AutoReverse = true,
                    //        RepeatBehavior = new RepeatBehavior(2) // Two quick pulses
                    //    };
                    //    scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, pulseScale);
                    //    scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, pulseScale);
                    //    break;

                    case LaunchEffect.Agitate:
                        // Agitate: Quick left-right bounce 7 times
                        var agitateTranslate = new DoubleAnimationUsingKeyFrames
                        {
                            Duration = TimeSpan.FromSeconds(0.7) // 0.1s per bounce, 7 bounces
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
    }
}