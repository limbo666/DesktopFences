using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;

namespace Desktop_Fences
{
    public static class OptionsFormManager
    {
        private static int _lastSelectedTabIndex = 0;
        private static TabControl _tabControl;
        private static Window _optionsWindow;
        private static Color _userAccentColor;

        // Colors for tabs
        private static readonly Color ColorStyle = Color.FromRgb(128, 0, 128); // Purple
        private static readonly Color ColorTools = Color.FromRgb(34, 139, 34); // Green
        private static readonly Color ColorLookDeeper = Color.FromRgb(220, 53, 69); // Red

        public static void ShowOptionsForm()
        {
            try
            {
                _userAccentColor = Utility.GetColorFromName(SettingsManager.SelectedColor);

                _optionsWindow = new Window
                {
                    Title = "Desktop Fences + Options",
                    Width = 800,
                    Height = 850,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStyle = WindowStyle.None,
                    Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                    AllowsTransparency = true
                };

                try
                {
                    _optionsWindow.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions());
                }
                catch { }

                Border mainBorder = new Border
                {
                    Background = Brushes.White,
                    Margin = new Thickness(8),
                    Effect = new DropShadowEffect { Color = Colors.Black, Direction = 270, ShadowDepth = 2, BlurRadius = 10, Opacity = 0.2 }
                };

                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(40) }); // Header
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(60) }); // Footer
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(80) }); // Donation

                // Header
                Border headerBorder = new Border { Background = new SolidColorBrush(_userAccentColor), Height = 40 };
                Grid.SetRow(headerBorder, 0);

                Grid headerGrid = new Grid();
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });

                TextBlock titleBlock = new TextBlock
                {
                    Text = "Options",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 16,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(20, 0, 0, 0)
                };

                Button closeButton = new Button
                {
                    Content = "✕",
                    Width = 32,
                    Height = 32,
                    Foreground = Brushes.White,
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand
                };
                closeButton.Click += (s, e) => _optionsWindow.Close();

                Grid.SetColumn(titleBlock, 0); headerGrid.Children.Add(titleBlock);
                Grid.SetColumn(closeButton, 1); headerGrid.Children.Add(closeButton);
                headerBorder.Child = headerGrid;
                headerBorder.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) _optionsWindow.DragMove(); };

                CreateTabContent(mainGrid);
                CreateFooter(mainGrid);
                CreateDonationSection(mainGrid);

                mainGrid.Children.Add(headerBorder);
                mainBorder.Child = mainGrid;
                _optionsWindow.Content = mainBorder;
                _optionsWindow.KeyDown += (s, e) => { if (e.Key == Key.Enter) SaveOptions(); else if (e.Key == Key.Escape) _optionsWindow.Close(); };

                _optionsWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing Options: {ex.Message}");
            }
        }

        private static void CreateTabContent(Grid mainGrid)
        {
            Grid contentGrid = new Grid();
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) });
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Grid.SetRow(contentGrid, 1);

            StackPanel tabPanel = new StackPanel { Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)), Margin = new Thickness(0, 20, 0, 0) };
            Border contentBorder = new Border { Background = Brushes.White, BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1, 0, 0, 0), Padding = new Thickness(20), Margin = new Thickness(0, 20, 0, 0) };

            _tabControl = new TabControl { Background = Brushes.Transparent, BorderThickness = new Thickness(0) };
            var template = new ControlTemplate(typeof(TabControl));
            template.VisualTree = new FrameworkElementFactory(typeof(ContentPresenter));
            ((FrameworkElementFactory)template.VisualTree).SetValue(ContentPresenter.ContentSourceProperty, "SelectedContent");
            _tabControl.Template = template;

            CreateGeneralTab();
            CreateStyleTab();
            CreateToolsTab();
            CreateLookDeeperTab();

            _tabControl.SelectedIndex = _lastSelectedTabIndex;
            CreateTabButton(tabPanel, "General", 0, _lastSelectedTabIndex == 0);
            CreateTabButton(tabPanel, "Style", 1, _lastSelectedTabIndex == 1);
            CreateTabButton(tabPanel, "Tools", 2, _lastSelectedTabIndex == 2);
            CreateTabButton(tabPanel, "Look Deeper", 3, _lastSelectedTabIndex == 3);

            contentBorder.Child = _tabControl;
            Grid.SetColumn(tabPanel, 0); contentGrid.Children.Add(tabPanel);
            Grid.SetColumn(contentBorder, 1); contentGrid.Children.Add(contentBorder);
            mainGrid.Children.Add(contentGrid);
        }

        private static void CreateTabButton(StackPanel parent, string title, int tabIndex, bool isSelected)
        {
            Button tabButton = new Button
            {
                Content = title,
                Height = 40,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Padding = new Thickness(20, 0, 0, 0),
                Margin = new Thickness(0, 0, 0, 2)
            };
            SetTabButtonColors(tabButton, title, isSelected);
            tabButton.Click += (s, e) => SelectTab(tabIndex, tabButton);
            tabButton.MouseEnter += (s, e) => { if (_tabControl.SelectedIndex != tabIndex) SetTabButtonColors(tabButton, title, false, true); };
            tabButton.MouseLeave += (s, e) => { if (_tabControl.SelectedIndex != tabIndex) SetTabButtonColors(tabButton, title, false, false); };
            parent.Children.Add(tabButton);
        }

        private static void SetTabButtonColors(Button button, string title, bool isSelected, bool isHover = false)
        {
            Color activeColor = title switch { "Style" => ColorStyle, "Tools" => ColorTools, "Look Deeper" => ColorLookDeeper, _ => _userAccentColor };
            if (isSelected) { button.Background = new SolidColorBrush(activeColor); button.Foreground = Brushes.White; }
            else if (isHover) { button.Background = new SolidColorBrush(Color.FromRgb((byte)(activeColor.R + 40), (byte)(activeColor.G + 40), (byte)(activeColor.B + 40))); button.Foreground = Brushes.White; }
            else { button.Background = new SolidColorBrush(Color.FromRgb(200, 200, 200)); button.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 60)); }
        }

        private static void SelectTab(int tabIndex, Button selectedButton)
        {
            _lastSelectedTabIndex = tabIndex;
            _tabControl.SelectedIndex = tabIndex;
            StackPanel tabPanel = (StackPanel)selectedButton.Parent;
            for (int i = 0; i < tabPanel.Children.Count; i++) if (tabPanel.Children[i] is Button btn) SetTabButtonColors(btn, btn.Content.ToString(), i == tabIndex);
        }

        // --- Tabs ---
        private static void CreateGeneralTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();
            CreateSectionHeader(c, "Startup", _userAccentColor);
            CreateCheckBox(c, "Start with Windows", "StartWithWindows", TrayManager.IsStartWithWindows);
            CreateSectionHeader(c, "Selections", _userAccentColor);
            CreateCheckBox(c, "Single Click to Launch", "SingleClickToLaunch", SettingsManager.SingleClickToLaunch);
            CreateCheckBox(c, "Enable Snap Near Fences", "EnableSnapNearFences", SettingsManager.IsSnapEnabled);
            CreateCheckBox(c, "Enable Dimension Snap", "EnableDimensionSnap", SettingsManager.EnableDimensionSnap);
            CreateCheckBox(c, "Enable Tray Icon", "EnableTrayIcon", SettingsManager.ShowInTray);
            CreateCheckBox(c, "Use Recycle Bin on Portal Fences 'Delete item' command", "UseRecycleBin", SettingsManager.UseRecycleBin);
            t.Content = new ScrollViewer { Content = c, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            _tabControl.Items.Add(t);
        }

        private static void CreateStyleTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();
            CreateSectionHeader(c, "Choices", ColorStyle);
            CreateCheckBox(c, "Enable Portal Fences Watermark", "EnablePortalWatermark", SettingsManager.ShowBackgroundImageOnPortalFences);
            var n = CreateCheckBoxReturn(c, "Enable Note Fences Watermark (Coming Soon)", "EnableNoteWatermark", false);
            n.IsEnabled = false; n.Foreground = Brushes.Gray;
            CreateCheckBox(c, "Disable Fence Scrollbars", "DisableFenceScrollbars", SettingsManager.DisableFenceScrollbars);
            CreateCheckBox(c, "Enable Sounds", "EnableSounds", SettingsManager.EnableSounds);

            CreateSectionHeader(c, "Appearance", ColorStyle);
            CreateSliderControl(c, "Fence Tint", "TintSlider", SettingsManager.TintValue);
            CreateSliderControl(c, "Menu Tint", "MenuTintSlider", SettingsManager.MenuTintValue);
            CreateColorComboBox(c);
            CreateLaunchEffectComboBox(c);

            CreateSectionHeader(c, "Icons", ColorStyle);
            c.Children.Add(new TextBlock { Text = "Menu Icon", FontWeight = FontWeights.SemiBold, Margin = new Thickness(15, 5, 0, 5) });
            CreateIconRadioButtonGroup(c, "MenuIconGroup", new Dictionary<string, int> { { "♥", 0 }, { "☰", 1 }, { "≣", 2 }, { "𓃑", 3 } }, SettingsManager.MenuIcon);
            c.Children.Add(new TextBlock { Text = "Lock Icon", FontWeight = FontWeights.SemiBold, Margin = new Thickness(15, 10, 0, 5) });
            CreateIconRadioButtonGroup(c, "LockIconGroup", new Dictionary<string, int> { { "🛡️", 0 }, { "🔑", 1 }, { "🔐", 2 }, { "🔒", 3 } }, SettingsManager.LockIcon);

            t.Content = new ScrollViewer { Content = c, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            _tabControl.Items.Add(t);
        }

        private static void CreateToolsTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();
            CreateSectionHeader(c, "Tools", ColorTools);

            Grid g = new Grid { Margin = new Thickness(0, 10, 0, 0) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(15) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) });
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(15) });
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) });

            Button b1 = CreateStyledButton("Backup", ColorTools); b1.Click += (s, e) => BackupManager.BackupData();
            Button b2 = CreateStyledButton("Restore...", Color.FromRgb(255, 152, 0)); b2.Click += (s, e) => RestoreBackup();
            Button b3 = CreateStyledButton("Open Backups Folder", Color.FromRgb(0, 123, 191)); b3.Click += (s, e) => OpenBackupsFolder();
            Grid.SetRow(b1, 0); Grid.SetColumn(b1, 0);
            Grid.SetRow(b2, 0); Grid.SetColumn(b2, 2);
            Grid.SetRow(b3, 2); Grid.SetColumn(b3, 0); Grid.SetColumnSpan(b3, 3);
            g.Children.Add(b1); g.Children.Add(b2); g.Children.Add(b3);
            c.Children.Add(g);

            CreateCheckBox(c, "Automatic Backup (Daily)", "EnableAutoBackup", SettingsManager.EnableAutoBackup);

            CreateSectionHeader(c, "Reset", Colors.Red);
            Button r1 = CreateStyledButton("Reset Styles", Color.FromRgb(108, 117, 125));
            r1.Width = 255; r1.Height = 45; r1.Margin = new Thickness(0, 0, 0, 15);
            r1.Click += (s, e) => { if (MessageBoxesManager.ShowCustomYesNoMessageBox("Reset all visual customizations?", "Reset")) { FenceManager.ResetAllCustomizations(); _optionsWindow.Close(); } };

            Button r2 = CreateStyledButton("Clear All Data", Color.FromRgb(220, 53, 69));
            r2.Width = 255; r2.Height = 45;
            r2.Click += (s, e) => PerformFullFactoryReset();

            StackPanel rs = new StackPanel { HorizontalAlignment = HorizontalAlignment.Left };
            rs.Children.Add(r1); rs.Children.Add(r2);
            c.Children.Add(rs);

            t.Content = c;
            _tabControl.Items.Add(t);
        }

        private static void CreateLookDeeperTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();
            CreateSectionHeader(c, "Log", ColorLookDeeper);
            CreateCheckBox(c, "Enable logging", "EnableLogging", SettingsManager.IsLogEnabled);
            Button b = CreateStyledButton("Open Log", ColorLookDeeper); b.Width = 100; b.Height = 25; b.HorizontalAlignment = HorizontalAlignment.Left;
            b.Click += (s, e) => OpenLogFile();
            c.Children.Add(b);

            CreateSectionHeader(c, "Log configuration", ColorLookDeeper);
            CreateLogLevelComboBox(c);
            CreateSectionHeader(c, "Log Categories", ColorLookDeeper);

            // This method creates checkboxes for all Enums (except Error now)
            CreateLogCategoryCheckBoxes(c);

            t.Content = new ScrollViewer { Content = c, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            _tabControl.Items.Add(t);
        }

        // --- Helpers ---
        private static void CreateSectionHeader(StackPanel p, string t, Color c) => p.Children.Add(new TextBlock { Text = t, FontFamily = new FontFamily("Segoe UI"), FontSize = 16, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(c), Margin = new Thickness(0, 10, 0, 15) });
        private static void CreateCheckBox(StackPanel p, string t, string n, bool c) => p.Children.Add(new CheckBox { Name = n, Content = t, IsChecked = c, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, Margin = new Thickness(15, 8, 0, 8) });
        private static CheckBox CreateCheckBoxReturn(StackPanel p, string t, string n, bool c) { var cb = new CheckBox { Name = n, Content = t, IsChecked = c, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, Margin = new Thickness(15, 8, 0, 8) }; p.Children.Add(cb); return cb; }

        private static void CreateSliderControl(StackPanel p, string l, string n, int v)
        {
            Grid g = new Grid { Margin = new Thickness(0, 5, 0, 5) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(50) });
            TextBlock lbl = new TextBlock { Text = l, FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 10, 0) };
            Slider sl = new Slider { Name = n, Minimum = 1, Maximum = 100, Value = v, TickFrequency = 1, IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center };
            TextBlock val = new TextBlock { Text = v.ToString(), FontSize = 13, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(15, 0, 0, 0) };
            sl.ValueChanged += (s, e) => val.Text = ((int)e.NewValue).ToString();
            Grid.SetColumn(lbl, 0); Grid.SetColumn(sl, 1); Grid.SetColumn(val, 2);
            g.Children.Add(lbl); g.Children.Add(sl); g.Children.Add(val);
            p.Children.Add(g);
        }

        private static void CreateIconRadioButtonGroup(StackPanel p, string gName, Dictionary<string, int> icons, int sel)
        {
            StackPanel sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(15, 5, 0, 15), Tag = gName };
            foreach (var i in icons) sp.Children.Add(new RadioButton { Content = i.Key, Tag = i.Value, GroupName = gName, IsChecked = i.Value == sel, Margin = new Thickness(0, 0, 15, 0), FontSize = 16, FontFamily = new FontFamily("Segoe UI Symbol") });
            p.Children.Add(sp);
        }

        private static void CreateColorComboBox(StackPanel p)
        {
            Grid g = new Grid { Margin = new Thickness(0, 10, 0, 10) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            g.Children.Add(new TextBlock { Text = "Color", FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 10, 0) });
            ComboBox cb = new ComboBox { Name = "ColorComboBox", Height = 25, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            foreach (string c in new[] { "Gray", "Black", "White", "Beige", "Green", "Purple", "Fuchsia", "Yellow", "Orange", "Red", "Blue", "Bismark" }) cb.Items.Add(c);
            cb.SelectedItem = SettingsManager.SelectedColor;
            Grid.SetColumn(cb, 1); g.Children.Add(cb); p.Children.Add(g);
        }

        private static void CreateLaunchEffectComboBox(StackPanel p)
        {
            Grid g = new Grid { Margin = new Thickness(0, 10, 0, 10) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            g.Children.Add(new TextBlock { Text = "Effect", FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 10, 0) });
            ComboBox cb = new ComboBox { Name = "LaunchEffectComboBox", Height = 25, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            foreach (string e in new[] { "Zoom", "Bounce", "FadeOut", "SlideUp", "Rotate", "Agitate", "GrowAndFly", "Pulse", "Elastic", "Flip3D", "Spiral", "Shockwave", "Matrix", "Supernova", "Teleport" }) cb.Items.Add(e);
            cb.SelectedIndex = (int)SettingsManager.LaunchEffect;
            Grid.SetColumn(cb, 1); g.Children.Add(cb); p.Children.Add(g);
        }

        private static Button CreateStyledButton(string t, Color c) => new Button { Content = t, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, FontWeight = FontWeights.Bold, Background = new SolidColorBrush(c), Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand, HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch };

        private static void CreateLogLevelComboBox(StackPanel p)
        {
            Grid g = new Grid { Margin = new Thickness(0, 10, 0, 10) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            g.Children.Add(new TextBlock { Text = "Minimum Log Level", FontSize = 13, VerticalAlignment = VerticalAlignment.Center });
            ComboBox cb = new ComboBox { Name = "LogLevelComboBox", Height = 25, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            foreach (var l in new[] { "Debug", "Info", "Warn", "Error" }) cb.Items.Add(l);
            cb.SelectedItem = SettingsManager.MinLogLevel.ToString();
            Grid.SetColumn(cb, 1); g.Children.Add(cb); p.Children.Add(g);
        }

        // --- Log Categories (Optimized & Fixed) ---
        private static void CreateLogCategoryCheckBoxes(StackPanel p)
        {
            Grid g = new Grid { Name = "LogCategoryGrid" };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            StackPanel l = new StackPanel(); StackPanel r = new StackPanel();

            // FIX: Filter out the "Error" category from UI, as it's a Level, not a Category
            var cats = Enum.GetValues(typeof(LogManager.LogCategory))
                .Cast<LogManager.LogCategory>()
                .Where(c => c != LogManager.LogCategory.Error) // Hide Error
                .ToList();

            int half = (cats.Count + 1) / 2;

            for (int i = 0; i < cats.Count; i++)
            {
                var cb = new CheckBox { Content = cats[i].ToString(), Tag = cats[i], IsChecked = SettingsManager.EnabledLogCategories.Contains(cats[i]), FontSize = 13, Margin = new Thickness(15, 8, 0, 8) };
                if (i < half) l.Children.Add(cb); else r.Children.Add(cb);
            }

            Grid.SetColumn(l, 0); Grid.SetColumn(r, 1);
            g.Children.Add(l); g.Children.Add(r);
            p.Children.Add(g);
        }

        // --- SAVING ---
        private static void SaveOptions()
        {
            try
            {
                bool tempPortalImageState = SettingsManager.ShowBackgroundImageOnPortalFences;
                bool newPortalWatermarkState = false;
                bool newShowInTrayState = false;

                // 1. General
                var generalContent = (StackPanel)((ScrollViewer)((TabItem)_tabControl.Items[0]).Content).Content;
                foreach (var child in generalContent.Children)
                {
                    if (child is CheckBox cb)
                    {
                        if (cb.Name == "StartWithWindows" && cb.IsChecked != TrayManager.IsStartWithWindows) TrayManager.Instance?.ToggleStartWithWindows(cb.IsChecked == true);
                        if (cb.Name == "SingleClickToLaunch") SettingsManager.SingleClickToLaunch = cb.IsChecked == true;
                        if (cb.Name == "EnableSnapNearFences") SettingsManager.IsSnapEnabled = cb.IsChecked == true;
                        if (cb.Name == "EnableDimensionSnap") SettingsManager.EnableDimensionSnap = cb.IsChecked == true;
                        if (cb.Name == "UseRecycleBin") SettingsManager.UseRecycleBin = cb.IsChecked == true;
                        if (cb.Name == "EnableTrayIcon")
                        {
                            newShowInTrayState = cb.IsChecked == true;
                            SettingsManager.ShowInTray = newShowInTrayState;
                            if (TrayManager.Instance != null) TrayManager.Instance.GetType().GetField("Showintray", BindingFlags.NonPublic | BindingFlags.Instance)?.SetValue(TrayManager.Instance, newShowInTrayState);
                        }
                    }
                }

                // 2. Style
                var styleContent = (StackPanel)((ScrollViewer)((TabItem)_tabControl.Items[1]).Content).Content;
                foreach (var child in styleContent.Children)
                {
                    if (child is CheckBox cb)
                    {
                        if (cb.Name == "EnablePortalWatermark") { newPortalWatermarkState = cb.IsChecked == true; SettingsManager.ShowBackgroundImageOnPortalFences = newPortalWatermarkState; }
                        if (cb.Name == "DisableFenceScrollbars") SettingsManager.DisableFenceScrollbars = cb.IsChecked == true;
                        if (cb.Name == "EnableSounds") SettingsManager.EnableSounds = cb.IsChecked == true;
                    }
                    else if (child is Grid g)
                    {
                        var tint = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "TintSlider"); if (tint != null) SettingsManager.TintValue = (int)tint.Value;
                        var mtint = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "MenuTintSlider"); if (mtint != null) SettingsManager.MenuTintValue = (int)mtint.Value;
                        var col = g.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "ColorComboBox"); if (col?.SelectedItem != null) SettingsManager.SelectedColor = col.SelectedItem.ToString();
                        var eff = g.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "LaunchEffectComboBox"); if (eff != null) SettingsManager.LaunchEffect = (LaunchEffectsManager.LaunchEffect)eff.SelectedIndex;
                    }
                    else if (child is StackPanel sp)
                    {
                        if (sp.Tag?.ToString() == "MenuIconGroup") foreach (RadioButton rb in sp.Children.OfType<RadioButton>()) if (rb.IsChecked == true) SettingsManager.MenuIcon = (int)rb.Tag;
                        if (sp.Tag?.ToString() == "LockIconGroup") foreach (RadioButton rb in sp.Children.OfType<RadioButton>()) if (rb.IsChecked == true) SettingsManager.LockIcon = (int)rb.Tag;
                    }
                }

                // 3. Tools
                var toolsContent = (StackPanel)((TabItem)_tabControl.Items[2]).Content;
                foreach (var child in toolsContent.Children) if (child is CheckBox cb && cb.Name == "EnableAutoBackup") SettingsManager.EnableAutoBackup = cb.IsChecked == true;

                // 4. Look Deeper (Logs)
                var logContent = (StackPanel)((ScrollViewer)((TabItem)_tabControl.Items[3]).Content).Content;
                var newEnabledCategories = new List<LogManager.LogCategory>();

                // FIX: Force enable the hidden "Error" category so existing log calls don't break.
                // It will only be filtered by the "Minimum Log Level" dropdown now.
                newEnabledCategories.Add(LogManager.LogCategory.Error);

                foreach (var child in logContent.Children)
                {
                    if (child is CheckBox cb)
                    {
                        if (cb.Name == "EnableLogging") SettingsManager.IsLogEnabled = cb.IsChecked == true;
                    }
                    else if (child is Grid g)
                    {
                        var lvl = g.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "LogLevelComboBox");
                        if (lvl?.SelectedItem != null && Enum.TryParse<LogManager.LogLevel>(lvl.SelectedItem.ToString(), out var ll)) SettingsManager.SetMinLogLevel(ll);

                        if (g.Name == "LogCategoryGrid")
                        {
                            foreach (var stack in g.Children.OfType<StackPanel>())
                            {
                                // FIX: Renamed inner variable 'catBox' to prevent conflict
                                foreach (var catBox in stack.Children.OfType<CheckBox>())
                                {
                                    if (catBox.IsChecked == true && catBox.Tag is LogManager.LogCategory cat)
                                    {
                                        newEnabledCategories.Add(cat);
                                        // Sync the boolean for background validation
                                        if (cat == LogManager.LogCategory.BackgroundValidation)
                                            SettingsManager.EnableBackgroundValidationLogging = true;
                                    }
                                }
                            }
                        }
                    }
                }

                if (!newEnabledCategories.Contains(LogManager.LogCategory.BackgroundValidation))
                    SettingsManager.EnableBackgroundValidationLogging = false;

                SettingsManager.SetEnabledLogCategories(newEnabledCategories);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.Settings, "Options saved successfully");

                if (tempPortalImageState != newPortalWatermarkState) TrayManager.reloadAllFences();
                TrayManager.Instance?.UpdateTrayIcon();
                Utility.UpdateFenceVisuals();

                _optionsWindow.Close();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings, $"Error saving options: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error: {ex.Message}", "Save Error");
            }
        }

        private static void CreateFooter(Grid mainGrid)
        {
            Border f = new Border { Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(0, 1, 0, 0), Padding = new Thickness(20, 8, 20, 8) };
            Grid.SetRow(f, 2);
            StackPanel sp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

            Button c = new Button { Content = "Cancel", Width = 100, Height = 34, FontWeight = FontWeights.Bold, Background = Brushes.White, BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Margin = new Thickness(0, 0, 10, 0), Cursor = Cursors.Hand };
            c.Click += (s, e) => _optionsWindow.Close();

            Button sv = new Button { Content = "Save", Width = 100, Height = 34, FontWeight = FontWeights.Bold, Background = new SolidColorBrush(_userAccentColor), Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
            sv.Click += (s, e) => SaveOptions();

            sp.Children.Add(c); sp.Children.Add(sv); f.Child = sp; mainGrid.Children.Add(f);
        }

        private static void CreateDonationSection(Grid mainGrid)
        {
            Border d = new Border { Background = new SolidColorBrush(Color.FromRgb(255, 248, 225)), BorderBrush = new SolidColorBrush(Color.FromRgb(255, 193, 7)), BorderThickness = new Thickness(0, 1, 0, 0), Padding = new Thickness(20) };
            Grid.SetRow(d, 3);
            StackPanel sp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
            sp.Children.Add(new TextBlock { Text = "Support the Maintenance and Enhancement of This Project by Donating", FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(102, 77, 3)), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 15, 0) });
            Button b = new Button { Content = "♥ Donate via PayPal", FontSize = 14, Background = new SolidColorBrush(Color.FromRgb(255, 193, 7)), Foreground = Brushes.White, BorderThickness = new Thickness(0), Padding = new Thickness(15, 6, 15, 6), Cursor = Cursors.Hand };
            b.Click += (s, e) => { try { Process.Start(new ProcessStartInfo { FileName = "https://www.paypal.com/donate/?hosted_button_id=M8H4M4R763RBE", UseShellExecute = true }); } catch { } };
            sp.Children.Add(b); d.Child = sp; mainGrid.Children.Add(d);
        }

        private static void RestoreBackup()
        {
            try
            {
                string r = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string b = System.IO.Path.Combine(r, "Backups");
                using (var d = new System.Windows.Forms.FolderBrowserDialog { SelectedPath = System.IO.Directory.Exists(b) ? b : r })
                {
                    if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK) { BackupManager.RestoreFromBackup(d.SelectedPath); _optionsWindow.Close(); TrayManager.reloadAllFences(); }
                }
            }
            catch (Exception ex) { MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Restore failed: {ex.Message}", "Error"); }
        }

        private static void OpenBackupsFolder()
        {
            try
            {
                string p = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Backups");
                if (!System.IO.Directory.Exists(p)) System.IO.Directory.CreateDirectory(p);
                Process.Start(new ProcessStartInfo { FileName = p, UseShellExecute = true });
            }
            catch (Exception ex) { MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error: {ex.Message}", "Error"); }
        }

        private static void OpenLogFile()
        {
            try
            {
                string p = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                if (System.IO.File.Exists(p)) Process.Start(new ProcessStartInfo { FileName = p, UseShellExecute = true });
                else MessageBoxesManager.ShowOKOnlyMessageBoxForm("Log file not found.", "Information");
            }
            catch { }
        }

        private static void PerformFullFactoryReset()
        {
            if (MessageBoxesManager.ShowCustomYesNoMessageBox("WARNING: This will delete ALL fences, shortcuts, and settings!\n\nAre you sure you want to proceed?", "Factory Reset"))
            {
                try
                {
                    string ts = DateTime.Now.ToString("yyMMddHHmm");
                    BackupManager.CreateBackup($"{ts}_backup_reset", silent: true);
                    string ed = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                    foreach (string f in new[] { "Temp Shortcuts", "Shortcuts", "Last Fence Deleted", "CopiedItem" })
                    {
                        string p = System.IO.Path.Combine(ed, f);
                        if (System.IO.Directory.Exists(p)) { try { System.IO.Directory.Delete(p, true); System.IO.Directory.CreateDirectory(p); } catch { } }
                    }
                    string fj = System.IO.Path.Combine(ed, "fences.json"); if (System.IO.File.Exists(fj)) System.IO.File.Delete(fj);
                    string oj = System.IO.Path.Combine(ed, "options.json"); if (System.IO.File.Exists(oj)) System.IO.File.Delete(oj);
                    SettingsManager.LoadSettings();
                    FenceManager.ReloadFences();
                    _optionsWindow.Close();
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Error, $"Factory reset failed: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Reset failed: {ex.Message}", "Error");
                }
            }
        }
    }
}