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
        private static readonly Color ColorProfiles = Color.FromRgb(255, 20, 147); // Deep Pink
        private static readonly Color ColorHotkeys = Color.FromRgb(139, 69, 19); // SaddleBrown
        private static readonly Color ColorSmartDesktop = Color.FromRgb(41, 74, 122); // Semi Dark Blue

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
                // Add Pause Here:
                AutoOrganizeManager.Pause();

                _optionsWindow.ShowDialog();

                // Add Resume Here:
                AutoOrganizeManager.Resume();

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
            CreateProfilesTab();
            CreateHotkeysTab();
            CreateSmartDesktopTab();
            CreateLookDeeperTab();

            _tabControl.SelectedIndex = _lastSelectedTabIndex;
            CreateTabButton(tabPanel, "General", 0, _lastSelectedTabIndex == 0);
            CreateTabButton(tabPanel, "Style", 1, _lastSelectedTabIndex == 1);
            CreateTabButton(tabPanel, "Tools", 2, _lastSelectedTabIndex == 2);
            CreateTabButton(tabPanel, "Profiles", 3, _lastSelectedTabIndex == 3);
            CreateTabButton(tabPanel, "Hotkeys", 4, _lastSelectedTabIndex == 4);
            CreateTabButton(tabPanel, "Smart Desktop", 5, _lastSelectedTabIndex == 5);
            CreateTabButton(tabPanel, "Look Deeper", 6, _lastSelectedTabIndex == 6);

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
            Color activeColor = title switch { "Style" => ColorStyle, "Tools" => ColorTools, "Profiles" => ColorProfiles, "Hotkeys" => ColorHotkeys, "Smart Desktop" => ColorSmartDesktop, "Look Deeper" => ColorLookDeeper, _ => _userAccentColor };
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

            // NEW: Context Menu Option
            CreateCheckBox(c, "Show 'New Fence' in Desktop Context Menu", "EnableContextMenu", SettingsManager.EnableContextMenu);

            // --- NEW: Auto-Hide Options ---
            CreateSectionHeader(c, "Auto-Hide Fences", _userAccentColor);
            CreateCheckBox(c, "Auto hide fences", "AutoHideFences", SettingsManager.AutoHideFences);
            // FIX: Pass 300 as the explicit max for this specific slider
            CreateSliderControl(c, "Auto hide time (sec)", "AutoHideTimeSlider", SettingsManager.AutoHideTime, 300);

           
       //     CreateCheckBox(c, "Enable Profile Automation", "EnableProfileAutomation", SettingsManager.EnableProfileAutomation);
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

            // --- CHAMELEON TOGGLE ---
            var chamCb = CreateCheckBoxReturn(c, "Enable Chameleon Mode (Auto-match Wallpaper Color)", "EnableChameleon", SettingsManager.EnableChameleonMode);
            chamCb.ToolTip = "Fences will automatically change color to blend perfectly with your desktop background.";

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



            // --- Maintenance Section ---
            Color darkPink = Color.FromRgb(199, 21, 133); // MediumVioletRed
            CreateSectionHeader(c, "Maintenance", darkPink);

            Button btnBound = CreateStyledButton("Screen Bound Fences", darkPink);
            btnBound.Width = 255;
            btnBound.Height = 45;
            btnBound.Margin = new Thickness(0, 0, 0, 15);
            btnBound.HorizontalAlignment = HorizontalAlignment.Left;

            // Enable ONLY if Auto-Reposition is OFF (Manual Mode)
            btnBound.IsEnabled = !SettingsManager.AllowAutoReposition;

            if (btnBound.IsEnabled)
            {
                btnBound.Click += (s, e) =>
                {
                    // This calls the wrapper that handles the variable flipping
                    FenceManager.ForceRepositionAllFences();

                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("All fences have been checked and moved within valid screen bounds.", "Success");
                };
            }
            else
            {
                btnBound.Opacity = 0.90;
                btnBound.ToolTip = "Auto-reposition is active (Hidden Setting). Fences are already managed automatically.";
            }

            c.Children.Add(btnBound);



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





        private static void CreateProfilesTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();
            CreateSectionHeader(c, "Profile Management", ColorProfiles);

            // Button 1: Manage Profiles (Green)
            Button btnManageProfiles = CreateStyledButton("Manage Profiles", Color.FromRgb(34, 139, 34)); // Tools Green
            btnManageProfiles.Width = 255; btnManageProfiles.Height = 45; btnManageProfiles.Margin = new Thickness(15, 0, 0, 15);
            btnManageProfiles.HorizontalAlignment = HorizontalAlignment.Left;
            btnManageProfiles.Click += (s, e) => { new ProfileManagerForm().ShowDialog(); };

            // Button 2: Manage Automation (Blue)
            Button btnManageAutomation = CreateStyledButton("Manage Automation", Color.FromRgb(0, 123, 191)); // Tools Blue
            btnManageAutomation.Width = 255; btnManageAutomation.Height = 45; btnManageAutomation.Margin = new Thickness(15, 0, 0, 15);
            btnManageAutomation.HorizontalAlignment = HorizontalAlignment.Left;
            btnManageAutomation.Click += (s, e) => { new AutomationRulesForm().ShowDialog(); };

            // Separator and Toggle
            c.Children.Add(btnManageProfiles);
            c.Children.Add(btnManageAutomation);

            // Checkbox for Automation (Synchronized with Tray)
            CheckBox autoCb = new CheckBox
            {
                Name = "EnableProfileAutomation",
                Content = "Enable Profile Automation",
                IsChecked = SettingsManager.EnableProfileAutomation,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Margin = new Thickness(15, 10, 0, 8)
            };
            // Use Click event to ensure it only fires on user interaction, then SaveSettings immediately
            autoCb.Click += (s, e) => {
                bool isChecked = autoCb.IsChecked == true;
                SettingsManager.EnableProfileAutomation = isChecked;
                SettingsManager.SaveSettings(); // Force write to JSON immediately
                TrayManager.Instance?.UpdateAutomationMenuCheck(isChecked);
                if (isChecked) AutomationManager.Start();
            };
            c.Children.Add(autoCb);

            t.Content = new ScrollViewer { Content = c, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            _tabControl.Items.Add(t);
        }

        private static readonly Dictionary<string, int> AvailableKeys = new Dictionary<string, int>
        {
            {"A", 0x41}, {"B", 0x42}, {"C", 0x43}, {"D", 0x44}, {"E", 0x45}, {"F", 0x46}, {"G", 0x47}, {"H", 0x48}, {"I", 0x49}, {"J", 0x4A}, {"K", 0x4B}, {"L", 0x4C}, {"M", 0x4D}, {"N", 0x4E}, {"O", 0x4F}, {"P", 0x50}, {"Q", 0x51}, {"R", 0x52}, {"S", 0x53}, {"T", 0x54}, {"U", 0x55}, {"V", 0x56}, {"W", 0x57}, {"X", 0x58}, {"Y", 0x59}, {"Z", 0x5A},
            {"0", 0x30}, {"1", 0x31}, {"2", 0x32}, {"3", 0x33}, {"4", 0x34}, {"5", 0x35}, {"6", 0x36}, {"7", 0x37}, {"8", 0x38}, {"9", 0x39},
            {"F1", 0x70}, {"F2", 0x71}, {"F3", 0x72}, {"F4", 0x73}, {"F5", 0x74}, {"F6", 0x75}, {"F7", 0x76}, {"F8", 0x77}, {"F9", 0x78}, {"F10", 0x79}, {"F11", 0x7A}, {"F12", 0x7B},
            {"Comma (,)", 0xBC}, {"Period (.)", 0xBE}, {"Tilde (~)", 192}, {"Space", 32}, {"Tab", 9}, {"Enter", 13}, {"Escape", 27}
        };

        private static void CreateHotkeysTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();

            CreateSectionHeader(c, "Profile Switching", ColorHotkeys);
            CheckBox cbProf = CreateCheckBoxReturn(c, "Enable Profile Switching Hotkeys", "EnableProfileHotkeys", SettingsManager.EnableProfileHotkeys);
            Grid gProf1 = CreateHotkeyEditor(c, "Direct Profile [0-9]", "ProfSwitch", SettingsManager.ProfileSwitchModifier, 0, false);
            Grid gProf2 = CreateHotkeyEditor(c, "Previous Profile", "ProfPrev", SettingsManager.ProfilePrevModifier, SettingsManager.ProfilePrevKey, true);
            Grid gProf3 = CreateHotkeyEditor(c, "Next Profile", "ProfNext", SettingsManager.ProfileNextModifier, SettingsManager.ProfileNextKey, true);

            // Bind initial state and live toggling
            gProf1.IsEnabled = gProf2.IsEnabled = gProf3.IsEnabled = cbProf.IsChecked == true;
            cbProf.Click += (s, e) => gProf1.IsEnabled = gProf2.IsEnabled = gProf3.IsEnabled = cbProf.IsChecked == true;

            CreateSectionHeader(c, "Utilities", ColorHotkeys);

            CheckBox cbFocus = CreateCheckBoxReturn(c, "Enable Focus Fence Hotkey", "EnableFocusFenceHotkey", SettingsManager.EnableFocusFenceHotkey);
            Grid gFocus = CreateHotkeyEditor(c, "Focus Fence", "FocusFence", SettingsManager.FocusFenceModifier, SettingsManager.FocusFenceKey, true);
            gFocus.IsEnabled = cbFocus.IsChecked == true;
            cbFocus.Click += (s, e) => gFocus.IsEnabled = cbFocus.IsChecked == true;

            CheckBox cbSpot = CreateCheckBoxReturn(c, "Enable Spot Search Hotkey", "EnableSpotSearchHotkey", SettingsManager.EnableSpotSearchHotkey);
            Grid gSpot = CreateHotkeyEditor(c, "Spot Search", "SpotSearch", SettingsManager.SpotSearchModifier, SettingsManager.SpotSearchKey, true);
            gSpot.IsEnabled = cbSpot.IsChecked == true;
            cbSpot.Click += (s, e) => gSpot.IsEnabled = cbSpot.IsChecked == true;

            TextBlock infoText = new TextBlock
            {
                Text = "Note: Changes to Global Hotkeys require an application restart to take effect.",
                FontStyle = FontStyles.Italic,
                Foreground = Brushes.Gray,
                Margin = new Thickness(15, 20, 0, 0)
            };
            c.Children.Add(infoText);

            t.Content = new ScrollViewer { Content = c, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            _tabControl.Items.Add(t);
        }

        private static Grid CreateHotkeyEditor(StackPanel p, string label, string namePrefix, string currentMod, int currentKey, bool hasKeySelector)
        {
            Grid g = new Grid { Margin = new Thickness(15, 5, 0, 15) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock lbl = new TextBlock { Text = label, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(lbl, 0); g.Children.Add(lbl);

            StackPanel spMods = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };

            string curModLower = (currentMod ?? "").ToLower();

            CheckBox chkCtrl = new CheckBox { Name = namePrefix + "Ctrl", Content = "Ctrl", IsChecked = curModLower.Contains("ctrl") || curModLower.Contains("control"), Margin = new Thickness(0, 0, 10, 0), VerticalAlignment = VerticalAlignment.Center };
            CheckBox chkAlt = new CheckBox { Name = namePrefix + "Alt", Content = "Alt", IsChecked = curModLower.Contains("alt"), Margin = new Thickness(0, 0, 10, 0), VerticalAlignment = VerticalAlignment.Center };
            CheckBox chkShift = new CheckBox { Name = namePrefix + "Shift", Content = "Shift", IsChecked = curModLower.Contains("shift"), Margin = new Thickness(0, 0, 10, 0), VerticalAlignment = VerticalAlignment.Center };
            CheckBox chkWin = new CheckBox { Name = namePrefix + "Win", Content = "Win", IsChecked = curModLower.Contains("win"), Margin = new Thickness(0, 0, 15, 0), VerticalAlignment = VerticalAlignment.Center };

            spMods.Children.Add(chkCtrl);
            spMods.Children.Add(chkAlt);
            spMods.Children.Add(chkShift);
            spMods.Children.Add(chkWin);

            if (hasKeySelector)
            {
                spMods.Children.Add(new TextBlock { Text = "+", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 0) });
                ComboBox cmb = new ComboBox { Name = namePrefix + "Key", Width = 100, VerticalAlignment = VerticalAlignment.Center };
                foreach (var kvp in AvailableKeys)
                {
                    ComboBoxItem item = new ComboBoxItem { Content = kvp.Key, Tag = kvp.Value };
                    cmb.Items.Add(item);
                    if (kvp.Value == currentKey) cmb.SelectedItem = item;
                }
                if (cmb.SelectedIndex == -1 && cmb.Items.Count > 0) cmb.SelectedIndex = 0;
                spMods.Children.Add(cmb);
            }

            Grid.SetColumn(spMods, 1);
            g.Children.Add(spMods);
            p.Children.Add(g);
            return g;
        }


        private static void CreateSmartDesktopTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();

            CreateSectionHeader(c, "Smart Desktop (Auto-Organize)", ColorSmartDesktop);

            CheckBox cbMain = CreateCheckBoxReturn(c, "Enable Auto-Organize", "EnableAutoOrganize", SettingsManager.EnableAutoOrganize);

            CheckBox cbNotif = CreateCheckBoxReturn(c, "Show execution toast notifications", "EnableAutoOrganizeNotifications", SettingsManager.EnableAutoOrganizeNotifications);
            cbNotif.Margin = new Thickness(35, 0, 0, 8); // Indent it!
            cbNotif.IsEnabled = cbMain.IsChecked == true;

            // NEW: Live Rule Statistics (Horizontal Layout)
            StackPanel statsPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(15, 15, 0, 15) };
            TextBlock txtTotalRules = new TextBlock { Text = $"Total number of rules: {AutoOrganizeManager.Rules.Count}", FontFamily = new FontFamily("Segoe UI"), FontSize = 13, FontWeight = FontWeights.Medium };
            TextBlock txtSeparator = new TextBlock { Text = "   -   ", FontFamily = new FontFamily("Segoe UI"), FontSize = 13, FontWeight = FontWeights.Medium, Foreground = Brushes.Gray };
            TextBlock txtEnabledRules = new TextBlock { Text = $"Enabled: {AutoOrganizeManager.Rules.Count(r => r.IsEnabled)}", FontFamily = new FontFamily("Segoe UI"), FontSize = 13, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(Color.FromRgb(34, 139, 34)) };
            statsPanel.Children.Add(txtTotalRules);
            statsPanel.Children.Add(txtSeparator);
            statsPanel.Children.Add(txtEnabledRules);
            c.Children.Add(statsPanel);

            // Navy Blue - Manage Rules Button
            Button btnManageRules = CreateStyledButton("Smart Desktop Rules...", Color.FromRgb(0, 0, 128));
            btnManageRules.Width = 255;
            btnManageRules.Height = 45;
            btnManageRules.Margin = new Thickness(15, 0, 0, 15);
            btnManageRules.HorizontalAlignment = HorizontalAlignment.Left;
            btnManageRules.Click += (s, e) =>
            {
                new AutoOrganizeForm().ShowDialog();
                // Refresh statistics when the editor closes
                txtTotalRules.Text = $"Total number of rules: {AutoOrganizeManager.Rules.Count}";
                txtEnabledRules.Text = $"Enabled: {AutoOrganizeManager.Rules.Count(r => r.IsEnabled)}";
            };
            c.Children.Add(btnManageRules);

            // Dark Red - Organize Desktop Now Button
            Button btnOrganizeNow = CreateStyledButton("Organize Now (Run)", Color.FromRgb(139, 0, 0));
            btnOrganizeNow.Width = 255;
            btnOrganizeNow.Height = 45;
            btnOrganizeNow.Margin = new Thickness(15, 0, 0, 15);
            btnOrganizeNow.HorizontalAlignment = HorizontalAlignment.Left;
            btnOrganizeNow.Click += (s, e) =>
            {
                if (MessageBoxesManager.ShowCustomYesNoMessageBox("This will move existing files on your desktop to your target folders based on your rules.\n\nProceed?", "Sweep Desktop"))
                {
                    AutoOrganizeManager.ProcessDesktopNow();
                }
            };
            c.Children.Add(btnOrganizeNow);

            TextBlock infoText = new TextBlock
            {
                Text = "Note: Auto-Organize continuously monitors your Desktop for new files. When a file matches an enabled rule's conditions, it is automatically and physically moved to your target Portal Fence or Folder. Use this to keep your Desktop permanently clean and automatically route downloads to their proper locations.",
                FontStyle = FontStyles.Italic,
                Foreground = Brushes.Gray,
                Margin = new Thickness(15, 20, 0, 0),
                TextWrapping = TextWrapping.Wrap
            };
            c.Children.Add(infoText);

            t.Content = new ScrollViewer { Content = c, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
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

        // FIX: Added 'max' parameter (defaulting to 100) to fix the Tint sliders while supporting AutoHideTime
        private static void CreateSliderControl(StackPanel p, string l, string n, int v, int max = 100)
        {
            Grid g = new Grid { Margin = new Thickness(0, 5, 0, 5) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Responsive Label Width
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(50) });
            // Add a slight margin to the slider itself so it spaces out nicely from the dynamic label
            g.Margin = new Thickness(15, 5, 0, 5);
            TextBlock lbl = new TextBlock { Text = l, FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 10, 0) };
            Slider sl = new Slider { Name = n, Minimum = 1, Maximum = max, Value = v, TickFrequency = 1, IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center };
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

                        // NEW: Context Menu Registry Update
                        if (cb.Name == "EnableContextMenu")
                        {
                            bool newState = cb.IsChecked == true;
                            if (SettingsManager.EnableContextMenu != newState)
                            {
                                SettingsManager.EnableContextMenu = newState;
                                RegistryHelper.ToggleContextMenu(newState);
                            }
                        }

                        // NEW: Auto-Hide Options
                        if (cb.Name == "AutoHideFences")
                        {
                            SettingsManager.AutoHideFences = cb.IsChecked == true;
                            FenceManager.ResetAutoHideTimer();
                        }
                    }
                    else if (child is Grid g)
                    {
                        var autoHideTime = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "AutoHideTimeSlider");
                        if (autoHideTime != null)
                        {
                            SettingsManager.AutoHideTime = (int)autoHideTime.Value;
                            FenceManager.ResetAutoHideTimer();
                        }
                    }

                    // REMOVED: EnableProfileAutomation logic is now handled exclusively in the Profiles tab.
                    //if (cb.Name == "EnableProfileAutomation")
                    //{
                    //    SettingsManager.EnableProfileAutomation = cb.IsChecked == true;
                    //    if (SettingsManager.EnableProfileAutomation) AutomationManager.Start();
                    //}

                }


                // 2. Style
                var styleContent = (StackPanel)((ScrollViewer)((TabItem)_tabControl.Items[1]).Content).Content;
                foreach (var child in styleContent.Children)
                {
                    if (child is CheckBox cb)
                    {
                        if (cb.Name == "EnableChameleon") SettingsManager.EnableChameleonMode = cb.IsChecked == true;
                        if (cb.Name == "EnablePortalWatermark") { newPortalWatermarkState = cb.IsChecked == true; SettingsManager.ShowBackgroundImageOnPortalFences = newPortalWatermarkState; }
                        if (cb.Name == "DisableFenceScrollbars") SettingsManager.DisableFenceScrollbars = cb.IsChecked == true;
                        if (cb.Name == "EnableSounds") SettingsManager.EnableSounds = cb.IsChecked == true;
                    }
                    else if (child is Grid g)
                    {
                        var tint = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "TintSlider"); if (tint != null) SettingsManager.TintValue = (int)tint.Value; var mtint = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "MenuTintSlider"); if (mtint != null) SettingsManager.MenuTintValue = (int)mtint.Value;
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

                // 4. Hotkeys (NEW)
                var hotkeysContent = (StackPanel)((ScrollViewer)((TabItem)_tabControl.Items[4]).Content).Content;
                bool hotkeysChanged = false;
                foreach (var child in hotkeysContent.Children)
                {
                    if (child is CheckBox hotkeyCb)
                    {
                        if (hotkeyCb.Name == "EnableProfileHotkeys" && SettingsManager.EnableProfileHotkeys != (hotkeyCb.IsChecked == true)) { SettingsManager.EnableProfileHotkeys = hotkeyCb.IsChecked == true; hotkeysChanged = true; }
                        if (hotkeyCb.Name == "EnableFocusFenceHotkey" && SettingsManager.EnableFocusFenceHotkey != (hotkeyCb.IsChecked == true)) { SettingsManager.EnableFocusFenceHotkey = hotkeyCb.IsChecked == true; hotkeysChanged = true; }
                        if (hotkeyCb.Name == "EnableSpotSearchHotkey" && SettingsManager.EnableSpotSearchHotkey != (hotkeyCb.IsChecked == true)) { SettingsManager.EnableSpotSearchHotkey = hotkeyCb.IsChecked == true; hotkeysChanged = true; }
                    }

                    if (child is Grid g && g.Children.Count > 1 && g.Children[1] is StackPanel spMods)
                    {
                        string prefix = "";
                        foreach (var elem in spMods.Children)
                        {
                            if (elem is CheckBox cb && cb.Name.EndsWith("Ctrl"))
                            {
                                prefix = cb.Name.Substring(0, cb.Name.Length - 4);
                                break;
                            }
                        }
                        if (!string.IsNullOrEmpty(prefix))
                        {
                            List<string> mods = new List<string>();
                            int key = 0;
                            foreach (var elem in spMods.Children)
                            {
                                if (elem is CheckBox cb && cb.IsChecked == true)
                                {
                                    if (cb.Name.EndsWith("Ctrl")) mods.Add("Control");
                                    else if (cb.Name.EndsWith("Alt")) mods.Add("Alt");
                                    else if (cb.Name.EndsWith("Shift")) mods.Add("Shift");
                                    else if (cb.Name.EndsWith("Win")) mods.Add("Win");
                                }
                                if (elem is ComboBox cmb && cmb.SelectedItem is ComboBoxItem item && item.Tag is int val)
                                {
                                    key = val;
                                }
                            }
                            string modString = string.Join(", ", mods);

                            if (prefix == "ProfSwitch") { if (SettingsManager.ProfileSwitchModifier != modString) { SettingsManager.ProfileSwitchModifier = modString; hotkeysChanged = true; } }
                            if (prefix == "ProfPrev") { if (SettingsManager.ProfilePrevModifier != modString || SettingsManager.ProfilePrevKey != key) { SettingsManager.ProfilePrevModifier = modString; SettingsManager.ProfilePrevKey = key; hotkeysChanged = true; } }
                            if (prefix == "ProfNext") { if (SettingsManager.ProfileNextModifier != modString || SettingsManager.ProfileNextKey != key) { SettingsManager.ProfileNextModifier = modString; SettingsManager.ProfileNextKey = key; hotkeysChanged = true; } }
                            if (prefix == "FocusFence") { if (SettingsManager.FocusFenceModifier != modString || SettingsManager.FocusFenceKey != key) { SettingsManager.FocusFenceModifier = modString; SettingsManager.FocusFenceKey = key; hotkeysChanged = true; } }
                            if (prefix == "SpotSearch") { if (SettingsManager.SpotSearchModifier != modString || SettingsManager.SpotSearchKey != key) { SettingsManager.SpotSearchModifier = modString; SettingsManager.SpotSearchKey = key; hotkeysChanged = true; } }
                        }
                    }
                }

                if (hotkeysChanged)
                {
                    // Propagate the new hotkeys across all existing profiles
                    SettingsManager.BroadcastHotkeysToAllProfiles();

                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Global Hotkey changes have been saved and applied to all profiles.\n\nPlease restart Desktop Fences to activate the new shortcuts.", "Restart Required");
                }

                // 5. Smart Desktop (Auto-Organize)
                var smartDesktopContent = (StackPanel)((ScrollViewer)((TabItem)_tabControl.Items[5]).Content).Content;
                foreach (var child in smartDesktopContent.Children)
                {
                    if (child is CheckBox cb && cb.Name == "EnableAutoOrganize")
                    {
                        bool wasEnabled = SettingsManager.EnableAutoOrganize;
                        SettingsManager.EnableAutoOrganize = cb.IsChecked == true;

                        // Sync with the Tray icon context menu!
                        TrayManager.Instance?.UpdateAutoOrganizeMenuCheck(SettingsManager.EnableAutoOrganize);

                        // Live toggle the background engine
                        if (!wasEnabled && SettingsManager.EnableAutoOrganize) AutoOrganizeManager.Start();
                        else if (wasEnabled && !SettingsManager.EnableAutoOrganize) AutoOrganizeManager.Stop();
                    }
                    if (child is CheckBox cbn && cbn.Name == "EnableAutoOrganizeNotifications")
                    {
                        SettingsManager.EnableAutoOrganizeNotifications = cbn.IsChecked == true;
                    }
                }

                // 6. Look Deeper (Logs) - Index shifted to 6
                var logContent = (StackPanel)((ScrollViewer)((TabItem)_tabControl.Items[6]).Content).Content;
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
                using (var d = new System.Windows.Forms.FolderBrowserDialog())
                {
                    // FIX: Use the Profile-Aware path helper
                    d.SelectedPath = BackupManager.GetBackupsFolderPath();
                    d.Description = "Select a backup folder to restore from";

                    if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        BackupManager.RestoreFromBackup(d.SelectedPath);
                        _optionsWindow.Close();
                        TrayManager.reloadAllFences();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Restore failed: {ex.Message}", "Error");
            }
        }
        private static void OpenBackupsFolder()
        {
            // Use the centralized BackupManager helper
            BackupManager.OpenBackupsFolder();
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
            if (MessageBoxesManager.ShowCustomYesNoMessageBox("WARNING: This will delete ALL fences, shortcuts, and settings for the CURRENT PROFILE!\n\nAre you sure you want to proceed?", "Factory Reset"))
            {
                try
                {
                    // 1. Create a safety backup before wiping
                    string ts = DateTime.Now.ToString("yyMMddHHmm");
                    BackupManager.CreateBackup($"{ts}_backup_reset", silent: true);

                    // 2. Wipe Profile-Specific Folders
                    // FIX: Use ProfileManager.GetProfileFilePath to target the active profile
                    foreach (string f in new[] { "Temp Shortcuts", "Shortcuts", "Last Fence Deleted", "CopiedItem" })
                    {
                        string p = ProfileManager.GetProfileFilePath(f);
                        if (System.IO.Directory.Exists(p))
                        {
                            try
                            {
                                System.IO.Directory.Delete(p, true);
                                System.IO.Directory.CreateDirectory(p);
                            }
                            catch { }
                        }
                    }

                    // 3. Wipe Profile-Specific Config Files
                    string fj = ProfileManager.GetProfileFilePath("fences.json");
                    if (System.IO.File.Exists(fj)) System.IO.File.Delete(fj);

                    string oj = ProfileManager.GetProfileFilePath("options.json");
                    if (System.IO.File.Exists(oj)) System.IO.File.Delete(oj);

                    // 4. Reload to apply the "Empty" state
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