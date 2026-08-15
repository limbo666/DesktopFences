using Desktop_Frames.Localization;
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

namespace Desktop_Frames
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
                    Title = Strings.OptionsTitle,
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
                    Text = Strings.OptionsHeading,
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
                // Pause Here:
                AutoOrganizeManager.Pause();

                _optionsWindow.ShowDialog();

                // Resume Here:
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
            CreateTabButton(tabPanel, Strings.TabGeneral, 0, _lastSelectedTabIndex == 0);
            CreateTabButton(tabPanel, Strings.TabStyleFx, 1, _lastSelectedTabIndex == 1);
            CreateTabButton(tabPanel, Strings.TabTools, 2, _lastSelectedTabIndex == 2);
            CreateTabButton(tabPanel, Strings.TabProfiles, 3, _lastSelectedTabIndex == 3);
            CreateTabButton(tabPanel, Strings.TabHotkeys, 4, _lastSelectedTabIndex == 4);
            CreateTabButton(tabPanel, Strings.TabSmartDesktop, 5, _lastSelectedTabIndex == 5);
            CreateTabButton(tabPanel, Strings.TabLookDeeper, 6, _lastSelectedTabIndex == 6);

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
            SetTabButtonColors(tabButton, tabIndex, isSelected);
            tabButton.Click += (s, e) => SelectTab(tabIndex, tabButton);
            tabButton.MouseEnter += (s, e) => { if (_tabControl.SelectedIndex != tabIndex) SetTabButtonColors(tabButton, tabIndex, false, true); };
            tabButton.MouseLeave += (s, e) => { if (_tabControl.SelectedIndex != tabIndex) SetTabButtonColors(tabButton, tabIndex, false, false); };
            parent.Children.Add(tabButton);
        }

        private static void SetTabButtonColors(Button button, int tabIndex, bool isSelected, bool isHover = false)
        {
            // Keyed on the tab index, not on its label. The labels are translated
            // now, and a switch needs compile-time constants anyway.
            Color activeColor = tabIndex switch { 1 => ColorStyle, 2 => ColorTools, 3 => ColorProfiles, 4 => ColorHotkeys, 5 => ColorSmartDesktop, 6 => ColorLookDeeper, _ => _userAccentColor };
            if (isSelected) { button.Background = new SolidColorBrush(activeColor); button.Foreground = Brushes.White; }
            else if (isHover) { button.Background = new SolidColorBrush(Color.FromRgb((byte)(activeColor.R + 40), (byte)(activeColor.G + 40), (byte)(activeColor.B + 40))); button.Foreground = Brushes.White; }
            else { button.Background = new SolidColorBrush(Color.FromRgb(200, 200, 200)); button.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 60)); }
        }

        private static void SelectTab(int tabIndex, Button selectedButton)
        {
            _lastSelectedTabIndex = tabIndex;
            _tabControl.SelectedIndex = tabIndex;
            StackPanel tabPanel = (StackPanel)selectedButton.Parent;
            for (int i = 0; i < tabPanel.Children.Count; i++) if (tabPanel.Children[i] is Button btn) SetTabButtonColors(btn, i, i == tabIndex);
        }

        // --- Tabs ---
        private static void CreateGeneralTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();
            CreateSectionHeader(c, Strings.SecStartup, _userAccentColor);
            CreateCheckBox(c, Strings.OptStartWithWindows, "StartWithWindows", TrayManager.IsStartWithWindows);
            CreateSectionHeader(c, Strings.SecSelections, _userAccentColor);
            CreateCheckBox(c, Strings.OptSingleClick, "SingleClickToLaunch", SettingsManager.SingleClickToLaunch);
            CreateCheckBox(c, Strings.OptSnapNearFrames, "EnableSnapNearFrames", SettingsManager.IsSnapEnabled);
            CreateCheckBox(c, Strings.OptDimensionSnap, "EnableDimensionSnap", SettingsManager.EnableDimensionSnap);
            CreateCheckBox(c, Strings.OptTrayIcon, "EnableTrayIcon", SettingsManager.ShowInTray);
            CreateCheckBox(c, Strings.OptRecycleBin, "UseRecycleBin", SettingsManager.UseRecycleBin);

            // NEW: Context Menu Option
            CreateCheckBox(c, Strings.OptNewFrameContextMenu, "EnableContextMenu", SettingsManager.EnableContextMenu);

            // --- NEW: Default Portal View Dropdown ---
            Grid portalViewGrid = new Grid { Margin = new Thickness(15, 8, 0, 8) };
            portalViewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            portalViewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });

            TextBlock lblPortalView = new TextBlock { Text = Strings.LblDefaultPortalView, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(lblPortalView, 0);

            ComboBox cbPortalView = new ComboBox { Name = "DefaultPortalViewComboBox", Height = 25, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            // Content is what the user reads, Tag is what gets written to the
            // settings file. Never save the label: it changes with the language.
            cbPortalView.Items.Add(new ComboBoxItem { Content = Strings.ViewIcons, Tag = "Icons" });
            cbPortalView.Items.Add(new ComboBoxItem { Content = Strings.ViewDetails, Tag = "Details" });
            cbPortalView.SelectedIndex = string.Equals(SettingsManager.DefaultPortalView, "Details", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
            Grid.SetColumn(cbPortalView, 1);

            portalViewGrid.Children.Add(lblPortalView);
            portalViewGrid.Children.Add(cbPortalView);
            c.Children.Add(portalViewGrid);
            // -----------------------------------------

            // Moved from Style Tab (Choices)
            CreateCheckBox(c, Strings.OptPortalWatermark, "EnablePortalWatermark", SettingsManager.ShowBackgroundImageOnPortalFrames);
            var n = CreateCheckBoxReturn(c, Strings.OptNoteWatermark, "EnableNoteWatermark", false);
            n.IsEnabled = false; n.Foreground = Brushes.Gray;
            CreateCheckBox(c, Strings.OptDisableScrollbars, "DisableFrameScrollbars", SettingsManager.DisableFrameScrollbars);


            // --- NEW: Notification Sound Dropdown ---
            CheckBox cbSounds = CreateCheckBoxReturn(c, Strings.OptEnableSounds, "EnableSounds", SettingsManager.EnableSounds);

            Grid soundGrid = new Grid { Margin = new Thickness(35, 0, 0, 8) }; // Indented to show parent/child relationship
            soundGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            soundGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });

            TextBlock lblSound = new TextBlock { Text = Strings.LblNotificationSound, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(lblSound, 0);

            ComboBox cbSoundType = new ComboBox { Name = "NotificationSoundComboBox", Height = 25, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            cbSoundType.Items.Add(Strings.SndDefault);
            cbSoundType.Items.Add(Strings.SndDoubleDing);
            cbSoundType.Items.Add(Strings.SndSmoothTickle);
            cbSoundType.Items.Add(Strings.SndMessageDing);
            cbSoundType.Items.Add(Strings.SndGentleDing);
            cbSoundType.Items.Add(Strings.SndSoftDing);

            // Map the current Enum back to the UI index
            cbSoundType.SelectedIndex = SettingsManager.NotificationSound switch
            {
                NotificationSound.DoubleDing => 1,
                NotificationSound.SmoothTickle => 2,
                NotificationSound.MessageDing => 3,
                NotificationSound.GentleDing => 4,
                NotificationSound.SoftDing => 5,
                _ => 0
            };
            Grid.SetColumn(cbSoundType, 1);
            soundGrid.Children.Add(lblSound);
            soundGrid.Children.Add(cbSoundType);
            c.Children.Add(soundGrid);

            // Live-toggle the combobox based on the checkbox state
            soundGrid.IsEnabled = cbSounds.IsChecked == true;
            cbSounds.Click += (s, e) => soundGrid.IsEnabled = cbSounds.IsChecked == true;
            // ----------------------------------------

            //     CreateCheckBox(c, Strings.LblEnableProfileAutomation, "EnableProfileAutomation", SettingsManager.EnableProfileAutomation);
            t.Content = new ScrollViewer { Content = c, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            _tabControl.Items.Add(t);
        }

        private static void CreateStyleTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();

            CreateSectionHeader(c, Strings.SecAppearance, ColorStyle);

            // --- CHAMELEON TOGGLE ---
            var chamCb = CreateCheckBoxReturn(c, Strings.OptChameleon, "EnableChameleon", SettingsManager.EnableChameleonMode);
            chamCb.ToolTip = Strings.TooltipChameleon;

            CreateSliderControl(c, Strings.SldFrameTint, "TintSlider", SettingsManager.TintValue);
            CreateSliderControl(c, Strings.SldMenuTint, "MenuTintSlider", SettingsManager.MenuTintValue);

            // Pass the chamCb reference so we can wire up the toggle event
            CreateColorAndEffectComboBoxes(c, chamCb);

            CreateCheckBox(c, Strings.OptFrameTint, "ApplyTintToIcons", SettingsManager.ApplyTintToIcons);

            // --- Moved from General Tab (Auto-Hide Options) ---
            CreateSectionHeader(c, Strings.SecAutoHideFrames, ColorStyle);
            CreateCheckBox(c, Strings.OptAutoHideFrames, "AutoHideFrames", SettingsManager.AutoHideFrames);
            CreateSliderControl(c, Strings.SldAutoHideTime, "AutoHideTimeSlider", SettingsManager.AutoHideTime, 300);

            // --- Moved from General Tab (Idle Fade-Out Options) ---
            CreateSectionHeader(c, Strings.SecIdleFadeOut, ColorStyle);
            CreateCheckBox(c, Strings.OptIdleFadeOut, "FramesFadeOutFx", SettingsManager.FramesFadeOutFx);
            CreateSliderControl(c, Strings.SldIdleTime, "FadeOutTimeSlider", SettingsManager.FadeOutTime, 300);
            CreateSliderControl(c, Strings.SldFadeTargetOpacity, "FadeOutAlphaSlider", (int)(SettingsManager.FadeOutFxTargetAlpha * 100), 100);

            // --- Moved from General Tab (Idle Auto-Roll Options) ---
            CreateSectionHeader(c, Strings.SecIdleAutoRoll, ColorStyle);

            // --- NEW: Contextual Help Text ---
            c.Children.Add(new TextBlock
            {
                Text = Strings.NoteAutoRoll,
                FontStyle = FontStyles.Italic,
                Foreground = Brushes.Gray,
                FontSize = 12,
                Margin = new Thickness(15, -10, 0, 10), // Negative top margin pulls it closer to the header
                TextWrapping = TextWrapping.Wrap
            });

            CreateSliderControl(c, Strings.SldIdleTime, "AutoRollTimeSlider", SettingsManager.AutoRollTime, 300);


            // --- NEW: Desktop Icon Visibility ---
            CreateSectionHeader(c, Strings.SecDesktopIconVisibility, ColorStyle);
            CreateCheckBox(c, Strings.OptHideIconsRunning, "HideDesktopElementsOnStart", SettingsManager.HideDesktopElementsOnStart);
            CreateCheckBox(c, Strings.OptHideIconsWhenHidden, "HideDesktopElementsOnAllFramesHide", SettingsManager.HideDesktopElementsOnAllFramesHide);



            // --- Icons Section ---
            CreateSectionHeader(c, Strings.SecIcons, ColorStyle);
            Grid iconGrid = new Grid { Margin = new Thickness(15, 5, 0, 15) };
            iconGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            iconGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            StackPanel menuIconPanel = new StackPanel();
            menuIconPanel.Children.Add(new TextBlock { Text = Strings.LblMenuIcon, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 0) });
            CreateIconRadioButtonGroup(menuIconPanel, "MenuIconGroup", new Dictionary<string, int> { { "♥", 0 }, { "☰", 1 }, { "≣", 2 }, { "𓃑", 3 } }, SettingsManager.MenuIcon);

            StackPanel lockIconPanel = new StackPanel();
            lockIconPanel.Children.Add(new TextBlock { Text = Strings.LblLockIcon, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 0) });
            CreateIconRadioButtonGroup(lockIconPanel, "LockIconGroup", new Dictionary<string, int> { { "🛡️", 0 }, { "🔑", 1 }, { "🔐", 2 }, { "🔒", 3 } }, SettingsManager.LockIcon);

            Grid.SetColumn(menuIconPanel, 0);
            Grid.SetColumn(lockIconPanel, 1);
            iconGrid.Children.Add(menuIconPanel);
            iconGrid.Children.Add(lockIconPanel);
            c.Children.Add(iconGrid);

            t.Content = new ScrollViewer { Content = c, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            _tabControl.Items.Add(t);
        }

        private static void CreateToolsTab()
        {
            TabItem t = new TabItem();
            StackPanel c = new StackPanel();
            CreateSectionHeader(c, Strings.TabTools, ColorTools);

            Grid g = new Grid { Margin = new Thickness(0, 10, 0, 0) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(15) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) });
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(15) });
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) });

            Button b1 = CreateStyledButton(Strings.BtnBackup, ColorTools); b1.Click += (s, e) => BackupManager.BackupData();
            Button b2 = CreateStyledButton(Strings.BtnRestore, Color.FromRgb(255, 152, 0)); b2.Click += (s, e) => RestoreBackup();
            Button b3 = CreateStyledButton(Strings.BtnOpenBackupsFolder, Color.FromRgb(0, 123, 191)); b3.Click += (s, e) => OpenBackupsFolder();
            Grid.SetRow(b1, 0); Grid.SetColumn(b1, 0);
            Grid.SetRow(b2, 0); Grid.SetColumn(b2, 2);
            Grid.SetRow(b3, 2); Grid.SetColumn(b3, 0); Grid.SetColumnSpan(b3, 3);
            g.Children.Add(b1); g.Children.Add(b2); g.Children.Add(b3);
            c.Children.Add(g);

            CreateCheckBox(c, Strings.OptAutomaticBackup, "EnableAutoBackup", SettingsManager.EnableAutoBackup);



            // --- Maintenance Section ---
            Color darkPink = Color.FromRgb(199, 21, 133); // MediumVioletRed
            CreateSectionHeader(c, Strings.SecMaintenance, darkPink);

            Button btnBound = CreateStyledButton(Strings.BtnScreenBoundFrames, darkPink);
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
                    Framemanager.ForceRepositionallFrames();

                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("All frames have been checked and moved within valid screen bounds.", "Success");
                };
            }
            else
            {
                btnBound.Opacity = 0.90;
                btnBound.ToolTip = Strings.TooltipAutoReposition;
            }

            c.Children.Add(btnBound);



            CreateSectionHeader(c, Strings.BtnReset, Colors.Red);
            Button r1 = CreateStyledButton(Strings.BtnResetStyles, Color.FromRgb(108, 117, 125));
            r1.Width = 255; r1.Height = 45; r1.Margin = new Thickness(0, 0, 0, 15);
            r1.Click += (s, e) => { if (MessageBoxesManager.ShowCustomYesNoMessageBox("Reset all visual customizations?", Strings.BtnReset)) { Framemanager.ResetAllCustomizations(); _optionsWindow.Close(); } };

            Button r2 = CreateStyledButton(Strings.BtnClearAllData, Color.FromRgb(220, 53, 69));
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
            CreateSectionHeader(c, Strings.SecProfileManagement, ColorProfiles);

            // Button 1: Manage Profiles (Green)
            Button btnManageProfiles = CreateStyledButton(Strings.BtnManageProfiles, Color.FromRgb(34, 139, 34)); // Tools Green
            btnManageProfiles.Width = 255; btnManageProfiles.Height = 45; btnManageProfiles.Margin = new Thickness(15, 0, 0, 15);
            btnManageProfiles.HorizontalAlignment = HorizontalAlignment.Left;
            btnManageProfiles.Click += (s, e) => { new ProfileManagerForm().ShowDialog(); };

            // Button 2: Manage Automation (Blue)
            Button btnManageAutomation = CreateStyledButton(Strings.BtnManageAutomation, Color.FromRgb(0, 123, 191)); // Tools Blue
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
                Content = Strings.LblEnableProfileAutomation,
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

            CreateSectionHeader(c, Strings.SecProfileSwitching, ColorHotkeys);
            CheckBox cbProf = CreateCheckBoxReturn(c, Strings.OptProfileHotkeys, "EnableProfileHotkeys", SettingsManager.EnableProfileHotkeys);
            Grid gProf1 = CreateHotkeyEditor(c, Strings.HkDirectProfile, "ProfSwitch", SettingsManager.ProfileSwitchModifier, 0, false);
            Grid gProf2 = CreateHotkeyEditor(c, Strings.HkPreviousProfile, "ProfPrev", SettingsManager.ProfilePrevModifier, SettingsManager.ProfilePrevKey, true);
            Grid gProf3 = CreateHotkeyEditor(c, Strings.HkNextProfile, "ProfNext", SettingsManager.ProfileNextModifier, SettingsManager.ProfileNextKey, true);

            // Bind initial state and live toggling
            gProf1.IsEnabled = gProf2.IsEnabled = gProf3.IsEnabled = cbProf.IsChecked == true;
            cbProf.Click += (s, e) => gProf1.IsEnabled = gProf2.IsEnabled = gProf3.IsEnabled = cbProf.IsChecked == true;

            CreateSectionHeader(c, Strings.SecUtilities, ColorHotkeys);

            CheckBox cbFocus = CreateCheckBoxReturn(c, Strings.OptFocusFrameHotkey, "EnableFocusFrameHotkey", SettingsManager.EnableFocusFrameHotkey);
            Grid gFocus = CreateHotkeyEditor(c, Strings.HkFocusFrame, "FocusFrame", SettingsManager.FocusFrameModifier, SettingsManager.FocusFrameKey, true);
            gFocus.IsEnabled = cbFocus.IsChecked == true;
            cbFocus.Click += (s, e) => gFocus.IsEnabled = cbFocus.IsChecked == true;

            CheckBox cbSpot = CreateCheckBoxReturn(c, Strings.OptSpotSearchHotkey, "EnableSpotSearchHotkey", SettingsManager.EnableSpotSearchHotkey);
            Grid gSpot = CreateHotkeyEditor(c, Strings.HkSpotSearch, "SpotSearch", SettingsManager.SpotSearchModifier, SettingsManager.SpotSearchKey, true);
            gSpot.IsEnabled = cbSpot.IsChecked == true;
            cbSpot.Click += (s, e) => gSpot.IsEnabled = cbSpot.IsChecked == true;

            TextBlock infoText = new TextBlock
            {
                Text = Strings.NoteHotkeysRestart,
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
                    ComboBoxItem item = new ComboBoxItem { Content = Strings.KeyLabel(kvp.Key), Tag = kvp.Value };
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

            CreateSectionHeader(c, Strings.SecSmartDesktopAuto, ColorSmartDesktop);

            CheckBox cbMain = CreateCheckBoxReturn(c, Strings.OptAutoOrganize, "EnableAutoOrganize", SettingsManager.EnableAutoOrganize);

            CheckBox cbNotif = CreateCheckBoxReturn(c, Strings.OptExecutionToasts, "EnableAutoOrganizeNotifications", SettingsManager.EnableAutoOrganizeNotifications);
            cbNotif.Margin = new Thickness(35, 0, 0, 8); // Indent it!
            cbNotif.IsEnabled = cbMain.IsChecked == true;

            // NEW: Live Rule Statistics (Horizontal Layout)
            StackPanel statsPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(15, 15, 0, 15) };
            TextBlock txtTotalRules = new TextBlock { Text = Strings.Get("LblTotalRules", AutoOrganizeManager.Rules.Count), FontFamily = new FontFamily("Segoe UI"), FontSize = 13, FontWeight = FontWeights.Medium };
            TextBlock txtSeparator = new TextBlock { Text = "   -   ", FontFamily = new FontFamily("Segoe UI"), FontSize = 13, FontWeight = FontWeights.Medium, Foreground = Brushes.Gray };
            TextBlock txtEnabledRules = new TextBlock { Text = Strings.Get("LblEnabledRules", AutoOrganizeManager.Rules.Count(r => r.IsEnabled)), FontFamily = new FontFamily("Segoe UI"), FontSize = 13, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(Color.FromRgb(34, 139, 34)) };
            statsPanel.Children.Add(txtTotalRules);
            statsPanel.Children.Add(txtSeparator);
            statsPanel.Children.Add(txtEnabledRules);
            c.Children.Add(statsPanel);

            // Navy Blue - Manage Rules Button
            Button btnManageRules = CreateStyledButton(Strings.BtnSmartDesktopRules, Color.FromRgb(0, 0, 128));
            btnManageRules.Width = 255;
            btnManageRules.Height = 45;
            btnManageRules.Margin = new Thickness(15, 0, 0, 15);
            btnManageRules.HorizontalAlignment = HorizontalAlignment.Left;
            btnManageRules.Click += (s, e) =>
            {
                new AutoOrganizeForm().ShowDialog();
                // Refresh statistics when the editor closes
                txtTotalRules.Text = Strings.Get("LblTotalRules", AutoOrganizeManager.Rules.Count);
                txtEnabledRules.Text = Strings.Get("LblEnabledRules", AutoOrganizeManager.Rules.Count(r => r.IsEnabled));
            };
            c.Children.Add(btnManageRules);

            // Dark Red - Organize Desktop Now Button
            Button btnOrganizeNow = CreateStyledButton(Strings.BtnOrganizeNow, Color.FromRgb(139, 0, 0));
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
                Text = Strings.NoteAutoOrganize,
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
            CreateSectionHeader(c, Strings.SecLog, ColorLookDeeper);
            CreateCheckBox(c, Strings.OptEnableLogging, "EnableLogging", SettingsManager.IsLogEnabled);
            Button b = CreateStyledButton(Strings.BtnOpenLog, ColorLookDeeper); b.Width = 100; b.Height = 25; b.HorizontalAlignment = HorizontalAlignment.Left;
            b.Click += (s, e) => OpenLogFile();
            c.Children.Add(b);

            CreateSectionHeader(c, Strings.SecLogConfiguration, ColorLookDeeper);
            CreateLogLevelComboBox(c);
            CreateSectionHeader(c, Strings.SecLogCategories, ColorLookDeeper);

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
            Grid g = new Grid { Margin = new Thickness(15, 5, 0, 5) };

            // --- UI FIX: Increased to 205px to bump the sliders and "Effect" label to the right, adding a nice gap ---
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(205) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            // Slightly widened the third column from 50 to 65 to comfortably fit the NumericTextBox without clipping
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(65) });

            // UI FIX: Changed HorizontalAlignment to Left so labels sit flush on the left margin
            TextBlock lbl = new TextBlock { Text = l, FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Left, Margin = new Thickness(0, 0, 10, 0) };
            Slider sl = new Slider { Name = n, Minimum = 1, Maximum = max, Value = v, TickFrequency = 1, IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center };

            // --- TRIAL: Replaced TextBlock with interconnected NumericTextBox for micro-adjustments ---
            NumericTextBox nud = new NumericTextBox
            {
                Minimum = 1,
                Maximum = max,
                Value = v,
                Width = 55, // Exactly half of the standard width used in CustomizeFrameForm, plus a tiny bit for 3-digit numbers
                Height = 20, // Strict compact height to preserve the original line spacing and avoid blowing up vertical space
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(10, 0, 0, 0)
            };

            // Two-way binding between Slider and NumericTextBox
            sl.ValueChanged += (s, e) => { if (nud.Value != (int)e.NewValue) nud.Value = (int)e.NewValue; };
            nud.ValueChanged += (s, e) => { if (sl.Value != nud.Value) sl.Value = nud.Value; };

            Grid.SetColumn(lbl, 0); Grid.SetColumn(sl, 1); Grid.SetColumn(nud, 2);
            g.Children.Add(lbl); g.Children.Add(sl); g.Children.Add(nud);
            p.Children.Add(g);
        }

        private static void CreateIconRadioButtonGroup(StackPanel p, string gName, Dictionary<string, int> icons, int sel)
        {
            StackPanel sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(15, 5, 0, 15), Tag = gName };
            foreach (var i in icons) sp.Children.Add(new RadioButton { Content = i.Key, Tag = i.Value, GroupName = gName, IsChecked = i.Value == sel, Margin = new Thickness(0, 0, 15, 0), FontSize = 16, FontFamily = new FontFamily("Segoe UI Symbol") });
            p.Children.Add(sp);
        }

        //New Arrangement

        private static void CreateColorAndEffectComboBoxes(StackPanel p, CheckBox chamCb)
        {
            Grid g = new Grid { Margin = new Thickness(15, 10, 0, 10) };

            // --- UI FIX: Hardcode total to 205px (45 + 160) to match the sliders. 
            // The 160px column holds a 140px dropdown, leaving exactly a 20px gap before "Effect" ---
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(45) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });

            TextBlock lblColor = new TextBlock { Text = Strings.LblColor, FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 10, 0) };
            // Constrain Width to 140 and Left-align so it doesn't stretch to fill the 160px column, creating the gap automatically
            ComboBox cbColor = new ComboBox { Name = "ColorComboBox", Width = 140, HorizontalAlignment = HorizontalAlignment.Left, Height = 25, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            foreach (string c in new[] { "Gray", "Black", "White", "Beige", "Green", "Purple", "Fuchsia", "Yellow", "Orange", "Red", "Blue", "Bismark" })
                cbColor.Items.Add(new ComboBoxItem { Content = Strings.Get("Color" + c), Tag = c });
            cbColor.SelectedItem = cbColor.Items.OfType<ComboBoxItem>()
                .FirstOrDefault(i => string.Equals(i.Tag as string, SettingsManager.SelectedColor, StringComparison.OrdinalIgnoreCase))
                ?? cbColor.Items.OfType<ComboBoxItem>().FirstOrDefault();

            // --- BUG FIX: Disable Color dropdown if Chameleon mode is ON ---
            cbColor.IsEnabled = chamCb.IsChecked != true;
            chamCb.Click += (s, e) => cbColor.IsEnabled = chamCb.IsChecked != true;
            // ---------------------------------------------------------------

            // UI FIX: Starts perfectly flush at the new 205px mark
            TextBlock lblEffect = new TextBlock { Text = Strings.LblEffect, FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Left, Margin = new Thickness(0, 0, 10, 0) };
            ComboBox cbEffect = new ComboBox { Name = "LaunchEffectComboBox", Width = 140, HorizontalAlignment = HorizontalAlignment.Left, Height = 25, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            // Same order as the LaunchEffect enum: the value is the index, not the label.
            foreach (string e in new[] { Strings.FxZoom, Strings.FxBounce, Strings.FxFadeOut, Strings.FxSlideUp, Strings.FxRotate, Strings.FxAgitate, Strings.FxGrowAndFly, Strings.FxPulse, Strings.FxElastic, Strings.FxFlip3D, Strings.FxSpiral, Strings.FxShockwave, Strings.FxMatrix, Strings.FxSupernova, Strings.FxTeleport }) cbEffect.Items.Add(e);
            cbEffect.SelectedIndex = (int)SettingsManager.LaunchEffect;

            Grid.SetColumn(lblColor, 0); g.Children.Add(lblColor);
            Grid.SetColumn(cbColor, 1); g.Children.Add(cbColor);
            Grid.SetColumn(lblEffect, 2); g.Children.Add(lblEffect);
            Grid.SetColumn(cbEffect, 3); g.Children.Add(cbEffect);

            p.Children.Add(g);
        }

        private static Button CreateStyledButton(string t, Color c) => new Button { Content = t, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, FontWeight = FontWeights.Bold, Background = new SolidColorBrush(c), Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand, HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch };

        private static void CreateLogLevelComboBox(StackPanel p)
        {
            Grid g = new Grid { Margin = new Thickness(0, 10, 0, 10) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            g.Children.Add(new TextBlock { Text = Strings.LblMinimumLogLevel, FontSize = 13, VerticalAlignment = VerticalAlignment.Center });
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
                bool tempPortalImageState = SettingsManager.ShowBackgroundImageOnPortalFrames;
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
                        if (cb.Name == "EnableSnapNearFrames") SettingsManager.IsSnapEnabled = cb.IsChecked == true;
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

                        // Moved from Style Tab (Choices)
                        if (cb.Name == "EnablePortalWatermark") { newPortalWatermarkState = cb.IsChecked == true; SettingsManager.ShowBackgroundImageOnPortalFrames = newPortalWatermarkState; }
                        if (cb.Name == "DisableFrameScrollbars") SettingsManager.DisableFrameScrollbars = cb.IsChecked == true;
                        if (cb.Name == "EnableSounds") SettingsManager.EnableSounds = cb.IsChecked == true;
                    }

                    // --- NEW: Catch the Sound & Portal View Config Grids ---
                    else if (child is Grid genGrid)
                    {
                        var sndCombo = genGrid.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "NotificationSoundComboBox");
                        if (sndCombo != null)
                        {
                            SettingsManager.NotificationSound = sndCombo.SelectedIndex switch
                            {
                                1 => NotificationSound.DoubleDing,
                                2 => NotificationSound.SmoothTickle,
                                3 => NotificationSound.MessageDing,
                                4 => NotificationSound.GentleDing,
                                5 => NotificationSound.SoftDing,
                                _ => NotificationSound.DefaultSound
                            };
                        }

                        var pvCombo = genGrid.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "DefaultPortalViewComboBox");
                        if (pvCombo?.SelectedItem != null)
                        {
                            SettingsManager.DefaultPortalView = (pvCombo.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? SettingsManager.DefaultPortalView;
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
                        if (cb.Name == "ApplyTintToIcons") SettingsManager.ApplyTintToIcons = cb.IsChecked == true;
                        // NEW: Auto-Hide & Fade Options (Moved from General)
                        if (cb.Name == "AutoHideFrames") { SettingsManager.AutoHideFrames = cb.IsChecked == true; Framemanager.ResetAutoHideTimer(); }


                        if (cb.Name == "FramesFadeOutFx") SettingsManager.FramesFadeOutFx = cb.IsChecked == true;

                        // NEW: Desktop Icon Visibility
                        if (cb.Name == "HideDesktopElementsOnStart")
                        {
                            SettingsManager.HideDesktopElementsOnStart = cb.IsChecked == true;

                            // Immediately apply this state upon saving
                            DesktopIconManager.SetDesktopIconsVisible(!SettingsManager.HideDesktopElementsOnStart);
                        }
                        if (cb.Name == "HideDesktopElementsOnAllFramesHide") SettingsManager.HideDesktopElementsOnAllFramesHide = cb.IsChecked == true;
                    }
                    else if (child is Grid g)
                    {
                        var tint = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "TintSlider"); if (tint != null) SettingsManager.TintValue = (int)tint.Value; var mtint = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "MenuTintSlider"); if (mtint != null) SettingsManager.MenuTintValue = (int)mtint.Value;
                        var col = g.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "ColorComboBox"); if ((col?.SelectedItem as ComboBoxItem)?.Tag is string colTag) SettingsManager.SelectedColor = colTag;
                        var eff = g.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "LaunchEffectComboBox"); if (eff != null) SettingsManager.LaunchEffect = (LaunchEffectsManager.LaunchEffect)eff.SelectedIndex;

                        // Parse Sliders (Moved from General)
                        var autoHideTime = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "AutoHideTimeSlider"); if (autoHideTime != null) { SettingsManager.AutoHideTime = (int)autoHideTime.Value; Framemanager.ResetAutoHideTimer(); }
                        var fadeOutTime = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "FadeOutTimeSlider"); if (fadeOutTime != null) SettingsManager.FadeOutTime = (int)fadeOutTime.Value;
                        var fadeOutAlpha = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "FadeOutAlphaSlider"); if (fadeOutAlpha != null) SettingsManager.FadeOutFxTargetAlpha = fadeOutAlpha.Value / 100.0;
                        var autoRollTime = g.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "AutoRollTimeSlider"); if (autoRollTime != null) SettingsManager.AutoRollTime = (int)autoRollTime.Value;

                        // Parse Icons from the new side-by-side nested Grid layout
                        foreach (var innerChild in g.Children.OfType<StackPanel>())
                        {
                            foreach (var rbSp in innerChild.Children.OfType<StackPanel>())
                            {
                                if (rbSp.Tag?.ToString() == "MenuIconGroup") foreach (RadioButton rb in rbSp.Children.OfType<RadioButton>()) if (rb.IsChecked == true) SettingsManager.MenuIcon = (int)rb.Tag;
                                if (rbSp.Tag?.ToString() == "LockIconGroup") foreach (RadioButton rb in rbSp.Children.OfType<RadioButton>()) if (rb.IsChecked == true) SettingsManager.LockIcon = (int)rb.Tag;
                            }
                        }
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
                        if (hotkeyCb.Name == "EnableFocusFrameHotkey" && SettingsManager.EnableFocusFrameHotkey != (hotkeyCb.IsChecked == true)) { SettingsManager.EnableFocusFrameHotkey = hotkeyCb.IsChecked == true; hotkeysChanged = true; }
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
                            if (prefix == "FocusFrame") { if (SettingsManager.FocusFrameModifier != modString || SettingsManager.FocusFrameKey != key) { SettingsManager.FocusFrameModifier = modString; SettingsManager.FocusFrameKey = key; hotkeysChanged = true; } }
                            if (prefix == "SpotSearch") { if (SettingsManager.SpotSearchModifier != modString || SettingsManager.SpotSearchKey != key) { SettingsManager.SpotSearchModifier = modString; SettingsManager.SpotSearchKey = key; hotkeysChanged = true; } }
                        }
                    }
                }

                if (hotkeysChanged)
                {
                    // Propagate the new hotkeys across all existing profiles
                    SettingsManager.BroadcastHotkeysToAllProfiles();

                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Global Hotkey changes have been saved and applied to all profiles.\n\nPlease restart Desktop Frames to activate the new shortcuts.", "Restart Required");
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

                if (tempPortalImageState != newPortalWatermarkState) TrayManager.reloadallFrames();
                TrayManager.Instance?.UpdateTrayIcon();
                Utility.UpdateFrameVisuals();

				// --- NEW: Broadcast Idle Fade-Out Settings to all active frames ---
				var allFrames = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                foreach (var frame in allFrames)
                {
					frame.RefreshIdleSettings();
                }

                // --- NEW: Broadcast Auto-Roll Settings ---
                Framemanager.RefreshAutoRollSettings();

                // --- NEW: Broadcast Icon Tint Settings ---
                Framemanager.RefreshAllIconsTint();

                // --- NEW: Broadcast Scrollbar Settings ---
                Framemanager.RefreshScrollbarSettings();

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

            Button c = new Button { Content = Strings.BtnCancel, Width = 100, Height = 34, FontWeight = FontWeights.Bold, Background = Brushes.White, BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Margin = new Thickness(0, 0, 10, 0), Cursor = Cursors.Hand };
            c.Click += (s, e) => _optionsWindow.Close();

            Button sv = new Button { Content = Strings.BtnSave, Width = 100, Height = 34, FontWeight = FontWeights.Bold, Background = new SolidColorBrush(_userAccentColor), Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
            sv.Click += (s, e) => SaveOptions();

            sp.Children.Add(c); sp.Children.Add(sv); f.Child = sp; mainGrid.Children.Add(f);
        }

        private static void CreateDonationSection(Grid mainGrid)
        {
            Border d = new Border { Background = new SolidColorBrush(Color.FromRgb(255, 248, 225)), BorderBrush = new SolidColorBrush(Color.FromRgb(255, 193, 7)), BorderThickness = new Thickness(0, 1, 0, 0), Padding = new Thickness(20) };
            Grid.SetRow(d, 3);
            StackPanel sp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
            sp.Children.Add(new TextBlock { Text = Strings.LblDonate, FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(102, 77, 3)), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 15, 0) });
            Button b = new Button { Content = Strings.BtnDonate, FontSize = 14, Background = new SolidColorBrush(Color.FromRgb(255, 193, 7)), Foreground = Brushes.White, BorderThickness = new Thickness(0), Padding = new Thickness(15, 6, 15, 6), Cursor = Cursors.Hand };
            b.Click += (s, e) => { try { Process.Start(new ProcessStartInfo { FileName = "https://www.paypal.com/donate/?hosted_button_id=PPLWC66UC8Q42", UseShellExecute = true }); } catch { } };
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
                        TrayManager.reloadallFrames();
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
                string p = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Frames.log");
                if (System.IO.File.Exists(p)) Process.Start(new ProcessStartInfo { FileName = p, UseShellExecute = true });
                else MessageBoxesManager.ShowOKOnlyMessageBoxForm("Log file not found.", "Information");
            }
            catch { }
        }

        private static void PerformFullFactoryReset()
        {
            if (MessageBoxesManager.ShowCustomYesNoMessageBox("WARNING: This will delete ALL frames, shortcuts, and settings for the CURRENT PROFILE!\n\nAre you sure you want to proceed?", "Factory Reset"))
            {
                // KISS: Hijack cursor to show processing
                System.Windows.Application.Current?.Dispatcher.Invoke(() => System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait);
                try
                {
                    // 1. Create a safety backup before wiping
                    string ts = DateTime.Now.ToString("yyMMddHHmm");
                    BackupManager.CreateBackup($"{ts}_backup_reset", silent: true);

                    // 2. Wipe Profile-Specific Folders
                    foreach (string f in new[] { "Temp Shortcuts", "Shortcuts", "Last Frame Deleted", "CopiedItem" })
                    {
                        string p = ProfileManager.GetProfileFilePath(f);
                        if (System.IO.Directory.Exists(p))
                        {
                            try
                            {
                                System.IO.Directory.Delete(p, true);
                                System.IO.Directory.CreateDirectory(p); // Recreate empty folder
                            }
                            catch { }
                        }
                    }

                    // 3. Wipe Profile-Specific Config Files (OVERWRITE INSTEAD OF DELETE)
                    // FIX: Pointed to frames.json and wrote empty array to prevent read crashes
                    string fj = ProfileManager.GetProfileFilePath("frames.json");
                    System.IO.File.WriteAllText(fj, "[]");

                    string oj = ProfileManager.GetProfileFilePath("options.json");
                    System.IO.File.WriteAllText(oj, "{}");

                    // 4. Force a clean OS-level restart (Guarantees all UI clears properly)
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("Factory Reset complete.\nThe application will now restart.", "Reset Successful");

                    string appPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = $"/c ping 127.0.0.1 -n 3 > nul & start \"\" \"{appPath}\"",
                        WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                        CreateNoWindow = true,
                        UseShellExecute = false
                    });

                    // Instantly kill current process
                    System.Diagnostics.Process.GetCurrentProcess().Kill();
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Error, $"Factory reset failed: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Reset failed: {ex.Message}", "Error");
                }
                finally
                {
                    System.Windows.Application.Current?.Dispatcher.Invoke(() => System.Windows.Input.Mouse.OverrideCursor = null);
                }
            }
        }
    }
}