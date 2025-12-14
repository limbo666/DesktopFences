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


        private static int _lastSelectedTabIndex = 0; // Default to 0 (General)


        private static TabControl _tabControl;
        private static Window _optionsWindow;
        private static Color _userAccentColor;

        // Colors for tabs
        private static readonly Color ColorGeneral = Utility.GetColorFromName(SettingsManager.SelectedColor); // Dynamic
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
                    Width = 800, // Slightly wider for side-by-side options if needed
                    Height = 850,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStyle = WindowStyle.None,
                    Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                    AllowsTransparency = true
                };

                // Set icon
                try
                {
                    _optionsWindow.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { }

                Border mainBorder = new Border
                {
                    Background = Brushes.White,
                    CornerRadius = new CornerRadius(0),
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
                Grid.SetColumn(titleBlock, 0);

                Button closeButton = new Button
                {
                    Content = "✕",
                    Width = 32,
                    Height = 32,
                    FontSize = 14,
                    Foreground = Brushes.White,
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand
                };
                closeButton.Click += (s, e) => _optionsWindow.Close();
                Grid.SetColumn(closeButton, 1);

                headerGrid.Children.Add(titleBlock);
                headerGrid.Children.Add(closeButton);
                headerBorder.Child = headerGrid;
                headerBorder.MouseLeftButtonDown += (sender, e) => { if (e.ButtonState == MouseButtonState.Pressed) _optionsWindow.DragMove(); };

                CreateTabContent(mainGrid);
                CreateFooter(mainGrid);
                CreateDonationSection(mainGrid);

                mainGrid.Children.Add(headerBorder);
                mainBorder.Child = mainGrid;
                _optionsWindow.Content = mainBorder;

                _optionsWindow.KeyDown += OptionsForm_KeyDown;
                _optionsWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing Options form: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"An error occurred: {ex.Message}", "Error");
            }
        }

        private static void OptionsForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) { SaveOptions(); e.Handled = true; }
            else if (e.Key == Key.Escape) { _optionsWindow.Close(); e.Handled = true; }
        }

        private static void CreateTabContent(Grid mainGrid)
        {
            Grid contentGrid = new Grid();
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) });
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Grid.SetRow(contentGrid, 1);

            StackPanel tabPanel = new StackPanel { Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)), Margin = new Thickness(0, 20, 0, 0) };
            Grid.SetColumn(tabPanel, 0);

            Border contentBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1, 0, 0, 0),
                Padding = new Thickness(20),
                Margin = new Thickness(0, 20, 0, 0)
            };
            Grid.SetColumn(contentBorder, 1);

            _tabControl = new TabControl { Background = Brushes.Transparent, BorderThickness = new Thickness(0) };
            var template = new ControlTemplate(typeof(TabControl));
            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.SetValue(ContentPresenter.ContentSourceProperty, "SelectedContent");
            template.VisualTree = contentPresenter;
            _tabControl.Template = template;

            // Create Tabs
            CreateGeneralTab();
            CreateStyleTab();
            CreateToolsTab();
            CreateLookDeeperTab();

            // --- CHANGE: Set the index from the static variable ---
            _tabControl.SelectedIndex = _lastSelectedTabIndex;

            // --- CHANGE: Set the button state based on the variable ---
            CreateTabButton(tabPanel, "General", 0, _lastSelectedTabIndex == 0);
            CreateTabButton(tabPanel, "Style", 1, _lastSelectedTabIndex == 1);
            CreateTabButton(tabPanel, "Tools", 2, _lastSelectedTabIndex == 2);
            CreateTabButton(tabPanel, "Look Deeper", 3, _lastSelectedTabIndex == 3);

            contentBorder.Child = _tabControl;
            contentGrid.Children.Add(tabPanel);
            contentGrid.Children.Add(contentBorder);
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
            Color activeColor;
            switch (title)
            {
                case "General": activeColor = _userAccentColor; break;
                case "Style": activeColor = ColorStyle; break;
                case "Tools": activeColor = ColorTools; break;
                case "Look Deeper": activeColor = ColorLookDeeper; break;
                default: activeColor = Colors.Gray; break;
            }

            if (isSelected)
            {
                button.Background = new SolidColorBrush(activeColor);
                button.Foreground = Brushes.White;
            }
            else if (isHover)
            {
                button.Background = new SolidColorBrush(Color.FromRgb((byte)(activeColor.R + 40), (byte)(activeColor.G + 40), (byte)(activeColor.B + 40)));
                button.Foreground = Brushes.White;
            }
            else
            {
                button.Background = new SolidColorBrush(Color.FromRgb(200, 200, 200));
                button.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 60));
            }
        }

        private static void SelectTab(int tabIndex, Button selectedButton)
        {
            _lastSelectedTabIndex = tabIndex;
            _tabControl.SelectedIndex = tabIndex;
            StackPanel tabPanel = (StackPanel)selectedButton.Parent;
            for (int i = 0; i < tabPanel.Children.Count; i++)
            {
                if (tabPanel.Children[i] is Button btn)
                {
                    SetTabButtonColors(btn, btn.Content.ToString(), i == tabIndex);
                }
            }
        }

        private static void CreateGeneralTab()
        {
            TabItem generalTab = new TabItem();
            ScrollViewer scrollViewer = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            StackPanel content = new StackPanel();

            // Startup
            CreateSectionHeader(content, "Startup", _userAccentColor);
            CreateCheckBox(content, "Start with Windows", "StartWithWindows", TrayManager.IsStartWithWindows);

            // Selections
            CreateSectionHeader(content, "Selections", _userAccentColor);
            CreateCheckBox(content, "Single Click to Launch", "SingleClickToLaunch", SettingsManager.SingleClickToLaunch);
            CreateCheckBox(content, "Enable Snap Near Fences", "EnableSnapNearFences", SettingsManager.IsSnapEnabled);
            CreateCheckBox(content, "Enable Dimension Snap", "EnableDimensionSnap", SettingsManager.EnableDimensionSnap);
            CreateCheckBox(content, "Enable Tray Icon", "EnableTrayIcon", SettingsManager.ShowInTray);
            CreateCheckBox(content, "Use Recycle Bin on Portal Fences 'Delete item' command", "UseRecycleBin", SettingsManager.UseRecycleBin);

            scrollViewer.Content = content;
            generalTab.Content = scrollViewer;
            _tabControl.Items.Add(generalTab);
        }

        private static void CreateStyleTab()
        {
            TabItem styleTab = new TabItem();
            ScrollViewer scrollViewer = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            StackPanel content = new StackPanel();

            // Choices
            CreateSectionHeader(content, "Choices", ColorStyle);
            CreateCheckBox(content, "Enable Portal Fences Watermark", "EnablePortalWatermark", SettingsManager.ShowBackgroundImageOnPortalFences);

            // Future implementation logic for Note Watermark
            var noteWatermark = CreateCheckBoxReturn(content, "Enable Note Fences Watermark (Coming Soon)", "EnableNoteWatermark", false);
            noteWatermark.IsEnabled = false; // Placeholder status
            noteWatermark.Foreground = Brushes.Gray;

            CreateCheckBox(content, "Disable Fence Scrollbars", "DisableFenceScrollbars", SettingsManager.DisableFenceScrollbars);
            CreateCheckBox(content, "Enable Sounds", "EnableSounds", SettingsManager.EnableSounds);

            // Appearance
            CreateSectionHeader(content, "Appearance", ColorStyle);
            CreateSliderControl(content, "Fence Tint", "TintSlider", SettingsManager.TintValue);
            CreateSliderControl(content, "Menu Tint", "MenuTintSlider", SettingsManager.MenuTintValue);
            CreateColorComboBox(content);
            CreateLaunchEffectComboBox(content);

            // Icons
            CreateSectionHeader(content, "Icons", ColorStyle);

            // Menu Icon Radio Buttons
            content.Children.Add(new TextBlock { Text = "Menu Icon", FontWeight = FontWeights.SemiBold, Margin = new Thickness(15, 5, 0, 5) });
            var menuIcons = new Dictionary<string, int> { { "♥", 0 }, { "☰", 1 }, { "≣", 2 }, { "𓃑", 3 } };
            CreateIconRadioButtonGroup(content, "MenuIconGroup", menuIcons, SettingsManager.MenuIcon);

            // Lock Icon Radio Buttons
            content.Children.Add(new TextBlock { Text = "Lock Icon", FontWeight = FontWeights.SemiBold, Margin = new Thickness(15, 10, 0, 5) });
            var lockIcons = new Dictionary<string, int> { { "🛡️", 0 }, { "🔑", 1 }, { "🔐", 2 }, { "🔒", 3 } };
            CreateIconRadioButtonGroup(content, "LockIconGroup", lockIcons, SettingsManager.LockIcon);

            scrollViewer.Content = content;
            styleTab.Content = scrollViewer;
            _tabControl.Items.Add(styleTab);
        }



        private static void CreateToolsTab()
        {
            TabItem toolsTab = new TabItem();
            StackPanel content = new StackPanel();

            CreateSectionHeader(content, "Tools", ColorTools);

            // --- Grid for Standard Tools ---
            // Width Calculation: 120 + 15 + 120 = 255 Total Width
            // Row Height: 45
            Grid buttonGrid = new Grid { Margin = new Thickness(0, 10, 0, 0) };
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(15) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            buttonGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) });
            buttonGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(15) });
            buttonGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) });

            Button backupButton = CreateStyledButton("Backup", ColorTools);
            backupButton.Click += (s, e) => BackupManager.BackupData();
            Grid.SetRow(backupButton, 0); Grid.SetColumn(backupButton, 0);

            Button restoreButton = CreateStyledButton("Restore...", Color.FromRgb(255, 152, 0));
            restoreButton.Click += (s, e) => RestoreBackup();
            Grid.SetRow(restoreButton, 0); Grid.SetColumn(restoreButton, 2);

            Button openBackupsButton = CreateStyledButton("Open Backups Folder", Color.FromRgb(0, 123, 191));
            openBackupsButton.Click += (s, e) => OpenBackupsFolder();
            Grid.SetRow(openBackupsButton, 2); Grid.SetColumn(openBackupsButton, 0); Grid.SetColumnSpan(openBackupsButton, 3);

            buttonGrid.Children.Add(backupButton);
            buttonGrid.Children.Add(restoreButton);
            buttonGrid.Children.Add(openBackupsButton);
            content.Children.Add(buttonGrid);

            // --- 1. Automatic Backup Checkbox ---
            CreateCheckBox(content, "Automatic Backup (Daily)", "EnableAutoBackup", SettingsManager.EnableAutoBackup);

            // --- 2. Reset Section ---
            CreateSectionHeader(content, "Reset", Colors.Red);

            StackPanel resetStack = new StackPanel
            {
                Orientation = Orientation.Vertical,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 0)
            };

            // Dimensions to match "Open Backups Folder" exactly
            double uniformWidth = 255;
            double uniformHeight = 45;

            // [Reset all Customizations]
            Button resetCustomButton = CreateStyledButton("Reset Styles", Color.FromRgb(108, 117, 125)); // Gray
            resetCustomButton.Width = uniformWidth;
            resetCustomButton.Height = uniformHeight;
            resetCustomButton.Margin = new Thickness(0, 0, 0, 15); // Gap between buttons
            resetCustomButton.ToolTip = "Resets colors, fonts, and sizes to default. Content remains safe.";
            resetCustomButton.Click += (s, e) =>
            {
                if (MessageBoxesManager.ShowCustomYesNoMessageBox("Reset all visual customizations to default?\nYour icons and fences will remain.", "Reset Styles"))
                {
                    FenceManager.ResetAllCustomizations();
                    _optionsWindow.Close();
                }
            };

            // [Clear all data]
            Button clearDataButton = CreateStyledButton("Clear All Data", Color.FromRgb(220, 53, 69)); // Red
            clearDataButton.Width = uniformWidth;
            clearDataButton.Height = uniformHeight;
            clearDataButton.ToolTip = "WARNING: Deletes all fences, shortcuts, and settings.";
            clearDataButton.Click += (s, e) => PerformFullFactoryReset();

            resetStack.Children.Add(resetCustomButton);
            resetStack.Children.Add(clearDataButton);
            content.Children.Add(resetStack);

            toolsTab.Content = content;
            _tabControl.Items.Add(toolsTab);
        }
        private static void PerformFullFactoryReset()
        {
            bool confirm = MessageBoxesManager.ShowCustomYesNoMessageBox(
                "WARNING: This will delete ALL fences, shortcuts, and settings!\n\n" +
                "Are you sure you want to proceed?",
                "Factory Reset");

            if (confirm)
            {
                try
                {
                    // 1. Silent Backup (Unique name, no success popup)
                    string timestamp = DateTime.Now.ToString("yyMMddHHmm");
                    string backupName = $"{timestamp}_backup_reset";
                    BackupManager.CreateBackup(backupName, silent: true);

                    // 2. Clean Folders
                    string exeDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                    string[] foldersToClean = { "Temp Shortcuts", "Shortcuts", "Last Fence Deleted", "CopiedItem" };

                    foreach (string folder in foldersToClean)
                    {
                        string path = System.IO.Path.Combine(exeDir, folder);
                        if (System.IO.Directory.Exists(path))
                        {
                            try
                            {
                                System.IO.Directory.Delete(path, true); // Recursive delete
                                System.IO.Directory.CreateDirectory(path); // Recreate empty
                            }
                            catch (Exception ex)
                            {
                                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Failed to clean folder {folder}: {ex.Message}");
                            }
                        }
                    }

                    // 3. Reset JSON Files (Delete them)
                    string fencesJson = System.IO.Path.Combine(exeDir, "fences.json");
                    string optionsJson = System.IO.Path.Combine(exeDir, "options.json");

                    if (System.IO.File.Exists(fencesJson)) System.IO.File.Delete(fencesJson);
                    if (System.IO.File.Exists(optionsJson)) System.IO.File.Delete(optionsJson);

                    // 4. Reload Application State
                    // A. Reset Settings in memory (loads defaults since file is gone)
                    SettingsManager.LoadSettings();

                    // B. Reload Fences (destroys old windows, creates defaults)
                    FenceManager.ReloadFences();

                    // C. Close Options Window (it reflects old state)
                    _optionsWindow.Close();

                    // Optional: Minimal confirmation
                    // MessageBoxesManager.ShowOKOnlyMessageBoxForm("Factory reset complete.", "Reset");
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Error, $"Factory reset failed: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Reset failed: {ex.Message}", "Error");
                }
            }
        }

        private static void CreateLookDeeperTab()
        {
            TabItem lookDeeperTab = new TabItem();
            ScrollViewer scrollViewer = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            StackPanel content = new StackPanel();

            CreateSectionHeader(content, "Log", ColorLookDeeper);
            CreateCheckBox(content, "Enable logging", "EnableLogging", SettingsManager.IsLogEnabled);

            Button openLogButton = CreateStyledButton("Open Log", ColorLookDeeper);
            openLogButton.Width = 100; openLogButton.Height = 25; openLogButton.HorizontalAlignment = HorizontalAlignment.Left;
            openLogButton.Click += (s, e) => OpenLogFile();
            content.Children.Add(openLogButton);

            CreateSectionHeader(content, "Log configuration", ColorLookDeeper);
            CreateLogLevelComboBox(content);

            CreateSectionHeader(content, "Log Categories", ColorLookDeeper);
            CreateLogCategoryCheckBoxes(content);
            CreateCheckBox(content, "Enable Background Validation Logging", "EnableBackgroundValidation", SettingsManager.EnableBackgroundValidationLogging);

            scrollViewer.Content = content;
            lookDeeperTab.Content = scrollViewer;
            _tabControl.Items.Add(lookDeeperTab);
        }

        // --- Helper Methods ---

        private static void CreateSectionHeader(StackPanel parent, string title, Color headerColor)
        {
            parent.Children.Add(new TextBlock
            {
                Text = title,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(headerColor),
                Margin = new Thickness(0, 10, 0, 15)
            });
        }

        private static void CreateCheckBox(StackPanel parent, string text, string name, bool isChecked)
        {
            parent.Children.Add(new CheckBox
            {
                Name = name,
                Content = text,
                IsChecked = isChecked,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Margin = new Thickness(15, 8, 0, 8)
            });
        }

        private static CheckBox CreateCheckBoxReturn(StackPanel parent, string text, string name, bool isChecked)
        {
            var cb = new CheckBox
            {
                Name = name,
                Content = text,
                IsChecked = isChecked,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Margin = new Thickness(15, 8, 0, 8)
            };
            parent.Children.Add(cb);
            return cb;
        }

        private static void CreateSliderControl(StackPanel parent, string labelText, string sliderName, int currentValue)
        {
            Grid sliderGrid = new Grid();
            sliderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) }); // Label
            sliderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) }); // Slider
            sliderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(50) });  // Value
            sliderGrid.Margin = new Thickness(0, 5, 0, 5);

            TextBlock label = new TextBlock { Text = labelText, FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 10, 0) };
            Grid.SetColumn(label, 0);

            Slider slider = new Slider { Name = sliderName, Minimum = 1, Maximum = 100, Value = currentValue, TickFrequency = 1, IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(slider, 1);

            TextBlock valueDisplay = new TextBlock { Text = currentValue.ToString(), FontSize = 13, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(15, 0, 0, 0) };
            Grid.SetColumn(valueDisplay, 2);

            slider.ValueChanged += (s, e) => valueDisplay.Text = ((int)e.NewValue).ToString();

            sliderGrid.Children.Add(label);
            sliderGrid.Children.Add(slider);
            sliderGrid.Children.Add(valueDisplay);
            parent.Children.Add(sliderGrid);
        }

        private static void CreateIconRadioButtonGroup(StackPanel parent, string groupName, Dictionary<string, int> icons, int selectedIndex)
        {
            StackPanel panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(15, 5, 0, 15), Tag = groupName };

            foreach (var icon in icons)
            {
                RadioButton rb = new RadioButton
                {
                    Content = icon.Key,
                    Tag = icon.Value, // Store int value in Tag
                    GroupName = groupName,
                    IsChecked = icon.Value == selectedIndex,
                    Margin = new Thickness(0, 0, 15, 0),
                    FontSize = 16,
                    FontFamily = new FontFamily("Segoe UI Symbol") // Ensure symbol support
                };
                panel.Children.Add(rb);
            }
            parent.Children.Add(panel);
        }

        private static void CreateColorComboBox(StackPanel parent)
        {
            Grid colorGrid = new Grid { Margin = new Thickness(0, 10, 0, 10) };
            colorGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            colorGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });

            TextBlock label = new TextBlock { Text = "Color", FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 10, 0) };
            Grid.SetColumn(label, 0);

            ComboBox comboBox = new ComboBox { Name = "ColorComboBox", Height = 25, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            string[] colors = { "Gray", "Black", "White", "Beige", "Green", "Purple", "Fuchsia", "Yellow", "Orange", "Red", "Blue", "Bismark" };
            foreach (string color in colors) comboBox.Items.Add(color);
            comboBox.SelectedItem = SettingsManager.SelectedColor;
            Grid.SetColumn(comboBox, 1);

            colorGrid.Children.Add(label);
            colorGrid.Children.Add(comboBox);
            parent.Children.Add(colorGrid);
        }

        private static void CreateLaunchEffectComboBox(StackPanel parent)
        {
            Grid effectGrid = new Grid { Margin = new Thickness(0, 10, 0, 10) };
            effectGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            effectGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });

            TextBlock label = new TextBlock { Text = "Effect", FontSize = 13, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 10, 0) };
            Grid.SetColumn(label, 0);

            ComboBox comboBox = new ComboBox { Name = "LaunchEffectComboBox", Height = 25, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            string[] effects = { "Zoom", "Bounce", "FadeOut", "SlideUp", "Rotate", "Agitate", "GrowAndFly", "Pulse", "Elastic", "Flip3D", "Spiral", "Shockwave", "Matrix", "Supernova", "Teleport" };
            foreach (string effect in effects) comboBox.Items.Add(effect);
            comboBox.SelectedIndex = (int)SettingsManager.LaunchEffect;
            Grid.SetColumn(comboBox, 1);

            effectGrid.Children.Add(label);
            effectGrid.Children.Add(comboBox);
            parent.Children.Add(effectGrid);
        }

        private static Button CreateStyledButton(string text, Color color)
        {
            return new Button
            {
                Content = text,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(color),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
        }

        private static void SaveOptions()
        {
            try
            {
                bool tempPortalImageState = SettingsManager.ShowBackgroundImageOnPortalFences;
                bool newPortalWatermarkState = false;
                bool newShowInTrayState = false;

                // --- 1. General Tab ---
                TabItem generalTab = (TabItem)_tabControl.Items[0];
                StackPanel generalContent = (StackPanel)((ScrollViewer)generalTab.Content).Content;

                foreach (var child in generalContent.Children)
                {
                    if (child is CheckBox checkBox)
                    {
                        switch (checkBox.Name)
                        {
                            case "StartWithWindows":
                                if (checkBox.IsChecked != TrayManager.IsStartWithWindows)
                                    TrayManager.Instance?.ToggleStartWithWindows(checkBox.IsChecked == true);
                                break;
                            case "SingleClickToLaunch": SettingsManager.SingleClickToLaunch = checkBox.IsChecked == true; break;
                            case "EnableSnapNearFences": SettingsManager.IsSnapEnabled = checkBox.IsChecked == true; break;
                            case "EnableDimensionSnap": SettingsManager.EnableDimensionSnap = checkBox.IsChecked == true; break;
                            case "EnableTrayIcon":
                                newShowInTrayState = checkBox.IsChecked == true;
                                SettingsManager.ShowInTray = newShowInTrayState;
                                if (TrayManager.Instance != null)
                                    TrayManager.Instance.GetType().GetField("Showintray", BindingFlags.NonPublic | BindingFlags.Instance)?.SetValue(TrayManager.Instance, newShowInTrayState);
                                break;
                            case "UseRecycleBin": SettingsManager.UseRecycleBin = checkBox.IsChecked == true; break;
                        }
                    }
                }

                // --- 2. Style Tab ---
                TabItem styleTab = (TabItem)_tabControl.Items[1];
                StackPanel styleContent = (StackPanel)((ScrollViewer)styleTab.Content).Content;

                foreach (var child in styleContent.Children)
                {
                    if (child is CheckBox checkBox)
                    {
                        switch (checkBox.Name)
                        {
                            case "EnablePortalWatermark":
                                newPortalWatermarkState = checkBox.IsChecked == true;
                                SettingsManager.ShowBackgroundImageOnPortalFences = newPortalWatermarkState;
                                break;
                            case "DisableFenceScrollbars": SettingsManager.DisableFenceScrollbars = checkBox.IsChecked == true; break;
                            case "EnableSounds": SettingsManager.EnableSounds = checkBox.IsChecked == true; break;
                                // Case for NoteWatermark reserved for future use
                        }
                    }
                    else if (child is Grid grid)
                    {
                        var tintSlider = grid.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "TintSlider");
                        if (tintSlider != null) SettingsManager.TintValue = (int)tintSlider.Value;

                        var menuTintSlider = grid.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "MenuTintSlider");
                        if (menuTintSlider != null) SettingsManager.MenuTintValue = (int)menuTintSlider.Value;

                        var colorCombo = grid.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "ColorComboBox");
                        if (colorCombo?.SelectedItem != null) SettingsManager.SelectedColor = colorCombo.SelectedItem.ToString();

                        var effectCombo = grid.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "LaunchEffectComboBox");
                        if (effectCombo != null) SettingsManager.LaunchEffect = (LaunchEffectsManager.LaunchEffect)effectCombo.SelectedIndex;
                    }
                    else if (child is StackPanel stackPanel)
                    {
                        // Handle Radio Buttons for Icons
                        if (stackPanel.Tag?.ToString() == "MenuIconGroup")
                        {
                            foreach (RadioButton rb in stackPanel.Children.OfType<RadioButton>())
                                if (rb.IsChecked == true) SettingsManager.MenuIcon = (int)rb.Tag;
                        }
                        else if (stackPanel.Tag?.ToString() == "LockIconGroup")
                        {
                            foreach (RadioButton rb in stackPanel.Children.OfType<RadioButton>())
                                if (rb.IsChecked == true) SettingsManager.LockIcon = (int)rb.Tag;
                        }
                    }
                }


                // --- 3a. Tools Tab ---
                TabItem toolsTabItem = (TabItem)_tabControl.Items[2];
                if (toolsTabItem.Content is StackPanel toolsContent)
                {
                    foreach (var child in toolsContent.Children)
                    {
                        if (child is CheckBox checkBox && checkBox.Name == "EnableAutoBackup")
                        {
                            SettingsManager.EnableAutoBackup = checkBox.IsChecked == true;
                        }
                    }
                }


                // --- 3b. Look Deeper Tab ---
                TabItem lookDeeperTab = (TabItem)_tabControl.Items[3]; // Index 3 now
                StackPanel lookDeeperContent = (StackPanel)((ScrollViewer)lookDeeperTab.Content).Content;
                var newEnabledCategories = new List<LogManager.LogCategory>();

                foreach (var child in lookDeeperContent.Children)
                {
                    if (child is CheckBox checkBox)
                    {
                        if (checkBox.Name == "EnableLogging") SettingsManager.IsLogEnabled = checkBox.IsChecked == true;
                        if (checkBox.Name == "EnableBackgroundValidation") SettingsManager.EnableBackgroundValidationLogging = checkBox.IsChecked == true;
                    }
                    else if (child is Grid grid)
                    {
                        var logLevelCombo = grid.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "LogLevelComboBox");
                        if (logLevelCombo?.SelectedItem != null && Enum.TryParse<LogManager.LogLevel>(logLevelCombo.SelectedItem.ToString(), out var logLevel))
                            SettingsManager.SetMinLogLevel(logLevel);

                        if (grid.Name == "LogCategoryGrid") ProcessLogCategoryCheckboxes(grid, newEnabledCategories);
                    }
                }

                // Save and Apply
                SettingsManager.SetEnabledLogCategories(newEnabledCategories);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.Settings, "Options saved successfully");

                if (tempPortalImageState != newPortalWatermarkState)
                { 
                TrayManager.reloadAllFences();
            }
                TrayManager.Instance?.UpdateTrayIcon();


                // --- NEW UPDATE CALL ---
                Utility.UpdateFenceVisuals();
    

                _optionsWindow.Close();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings, $"Error saving options: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error saving settings: {ex.Message}", "Save Error");
            }
        }

        // --- Reused Helpers (Log Checkboxes, etc) ---
        private static void CreateLogLevelComboBox(StackPanel parent)
        {
            Grid levelGrid = new Grid { Margin = new Thickness(0, 10, 0, 10) };
            levelGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            levelGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });

            TextBlock label = new TextBlock { Text = "Minimum Log Level", FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(label, 0);

            ComboBox comboBox = new ComboBox { Name = "LogLevelComboBox", Height = 25, FontSize = 13, VerticalAlignment = VerticalAlignment.Center };
            comboBox.Items.Add("Debug"); comboBox.Items.Add("Info"); comboBox.Items.Add("Warn"); comboBox.Items.Add("Error");
            comboBox.SelectedItem = SettingsManager.MinLogLevel.ToString();
            Grid.SetColumn(comboBox, 1);

            levelGrid.Children.Add(label); levelGrid.Children.Add(comboBox);
            parent.Children.Add(levelGrid);
        }

        private static void CreateLogCategoryCheckBoxes(StackPanel parent)
        {
            Grid categoryGrid = new Grid { Name = "LogCategoryGrid" };
            categoryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            categoryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            StackPanel leftColumn = new StackPanel();
            StackPanel rightColumn = new StackPanel();

            var logCategories = new[] {
                LogManager.LogCategory.General, LogManager.LogCategory.FenceCreation,
                LogManager.LogCategory.FenceUpdate, LogManager.LogCategory.UI,
                LogManager.LogCategory.IconHandling, LogManager.LogCategory.Error,
                LogManager.LogCategory.ImportExport, LogManager.LogCategory.Settings
            };

            for (int i = 0; i < 4; i++) CreateLogCategoryCheckBox(leftColumn, logCategories[i].ToString(), SettingsManager.EnabledLogCategories.Contains(logCategories[i]));
            for (int i = 4; i < 8; i++) CreateLogCategoryCheckBox(rightColumn, logCategories[i].ToString(), SettingsManager.EnabledLogCategories.Contains(logCategories[i]));

            Grid.SetColumn(leftColumn, 0); Grid.SetColumn(rightColumn, 1);
            categoryGrid.Children.Add(leftColumn); categoryGrid.Children.Add(rightColumn);
            parent.Children.Add(categoryGrid);
        }

        private static void CreateLogCategoryCheckBox(StackPanel parent, string categoryName, bool isEnabled)
        {
            parent.Children.Add(new CheckBox { Name = "LogCategory" + categoryName, Content = categoryName, IsChecked = isEnabled, FontSize = 13, Margin = new Thickness(15, 8, 0, 8) });
        }

        private static void ProcessLogCategoryCheckboxes(Grid categoryGrid, List<LogManager.LogCategory> enabledCategories)
        {
            foreach (var gridChild in categoryGrid.Children)
            {
                if (gridChild is StackPanel stackPanel)
                {
                    foreach (var stackChild in stackPanel.Children)
                    {
                        if (stackChild is CheckBox checkBox && checkBox.Name.StartsWith("LogCategory") && checkBox.IsChecked == true)
                        {
                            string categoryName = checkBox.Name.Substring("LogCategory".Length);
                            if (Enum.TryParse<LogManager.LogCategory>(categoryName, out var category))
                                if (!enabledCategories.Contains(category)) enabledCategories.Add(category);
                        }
                    }
                }
            }
        }

        // Keep existing CreateFooter, CreateDonationSection, RestoreBackup, OpenBackupsFolder, OpenLogFile 
        // (They are unchanged logic, included here for compilation completeness if copy-pasted)
        private static void CreateFooter(Grid mainGrid)
        {
            Border footerBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(20, 8, 20, 8)
            };
            Grid.SetRow(footerBorder, 2);

            StackPanel buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

            Button cancelButton = new Button { Content = "Cancel", Width = 100, Height = 34, FontWeight = FontWeights.Bold, Background = Brushes.White, BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Margin = new Thickness(0, 0, 10, 0), Cursor = Cursors.Hand };
            cancelButton.Click += (s, e) => _optionsWindow.Close();

            Button saveButton = new Button { Content = "Save", Width = 100, Height = 34, FontWeight = FontWeights.Bold, Background = new SolidColorBrush(_userAccentColor), Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
            saveButton.Click += (s, e) => SaveOptions();

            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(saveButton);
            footerBorder.Child = buttonPanel;
            mainGrid.Children.Add(footerBorder);
        }

        private static void CreateDonationSection(Grid mainGrid)
        {
            Border donationBorder = new Border { Background = new SolidColorBrush(Color.FromRgb(255, 248, 225)), BorderBrush = new SolidColorBrush(Color.FromRgb(255, 193, 7)), BorderThickness = new Thickness(0, 1, 0, 0), Padding = new Thickness(20) };
            Grid.SetRow(donationBorder, 3);
            StackPanel donationPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
            TextBlock donationText = new TextBlock { Text = "Support the Maintenance and Enhancement of This Project by Donating", FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(102, 77, 3)), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 15, 0) };
            Button paypalButton = new Button { Content = "♥ Donate via PayPal", FontSize = 14, Background = new SolidColorBrush(Color.FromRgb(255, 193, 7)), Foreground = Brushes.White, BorderThickness = new Thickness(0), Padding = new Thickness(15, 6, 15, 6), Cursor = Cursors.Hand };
            paypalButton.Click += (s, e) => { try { Process.Start(new ProcessStartInfo { FileName = "https://www.paypal.com/donate/?hosted_button_id=M8H4M4R763RBE", UseShellExecute = true }); } catch { } };
            donationPanel.Children.Add(donationText); donationPanel.Children.Add(paypalButton);
            donationBorder.Child = donationPanel;
            mainGrid.Children.Add(donationBorder);
        }

        private static void RestoreBackup()
        {
            try
            {
                string rootDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string backupsDir = System.IO.Path.Combine(rootDir, "Backups");
                using (var dialog = new System.Windows.Forms.FolderBrowserDialog { SelectedPath = System.IO.Directory.Exists(backupsDir) ? backupsDir : rootDir })
                {
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        BackupManager.RestoreFromBackup(dialog.SelectedPath);
                        _optionsWindow.Close();
                        TrayManager.reloadAllFences();
                    }
                }
            }
            catch (Exception ex) { MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Restore failed: {ex.Message}", "Error"); }
        }

        private static void OpenBackupsFolder()
        {
            try
            {
                string rootDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string backupsDir = System.IO.Path.Combine(rootDir, "Backups");
                if (!System.IO.Directory.Exists(backupsDir)) System.IO.Directory.CreateDirectory(backupsDir);
                Process.Start(new ProcessStartInfo { FileName = backupsDir, UseShellExecute = true });
            }
            catch (Exception ex) { MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error: {ex.Message}", "Error"); }
        }

        private static void OpenLogFile()
        {
            try
            {
                string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                if (System.IO.File.Exists(logPath)) Process.Start(new ProcessStartInfo { FileName = logPath, UseShellExecute = true });
                else MessageBoxesManager.ShowOKOnlyMessageBoxForm("Log file not found.", "Information");
            }
            catch { }
        }
    }
}