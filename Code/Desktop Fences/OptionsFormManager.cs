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
    /// <summary>
    /// Manages the modern WPF Options form with tabbed interface and proper DPI scaling
    /// Replaces the Windows Forms options dialog from TrayManager
    /// </summary>
    public static class OptionsFormManager
    {
        private static TabControl _tabControl;
        private static Window _optionsWindow;
        private static Color _userAccentColor;

        /// <summary>
        /// Shows the modern Options form with DPI scaling and tabbed interface
        /// </summary>
        public static void ShowOptionsForm()
        {
            try
            {
                // Get user's accent color for modern design elements
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                _userAccentColor = mediaColor;

                _optionsWindow = new Window
                {
                    Title = "Desktop Fences + Options",
                    Width = 750,
                    Height = 800,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStyle = WindowStyle.None,
                    Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                    AllowsTransparency = true
                };

                // Set icon from executable
                try
                {
                    _optionsWindow.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { } // Ignore icon loading errors

                // Main container with white background and shadow
                Border mainBorder = new Border
                {
                    Background = Brushes.White,
                    CornerRadius = new CornerRadius(0),
                    Margin = new Thickness(8),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 270,
                        ShadowDepth = 2,
                        BlurRadius = 10,
                        Opacity = 0.2
                    }
                };

                // Main grid layout
                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(40) }); // Header
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(60) }); // Footer
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(80) }); // Donation

                // Header with accent color
                Border headerBorder = new Border
                {
                    Background = new SolidColorBrush(_userAccentColor),
                    Height = 40
                };
                Grid.SetRow(headerBorder, 0);

                Grid headerGrid = new Grid();
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });

                // Header title
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

                // Close button
                Button closeButton = new Button
                {
                    Content = "✕",
                    Width = 32,
                    Height = 32,
                    FontSize = 14,
                    FontFamily = new FontFamily("Segoe UI"),
                    Foreground = Brushes.White,
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center
                };
                closeButton.Click += (s, e) => _optionsWindow.Close();

                // Close button hover effect
                closeButton.MouseEnter += (s, e) => closeButton.Background = new SolidColorBrush(Color.FromArgb(50, 255, 255, 255));
                closeButton.MouseLeave += (s, e) => closeButton.Background = Brushes.Transparent;

                Grid.SetColumn(closeButton, 1);

                headerGrid.Children.Add(titleBlock);
                headerGrid.Children.Add(closeButton);
                headerBorder.Child = headerGrid;

                // Create main content with tabs
                CreateTabContent(mainGrid);

                // Footer with buttons
                CreateFooter(mainGrid);

                // Donation section
                CreateDonationSection(mainGrid);
                // Assemble the window
                mainGrid.Children.Add(headerBorder);
                mainBorder.Child = mainGrid;
                _optionsWindow.Content = mainBorder;

                // Add keyboard support for Enter/Escape keys
                _optionsWindow.KeyDown += OptionsForm_KeyDown;
                _optionsWindow.Focusable = true;
                _optionsWindow.Focus();

                // Show the window
                _optionsWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing Options form: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"An error occurred: {ex.Message}", "Error");
            }
        }

        /// <summary>
        /// Handles keyboard input for OptionsForm
        /// Enter = Save, Escape = Cancel
        /// </summary>
        private static void OptionsForm_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                    // Trigger Save button logic
                    SaveOptions();
                    e.Handled = true;
                    break;
                case Key.Escape:
                    // Trigger Cancel button logic (close window)
                    _optionsWindow.Close();
                    e.Handled = true;
                    break;
            }
        }

        private static void CreateTabContent(Grid mainGrid)
        {
            // Main content grid with left tabs and right content
            Grid contentGrid = new Grid();
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) }); // Left tab panel
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Right content
            Grid.SetRow(contentGrid, 1);

            // Left tab panel with top margin for gap from header
            StackPanel tabPanel = new StackPanel
            {
                Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                Margin = new Thickness(0, 20, 0, 0) // Add top margin for gap
            };
            Grid.SetColumn(tabPanel, 0);

            // Right content area with top margin
            Border contentBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1, 0, 0, 0),
                Padding = new Thickness(20),
                Margin = new Thickness(0, 20, 0, 0) // Add top margin for gap
            };
            Grid.SetColumn(contentBorder, 1);

            // Create tab control for content switching
            _tabControl = new TabControl
            {
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0)
            };

            // Create custom template to hide default tab headers
            var template = new ControlTemplate(typeof(TabControl));
            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.SetValue(ContentPresenter.ContentSourceProperty, "SelectedContent");
            template.VisualTree = contentPresenter;
            _tabControl.Template = template;

            // Create tabs
            CreateGeneralTab();
            CreateToolsTab();
            CreateLookDeeperTab();

            // Set initial tab selection
            _tabControl.SelectedIndex = 0;

            // Create custom tab buttons with gaps
            CreateTabButton(tabPanel, "General", 0, true);
            CreateTabButton(tabPanel, "Tools", 1, false);
            CreateTabButton(tabPanel, "Look Deeper", 2, false);

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
                HorizontalContentAlignment = System.Windows.HorizontalAlignment.Left,
                Padding = new Thickness(20, 0, 0, 0),
                Margin = new Thickness(0, 0, 0, 2)
            };

            // Set initial colors based on selection state
            SetTabButtonColors(tabButton, title, isSelected);

            tabButton.Click += (s, e) => SelectTab(tabIndex, tabButton);

            // Hover effects - ONLY for non-selected tabs
            tabButton.MouseEnter += (s, e) =>
            {
                if (_tabControl.SelectedIndex != tabIndex) // Only if not selected
                {
                    SetTabButtonColors(tabButton, title, false, true); // true = hover state
                }
            };

            tabButton.MouseLeave += (s, e) =>
            {
                if (_tabControl.SelectedIndex != tabIndex) // Only if not selected
                {
                    SetTabButtonColors(tabButton, title, false, false); // false = normal state
                }
            };

            parent.Children.Add(tabButton);
        }

        private static void SetTabButtonColors(Button button, string title, bool isSelected, bool isHover = false)
        {
            Color buttonColor, textColor;

            if (isSelected)
            {
                // Selected state colors
                if (title == "General")
                {
                    buttonColor = _userAccentColor;
                    textColor = Colors.White;
                }
                else if (title == "Tools")
                {
                    buttonColor = Color.FromRgb(34, 139, 34); // Green
                    textColor = Colors.White;
                }
                else // Look Deeper
                {
                    buttonColor = Color.FromRgb(220, 53, 69); // Red
                    textColor = Colors.White;
                }
            }
            else if (isHover)
            {
                // Hover state colors - lighter version of selected color
                if (title == "General")
                {
                    var accent = _userAccentColor;
                    buttonColor = Color.FromRgb((byte)(accent.R + 30), (byte)(accent.G + 30), (byte)(accent.B + 30));
                    textColor = Colors.White;
                }
                else if (title == "Tools")
                {
                    buttonColor = Color.FromRgb(64, 169, 64); // Lighter green
                    textColor = Colors.White;
                }
                else // Look Deeper
                {
                    buttonColor = Color.FromRgb(250, 83, 99); // Lighter red
                    textColor = Colors.White;
                }
            }
            else
            {
                // Normal unselected state
                buttonColor = Color.FromRgb(200, 200, 200);
                textColor = Color.FromRgb(60, 60, 60);
            }

            button.Background = new SolidColorBrush(buttonColor);
            button.Foreground = new SolidColorBrush(textColor);
        }

        private static void SelectTab(int tabIndex, Button selectedButton)
        {
            _tabControl.SelectedIndex = tabIndex;

            // Update all tab button appearances
            StackPanel tabPanel = (StackPanel)selectedButton.Parent;
            for (int i = 0; i < tabPanel.Children.Count; i++)
            {
                if (tabPanel.Children[i] is Button btn)
                {
                    string buttonText = btn.Content.ToString();
                    bool isSelected = (i == tabIndex);

                    // Use the centralized color setting method
                    SetTabButtonColors(btn, buttonText, isSelected);
                    btn.FontWeight = FontWeights.Bold;
                    btn.FontSize = 14;
                }
            }
        }

        private static void CreateGeneralTab()
        {
            TabItem generalTab = new TabItem();
            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };

            StackPanel content = new StackPanel();

            // Startup Section - aligned with top
            CreateSectionHeader(content, "Startup", _userAccentColor);
            CreateCheckBox(content, "Start with Windows", "StartWithWindows", TrayManager.IsStartWithWindows);

            // Selections Section  
            CreateSectionHeader(content, "Selections", _userAccentColor);
            CreateCheckBox(content, "Single Click to Launch", "SingleClickToLaunch", SettingsManager.SingleClickToLaunch);
            CreateCheckBox(content, "Enable Snap Near Fences", "EnableSnapNearFences", SettingsManager.IsSnapEnabled);
            CreateCheckBox(content, "Enable Dimension Snap", "EnableDimensionSnap", SettingsManager.EnableDimensionSnap);
            CreateCheckBox(content, "Enable Tray Icon", "EnableTrayIcon", SettingsManager.ShowInTray);
            CreateCheckBox(content, "Enable Portal Fences Watermark", "EnablePortalWatermark", SettingsManager.ShowBackgroundImageOnPortalFences);
            CreateCheckBox(content, "Use Recycle Bin on Portal Fences 'Delete item' command", "UseRecycleBin", SettingsManager.UseRecycleBin);
            CreateCheckBox(content, "Disable Fence Scrollbars", "DisableFenceScrollbars", SettingsManager.DisableFenceScrollbars);
            // Style Section
            CreateSectionHeader(content, "Style", _userAccentColor);
            CreateTintSlider(content);
            CreateColorComboBox(content);
            CreateLaunchEffectComboBox(content);

            scrollViewer.Content = content;
            generalTab.Content = scrollViewer;
            _tabControl.Items.Add(generalTab);
        }

        private static void CreateToolsTab()
        {
            TabItem toolsTab = new TabItem();
            StackPanel content = new StackPanel();

            // Tools Section - aligned with top
            CreateSectionHeader(content, "Tools", Color.FromRgb(34, 139, 34));

            // Button panel with proper alignment
            Grid buttonGrid = new Grid();
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(15) }); // Gap
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            buttonGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) });
            buttonGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(15) }); // Gap
            buttonGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) });
            buttonGrid.Margin = new Thickness(0, 10, 0, 0);

            // Backup button
            Button backupButton = new Button
            {
                Content = "Backup",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(Color.FromRgb(34, 139, 34)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            backupButton.Click += (s, e) => BackupManager.BackupData();
            Grid.SetRow(backupButton, 0);
            Grid.SetColumn(backupButton, 0);

            // Restore button
            Button restoreButton = new Button
            {
                Content = "Restore...",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(Color.FromRgb(255, 152, 0)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            restoreButton.Click += (s, e) => RestoreBackup();
            Grid.SetRow(restoreButton, 0);
            Grid.SetColumn(restoreButton, 2);

            // Open Backups Folder button
            Button openBackupsButton = new Button
            {
                Content = "Open Backups Folder",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(Color.FromRgb(0, 123, 191)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            openBackupsButton.Click += (s, e) => OpenBackupsFolder();
            Grid.SetRow(openBackupsButton, 2);
            Grid.SetColumn(openBackupsButton, 0);
            Grid.SetColumnSpan(openBackupsButton, 3);

            buttonGrid.Children.Add(backupButton);
            buttonGrid.Children.Add(restoreButton);
            buttonGrid.Children.Add(openBackupsButton);
            content.Children.Add(buttonGrid);

            toolsTab.Content = content;
            _tabControl.Items.Add(toolsTab);
        }

        private static void CreateLookDeeperTab()
        {
            TabItem lookDeeperTab = new TabItem();
            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };

            StackPanel content = new StackPanel();

            // Log Section - aligned with top
            CreateSectionHeader(content, "Log", Color.FromRgb(220, 53, 69));
            CreateCheckBox(content, "Enable logging", "EnableLogging", SettingsManager.IsLogEnabled);

            // Open Log button
            Button openLogButton = new Button
            {
                Content = "Open Log",
                Width = 100,
                Height = 25,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(Color.FromRgb(220, 53, 69)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Margin = new Thickness(0, 10, 0, 20),
                Cursor = Cursors.Hand,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            };
            openLogButton.Click += (s, e) => OpenLogFile();
            content.Children.Add(openLogButton);

            // Log configuration section
            CreateSectionHeader(content, "Log configuration", Color.FromRgb(220, 53, 69));
            CreateLogLevelComboBox(content);

            // Log Categories section
            CreateSectionHeader(content, "Log Categories", Color.FromRgb(220, 53, 69));
            CreateLogCategoryCheckBoxes(content);

            CreateCheckBox(content, "Enable Background Validation Logging", "EnableBackgroundValidation", SettingsManager.EnableBackgroundValidationLogging);

            scrollViewer.Content = content;
            lookDeeperTab.Content = scrollViewer;
            _tabControl.Items.Add(lookDeeperTab);
        }

        private static void CreateSectionHeader(StackPanel parent, string title, Color headerColor)
        {
            TextBlock header = new TextBlock
            {
                Text = title,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(headerColor),
                Margin = new Thickness(0, 0, 0, 15)
            };
            parent.Children.Add(header);
        }

        private static void CreateCheckBox(StackPanel parent, string text, string name, bool isChecked)
        {
            CheckBox checkBox = new CheckBox
            {
                Name = name,
                Content = text,
                IsChecked = isChecked,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Margin = new Thickness(15, 8, 0, 8),
                VerticalAlignment = VerticalAlignment.Center
            };
            parent.Children.Add(checkBox);
        }

        private static void CreateTintSlider(StackPanel parent)
        {
            Grid sliderGrid = new Grid();
            sliderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            sliderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            sliderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(50) });
            sliderGrid.Margin = new Thickness(0, 10, 0, 10);

            // Label
            TextBlock label = new TextBlock
            {
                Text = "Tint",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 10, 0)
            };
            Grid.SetColumn(label, 0);

            // Slider
            Slider slider = new Slider
            {
                Name = "TintSlider",
                Minimum = 1,
                Maximum = 100,
                Value = SettingsManager.TintValue,
                TickFrequency = 1,
                IsSnapToTickEnabled = true,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(slider, 1);

            // Value display
            TextBlock valueDisplay = new TextBlock
            {
                Name = "TintValue",
                Text = SettingsManager.TintValue.ToString(),
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(15, 0, 0, 0)
            };
            Grid.SetColumn(valueDisplay, 2);

            slider.ValueChanged += (s, e) => valueDisplay.Text = ((int)e.NewValue).ToString();

            sliderGrid.Children.Add(label);
            sliderGrid.Children.Add(slider);
            sliderGrid.Children.Add(valueDisplay);
            parent.Children.Add(sliderGrid);
        }

        private static void CreateColorComboBox(StackPanel parent)
        {
            Grid colorGrid = new Grid();
            colorGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            colorGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            colorGrid.Margin = new Thickness(0, 10, 0, 10);

            TextBlock label = new TextBlock
            {
                Text = "Color",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 10, 0)
            };
            Grid.SetColumn(label, 0);

            ComboBox comboBox = new ComboBox
            {
                Name = "ColorComboBox",
                Height = 25,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center
            };

            string[] colors = { "Gray", "Black", "White", "Beige", "Green", "Purple", "Fuchsia", "Yellow", "Orange", "Red", "Blue", "Bismark" };
            foreach (string color in colors)
            {
                comboBox.Items.Add(color);
            }
            comboBox.SelectedItem = SettingsManager.SelectedColor;
            Grid.SetColumn(comboBox, 1);

            colorGrid.Children.Add(label);
            colorGrid.Children.Add(comboBox);
            parent.Children.Add(colorGrid);
        }

        private static void CreateLaunchEffectComboBox(StackPanel parent)
        {
            Grid effectGrid = new Grid();
            effectGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            effectGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            effectGrid.Margin = new Thickness(0, 10, 0, 10);

            TextBlock label = new TextBlock
            {
                Text = "Effect",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 10, 0)
            };
            Grid.SetColumn(label, 0);

            ComboBox comboBox = new ComboBox
            {
                Name = "LaunchEffectComboBox",
                Height = 25,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center
            };

            string[] effects = { "Zoom", "Bounce", "FadeOut", "SlideUp", "Rotate", "Agitate", "GrowAndFly", "Pulse", "Elastic", "Flip3D", "Spiral", "Shockwave", "Matrix", "Supernova", "Teleport" };
            foreach (string effect in effects)
            {
                comboBox.Items.Add(effect);
            }
            comboBox.SelectedIndex = (int)SettingsManager.LaunchEffect;
            Grid.SetColumn(comboBox, 1);

            effectGrid.Children.Add(label);
            effectGrid.Children.Add(comboBox);
            parent.Children.Add(effectGrid);
        }

        private static void CreateLogLevelComboBox(StackPanel parent)
        {
            Grid levelGrid = new Grid();
            levelGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            levelGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            levelGrid.Margin = new Thickness(0, 10, 0, 10);

            TextBlock label = new TextBlock
            {
                Text = "Minimum Log Level",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            };
            Grid.SetColumn(label, 0);

            ComboBox comboBox = new ComboBox
            {
                Name = "LogLevelComboBox",
                Height = 25,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center
            };

            comboBox.Items.Add("Debug");
            comboBox.Items.Add("Info");
            comboBox.Items.Add("Warn");
            comboBox.Items.Add("Error");
            comboBox.SelectedItem = SettingsManager.MinLogLevel.ToString();
            Grid.SetColumn(comboBox, 1);

            levelGrid.Children.Add(label);
            levelGrid.Children.Add(comboBox);
            parent.Children.Add(levelGrid);
        }

        private static void CreateLogCategoryCheckBoxes(StackPanel parent)
        {
            Grid categoryGrid = new Grid();
            categoryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            categoryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            StackPanel leftColumn = new StackPanel();
            StackPanel rightColumn = new StackPanel();

            var logCategories = new[] {
                LogManager.LogCategory.General,
                LogManager.LogCategory.FenceCreation,
                LogManager.LogCategory.FenceUpdate,
                LogManager.LogCategory.UI,
                LogManager.LogCategory.IconHandling,
                LogManager.LogCategory.Error,
                LogManager.LogCategory.ImportExport,
                LogManager.LogCategory.Settings
            };

            for (int i = 0; i < 4; i++)
            {
                CreateLogCategoryCheckBox(leftColumn, logCategories[i].ToString(), SettingsManager.EnabledLogCategories.Contains(logCategories[i]));
            }

            for (int i = 4; i < 8; i++)
            {
                CreateLogCategoryCheckBox(rightColumn, logCategories[i].ToString(), SettingsManager.EnabledLogCategories.Contains(logCategories[i]));
            }

            Grid.SetColumn(leftColumn, 0);
            Grid.SetColumn(rightColumn, 1);
            categoryGrid.Children.Add(leftColumn);
            categoryGrid.Children.Add(rightColumn);

            parent.Children.Add(categoryGrid);
        }

        private static void CreateLogCategoryCheckBox(StackPanel parent, string categoryName, bool isEnabled)
        {
            CheckBox checkBox = new CheckBox
            {
                Name = "LogCategory" + categoryName,
                Content = categoryName,
                IsChecked = isEnabled,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Margin = new Thickness(15, 8, 0, 8)
            };
            parent.Children.Add(checkBox);
        }

        private static void RestoreBackup()
        {
            try
            {
                string rootDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string backupsDir = System.IO.Path.Combine(rootDir, "Backups");
                string initialPath = System.IO.Directory.Exists(backupsDir) ? backupsDir : rootDir;

                using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    dialog.Description = "Select backup folder to restore from";
                    dialog.ShowNewFolderButton = false;
                    dialog.SelectedPath = initialPath;

                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        BackupManager.RestoreFromBackup(dialog.SelectedPath);
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
            try
            {
                string rootDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string backupsDir = System.IO.Path.Combine(rootDir, "Backups");
                if (!System.IO.Directory.Exists(backupsDir))
                {
                    System.IO.Directory.CreateDirectory(backupsDir);
                }
                Process.Start(new ProcessStartInfo
                {
                    FileName = backupsDir,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
              MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error opening backups folder: {ex.Message}", "Error");
            }
        }

        private static void OpenLogFile()
        {
            try
            {
                string exeDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string logPath = System.IO.Path.Combine(exeDir, "Desktop_Fences.log");
                if (System.IO.File.Exists(logPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = logPath,
                        UseShellExecute = true
                    });
                }
                else
                {
                  MessageBoxesManager.ShowOKOnlyMessageBoxForm("Log file not found.", "Information");
                }
            }
            catch (Exception ex)
            {
              MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error opening log file: {ex.Message}", "Error");
            }
        }

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

            StackPanel buttonPanel = new StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            };

            System.Windows.Controls.Button cancelButton = new System.Windows.Controls.Button
            {
                Content = "Cancel",
                Width = 100,
                Height = 34,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            cancelButton.Click += (s, e) => _optionsWindow.Close();

            System.Windows.Controls.Button saveButton = new System.Windows.Controls.Button
            {
                Content = "Save",
                Width = 100,
                Height = 34,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(_userAccentColor),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            saveButton.Click += (s, e) => SaveOptions();

            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(saveButton);
            footerBorder.Child = buttonPanel;

            mainGrid.Children.Add(footerBorder);
        }

        private static void CreateDonationSection(Grid mainGrid)
        {
            Border donationBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(255, 248, 225)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(255, 193, 7)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(20)
            };
            Grid.SetRow(donationBorder, 3);

            StackPanel donationPanel = new StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            TextBlock donationText = new TextBlock
            {
                Text = "Support the Maintenance and Enhancement of This Project by Donating",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(102, 77, 3)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 15, 0)
            };

            Button paypalButton = new Button
            {
                Content = "♥ Donate via PayPal",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                Background = new SolidColorBrush(Color.FromRgb(255, 193, 7)),
               
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(15, 6, 15, 6),
                Cursor = Cursors.Hand
            };
            paypalButton.Click += (s, e) =>
            {



                        try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "https://www.paypal.com/donate/?hosted_button_id=M8H4M4R763RBE",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error opening PayPal donation link: {ex.Message}");
                }
            };

            donationPanel.Children.Add(donationText);
            donationPanel.Children.Add(paypalButton);
            donationBorder.Child = donationPanel;

            mainGrid.Children.Add(donationBorder);
        }

        // ADD this helper method to OptionsFormManager class:
        private static void ProcessLogCategoryCheckboxes(Grid categoryGrid, List<LogManager.LogCategory> enabledCategories)
        {
            // The log category checkboxes are in nested StackPanels within the Grid
            foreach (var gridChild in categoryGrid.Children)
            {
                if (gridChild is StackPanel stackPanel)
                {
                    foreach (var stackChild in stackPanel.Children)
                    {
                        if (stackChild is CheckBox checkBox && checkBox.Name.StartsWith("LogCategory"))
                        {
                            if (checkBox.IsChecked == true)
                            {
                                string categoryName = checkBox.Name.Substring("LogCategory".Length);
                                if (Enum.TryParse<LogManager.LogCategory>(categoryName, out var category))
                                {
                                    if (!enabledCategories.Contains(category))
                                    {
                                        enabledCategories.Add(category);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


        private static void SaveOptions()
        {
            try
            {
                bool tempPortalImageState = SettingsManager.ShowBackgroundImageOnPortalFences;

                TabItem generalTab = (TabItem)_tabControl.Items[0];
                ScrollViewer generalScrollViewer = (ScrollViewer)generalTab.Content;
                StackPanel generalContent = (StackPanel)generalScrollViewer.Content;

                bool newPortalWatermarkState = false;
                bool newShowInTrayState = false;

                foreach (var child in generalContent.Children)
                {
                    if (child is CheckBox checkBox)
                    {
                        switch (checkBox.Name)
                        {
                            case "StartWithWindows":
                                if (checkBox.IsChecked != TrayManager.IsStartWithWindows)
                                {
                                    TrayManager.Instance?.ToggleStartWithWindows(checkBox.IsChecked == true);
                                }
                                break;
                            case "SingleClickToLaunch":
                                SettingsManager.SingleClickToLaunch = checkBox.IsChecked == true;
                                break;
                            case "EnableSnapNearFences":
                                SettingsManager.IsSnapEnabled = checkBox.IsChecked == true;
                                break;
                            case "EnableDimensionSnap":
                                SettingsManager.EnableDimensionSnap = checkBox.IsChecked == true;
                                break;
                            case "EnableTrayIcon":
                                newShowInTrayState = checkBox.IsChecked == true;
                                SettingsManager.ShowInTray = newShowInTrayState;
                                if (TrayManager.Instance != null)
                                {
                                    TrayManager.Instance.GetType().GetField("Showintray", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.SetValue(TrayManager.Instance, newShowInTrayState);
                                }
                                break;
                            case "EnablePortalWatermark":
                                newPortalWatermarkState = checkBox.IsChecked == true;
                                SettingsManager.ShowBackgroundImageOnPortalFences = newPortalWatermarkState;
                                break;
                            case "UseRecycleBin":
                                SettingsManager.UseRecycleBin = checkBox.IsChecked == true;
                                break;
                            case "DisableFenceScrollbars":
                                SettingsManager.DisableFenceScrollbars = checkBox.IsChecked == true;
                                break;
                        }
                    }
                    else if (child is Grid grid)
                    {
                        var slider = grid.Children.OfType<Slider>().FirstOrDefault(s => s.Name == "TintSlider");
                        if (slider != null)
                        {
                            SettingsManager.TintValue = (int)slider.Value;
                        }

                        var colorCombo = grid.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "ColorComboBox");
                        if (colorCombo?.SelectedItem != null)
                        {
                            SettingsManager.SelectedColor = colorCombo.SelectedItem.ToString();
                        }

                        var effectCombo = grid.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "LaunchEffectComboBox");
                        if (effectCombo != null)
                        {
                            SettingsManager.LaunchEffect = (LaunchEffectsManager.LaunchEffect)effectCombo.SelectedIndex;
                        }
                    }
                }

                TabItem lookDeeperTab = (TabItem)_tabControl.Items[2];
                ScrollViewer lookDeeperScrollViewer = (ScrollViewer)lookDeeperTab.Content;
                StackPanel lookDeeperContent = (StackPanel)lookDeeperScrollViewer.Content;

                var newEnabledCategories = new List<LogManager.LogCategory>();

                foreach (var child in lookDeeperContent.Children)
                {
                    if (child is CheckBox checkBox)
                    {
                        switch (checkBox.Name)
                        {
                            case "EnableLogging":
                                SettingsManager.IsLogEnabled = checkBox.IsChecked == true;
                                break;
                            case "EnableBackgroundValidation":
                                SettingsManager.EnableBackgroundValidationLogging = checkBox.IsChecked == true;
                                break;
                        }
                    }
                    else if (child is Grid grid)
                    {
                        // Handle log level combo box
                        var logLevelCombo = grid.Children.OfType<ComboBox>().FirstOrDefault(c => c.Name == "LogLevelComboBox");
                        if (logLevelCombo?.SelectedItem != null)
                        {
                            if (Enum.TryParse<LogManager.LogLevel>(logLevelCombo.SelectedItem.ToString(), out var logLevel))
                            {
                                SettingsManager.SetMinLogLevel(logLevel);
                            }
                        }

                        // Handle log category checkboxes (they're in nested StackPanels within the Grid)
                        ProcessLogCategoryCheckboxes(grid, newEnabledCategories);
                    }
                }


                // Update categories once after collecting all checked ones
                SettingsManager.EnabledLogCategories = newEnabledCategories;

                if (tempPortalImageState != newPortalWatermarkState)
                {
                    TrayManager.reloadAllFences();
                }

                TrayManager.Instance?.UpdateTrayIcon();
                var fences = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
                foreach (var fence in fences)
                {
                    dynamic fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == fence.Title);
                    if (fenceData != null)
                    {
                        string customColor = fenceData.CustomColor?.ToString();
                        string appliedColor = string.IsNullOrEmpty(customColor) ? SettingsManager.SelectedColor : customColor;
                        Utility.ApplyTintAndColorToFence(fence, appliedColor);
                    }
                    else
                    {
                        Utility.ApplyTintAndColorToFence(fence, SettingsManager.SelectedColor);
                    }

                    // Apply scrollbar setting to existing fences
                    var border = fence.Content as Border;
                    if (border?.Child is DockPanel dockPanel)
                    {
                        var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                        if (scrollViewer != null)
                        {
                            scrollViewer.VerticalScrollBarVisibility = SettingsManager.DisableFenceScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto;
                        }
                    }
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.Settings, "Options saved successfully");
                _optionsWindow.Close();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings, $"Error saving options: {ex.Message}");
              MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error saving settings: {ex.Message}", "Save Error");
            }


        }
    }
}