using Desktop_Frames.Localization;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Diagnostics;

namespace Desktop_Frames
{
    /// <summary>
    /// Modern WPF form for customizing frame-specific properties
    /// Converted from Windows Forms CustomizeFrameForm with identical functionality
    /// DPI-aware with monitor positioning support
    /// </summary>
    public class CustomizeFrameFormManager : Window
    {
        #region Private Fields
        private dynamic _frame;
        private bool _result = false;

        // Controls for the 12 customization settings
        private ComboBox _cmbCustomColor;
        private ComboBox _cmbCustomLaunchEffect;
        private ComboBox _cmbframeBorderColor;
        private NumericTextBox _nudframeBorderThickness;
        private ComboBox _cmbTitleTextColor;
        private ComboBox _cmbTitleTextSize;
        private CheckBox _chkBoldTitleText;
        private ComboBox _cmbPortalView; // --- NEW: Portal View Mode Selector ---
        private ComboBox _cmbIconSize;
        private NumericTextBox _nudIconSpacing;
        private ComboBox _cmbTextColor;
        private CheckBox _chkDisableTextShadow;
        private CheckBox _chkGrayscaleIcons;

        // --- NEW: Global Action Tracking ---
        private Button _btnApply;
        private Button _btnSave;
        private bool _isCtrlPressed = false;

        // Valid options from existing code
        private readonly string[] _validColors = { "Red", "Green", "Teal", "Blue", "Bismark", "White", "Beige", "Gray", "Black", "Purple", "Fuchsia", "Yellow", "Orange" };
        private readonly string[] _validEffects = { "Zoom", "Bounce", "FadeOut", "SlideUp", "Rotate", "Agitate", "GrowAndFly", "Pulse", "Elastic", "Flip3D", "Spiral", "Shockwave", "Matrix", "Supernova", "Teleport" };
        private readonly string[] _validTextSizes = { "Small", "Medium", "Large" };
        private readonly string[] _validIconSizes = { "Tiny", "Small", "Medium", "Large", "Huge" };

        private Color _userAccentColor;
		#endregion

		#region Constructor
		/// <summary>
		/// Initialize the Customize Frame WPF form with frame-specific data
		/// </summary>
		/// <param name="frame">The frame object to customize</param>
		public CustomizeFrameFormManager(dynamic frame)
        {
			// Get the most current frame data from Framemanager to avoid stale references
			_frame = GetCurrentFrameData(frame);
            InitializeComponent();
            LoadCurrentValues();
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets whether the user clicked Save (true) or Cancel (false)
        /// </summary>
        public new bool DialogResult => _result;
        #endregion

        #region Form Initialization
        private void InitializeComponent()
        {
            try
            {
                // Get user's accent color for modern design elements
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                _userAccentColor = mediaColor;

                // Modern WPF window setup with DPI awareness
                this.Title = Strings.CustomizeTitle;
                this.Width = 500;
                this.Height = 675;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.ResizeMode = ResizeMode.NoResize;

                // Set icon from executable
                try
                {
                    this.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { } // Ignore icon loading errors

                // Main container with modern card design
                Border mainCard = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                    BorderThickness = new Thickness(1),
                    Margin = new Thickness(8),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 270,
                        ShadowDepth = 2,
                        BlurRadius = 10,
                        Opacity = 0.1
                    }
                };

                // Main grid layout
                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) }); // Header
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) }); // Footer

                CreateHeader(mainGrid);
                CreateContent(mainGrid);
                CreateFooter(mainGrid);

                mainCard.Child = mainGrid;
                this.Content = mainCard;

                // Position form on the screen where mouse is currently located
                PositionFormOnMouseScreen();

                // Add keyboard support for Enter/Escape keys
                this.KeyDown += CustomizeFrameForm_KeyDown;

                // --- NEW: Global Action CTRL Hooks ---
                this.PreviewKeyDown += Window_PreviewKeyDown;
                this.PreviewKeyUp += Window_PreviewKeyUp;
                this.Deactivated += Window_Deactivated; // Safety reset if window loses focus

                this.Focusable = true;
                this.Focus();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Initialized CustomizeFrameFormWPF for frame '{_frame.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error initializing CustomizeFrameFormWPF: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error initializing form: {ex.Message}", "Form Error");
            }
        }

        #region Global Action Input Handlers (CTRL Key)
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
            {
                UpdateButtonsState(true);
            }
        }

        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
            {
                UpdateButtonsState(false);
            }
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            UpdateButtonsState(false);
        }

        private void UpdateButtonsState(bool ctrlPressed)
        {
            if (_btnApply == null || _btnSave == null) return;
            if (_isCtrlPressed == ctrlPressed) return;

            _isCtrlPressed = ctrlPressed;

            if (_isCtrlPressed)
            {
                _btnApply.Content = Strings.BtnApplyToAll;
                _btnApply.Background = new SolidColorBrush(Color.FromRgb(204, 85, 0)); // Warning Orange

                _btnSave.Content = Strings.BtnSaveToAll;
                _btnSave.Background = new SolidColorBrush(Color.FromRgb(204, 0, 0)); // Warning Red
            }
            else
            {
                _btnApply.Content = Strings.BtnApply;
                _btnApply.Background = new SolidColorBrush(Color.FromRgb(34, 139, 34)); // Standard Green

                _btnSave.Content = Strings.BtnSave;
                _btnSave.Background = new SolidColorBrush(_userAccentColor); // Standard Accent
            }
        }
        #endregion

        /// <summary>
        /// Handles keyboard input for CustomizeFrameForm
        /// Enter = Save, Escape = Cancel
        /// </summary>
        private void CustomizeFrameForm_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                    // Trigger Save button logic
                    SaveButton_Click(this, new RoutedEventArgs());
                    e.Handled = true;
                    break;
                case Key.Escape:
                    // Trigger Cancel button logic
                    CancelButton_Click(this, new RoutedEventArgs());
                    e.Handled = true;
                    break;
            }
        }

        private void CreateHeader(Grid parent)
        {
            // Header border with accent color background
            Border headerBorder = new Border
            {
                Background = new SolidColorBrush(_userAccentColor),
                Height = 50
            };
            Grid.SetRow(headerBorder, 0);

            // Header grid for layout
            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });

            // Title label
            TextBlock titleBlock = new TextBlock
            {
                Text = Strings.CustomizeTitle,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0)
            };
            Grid.SetColumn(titleBlock, 0);

            // Close button (✕)
            Button closeButton = new Button
            {
                Content = "✕",
                Width = 32,
                Height = 32,
                FontSize = 16,
                FontFamily = new FontFamily("Segoe UI"),
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            closeButton.Click += CloseButton_Click;
            closeButton.MouseEnter += (s, e) => closeButton.Background = new SolidColorBrush(Color.FromArgb(50, 255, 255, 255));
            closeButton.MouseLeave += (s, e) => closeButton.Background = Brushes.Transparent;
            Grid.SetColumn(closeButton, 1);

            headerGrid.Children.Add(titleBlock);
            headerGrid.Children.Add(closeButton);
            headerBorder.Child = headerGrid;
            // Add drag functionality to header area (like EditShortcutWindow and OptionsForm)
            headerBorder.MouseLeftButtonDown += (sender, e) =>
            {
                if (e.ButtonState == MouseButtonState.Pressed)
                {
                    this.DragMove();
                }
            };



            parent.Children.Add(headerBorder);
        }

        private void CreateContent(Grid parent)
        {
            // Content scroll viewer
            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Background = Brushes.White,
                Padding = new Thickness(16)
            };
            Grid.SetRow(scrollViewer, 1);

            // Main content stack panel
            StackPanel contentStack = new StackPanel
            {
                Orientation = Orientation.Vertical
            };

            CreateFrameSection(contentStack);
            CreateTitleSection(contentStack);
            CreateIconsSection(contentStack);

            scrollViewer.Content = contentStack;
            parent.Children.Add(scrollViewer);
        }



        private void CreateFooter(Grid parent)
        {
            // Footer border
            Border footerBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(16, 8, 16, 8)
            };
            Grid.SetRow(footerBorder, 2);

            // Button panel
            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            // 1. Default Button (Moved Left & Restyled to Dark Gray)
            Button defaultButton = new Button
            {
                Content = Strings.BtnDefault,
                Width = 80, // Slightly smaller to fit 4 buttons
                Height = 34,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(Color.FromRgb(80, 80, 80)), // Dark Gray
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 10, 0)
            };
            defaultButton.Click += DefaultButton_Click;

            // 2. Apply Button (Now tracking class-level for Global Override)
            _btnApply = new Button
            {
                Content = Strings.BtnApply,
                Width = 100,
                Height = 34,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(Color.FromRgb(34, 139, 34)), // Green
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 10, 0)
            };
            _btnApply.Click += ApplyButton_Click;

            // 3. Cancel Button
            Button cancelButton = new Button
            {
                Content = Strings.BtnCancel,
                Width = 100,
                Height = 33,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 10, 0)
            };
            cancelButton.Click += CancelButton_Click;

            // 4. Save Button (Now tracking class-level for Global Override)
            _btnSave = new Button
            {
                Content = Strings.BtnSave,
                Width = 100,
                Height = 34,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(_userAccentColor),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            _btnSave.Click += SaveButton_Click;

            // Add buttons in the new order: Default -> Apply -> Cancel -> Save
            buttonPanel.Children.Add(defaultButton);
            buttonPanel.Children.Add(_btnApply);
            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(_btnSave);

            footerBorder.Child = buttonPanel;
            parent.Children.Add(footerBorder);
        }

        private void CreateFrameSection(StackPanel parent)
        {
            GroupBox frameGroupBox = new GroupBox
            {
                Header = Strings.TabFrame,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(_userAccentColor),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(8)
            };

            StackPanel frameStack = new StackPanel { Orientation = Orientation.Vertical };

            CreateDropdownField(frameStack, Strings.LblCustomColor, _validColors, out _cmbCustomColor);
            CreateDropdownField(frameStack, Strings.LblCustomLaunchEffect, _validEffects, out _cmbCustomLaunchEffect);
            CreateDropdownField(frameStack, Strings.LblFrameBorderColor, _validColors, out _cmbframeBorderColor);
            CreateNumericField(frameStack, Strings.LblFrameBorderThickness, 0, 5, out _nudframeBorderThickness);

			frameGroupBox.Content = frameStack;
            parent.Children.Add(frameGroupBox);
        }

        private void CreateTitleSection(StackPanel parent)
        {
            GroupBox titleGroupBox = new GroupBox
            {
                Header = Strings.TabTitle,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(_userAccentColor),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(8)
            };

            StackPanel titleStack = new StackPanel { Orientation = Orientation.Vertical };

            CreateDropdownField(titleStack, Strings.LblTitleTextColor, _validColors, out _cmbTitleTextColor);
            CreateDropdownField(titleStack, Strings.LblTitleTextSize, _validTextSizes, out _cmbTitleTextSize);
            CreateCheckboxField(titleStack, Strings.LblBoldTitleText, out _chkBoldTitleText);

            titleGroupBox.Content = titleStack;
            parent.Children.Add(titleGroupBox);
        }

        private void CreateIconsSection(StackPanel parent)
        {
            GroupBox iconsGroupBox = new GroupBox
            {
                Header = Strings.TabIcons,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(_userAccentColor),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(8)
            };

            StackPanel iconsStack = new StackPanel { Orientation = Orientation.Vertical };

            // --- NEW: Portal View Dropdown (Visible strictly for Portal Frames) ---
            CreateDropdownField(iconsStack, Strings.LblPortalView, new[] { "Icons", "Details" }, out _cmbPortalView);
            bool isPortal = _frame.ItemsType?.ToString() == "Portal";
            if (_cmbPortalView.Parent is FrameworkElement pvRow)
            {
                pvRow.Visibility = isPortal ? Visibility.Visible : Visibility.Collapsed;
            }
            // ----------------------------------------------------------------------

            CreateDropdownField(iconsStack, Strings.LblIconSize, _validIconSizes, out _cmbIconSize);
            CreateNumericField(iconsStack, Strings.LblIconSpacing, 0, 20, out _nudIconSpacing);
            CreateDropdownField(iconsStack, Strings.LblTextColor, _validColors, out _cmbTextColor);
            CreateCheckboxField(iconsStack, Strings.LblDisableTextShadow, out _chkDisableTextShadow);
            CreateCheckboxField(iconsStack, Strings.LblGrayscaleIcons, out _chkGrayscaleIcons);

            iconsGroupBox.Content = iconsStack;
            parent.Children.Add(iconsGroupBox);
        }

        #region Helper Methods for Control Creation
        private void CreateDropdownField(StackPanel parent, string labelText, string[] items, out ComboBox comboBox)
        {
            Grid fieldGrid = new Grid
            {
                Margin = new Thickness(0, 5, 0, 5)
            };
            fieldGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
            fieldGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock label = new TextBlock
            {
                Text = labelText,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0)
            };
            Grid.SetColumn(label, 0);

            comboBox = new ComboBox
            {
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Width = 180,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Center
            };

            // Content is read by the user, Tag is written to frames.json.
            comboBox.Items.Add(new ComboBoxItem { Content = Strings.Item("Default"), Tag = "Default" });
            foreach (var item in items)
            {
                comboBox.Items.Add(new ComboBoxItem { Content = Strings.Item(item), Tag = item });
            }
            comboBox.SelectedIndex = 0;

            Grid.SetColumn(comboBox, 1);

            fieldGrid.Children.Add(label);
            fieldGrid.Children.Add(comboBox);
            parent.Children.Add(fieldGrid);
        }

        private void CreateNumericField(StackPanel parent, string labelText, int min, int max, out NumericTextBox numericTextBox)
        {
            Grid fieldGrid = new Grid
            {
                Margin = new Thickness(0, 5, 0, 5)
            };
            fieldGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
            fieldGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock label = new TextBlock
            {
                Text = labelText,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0)
            };
            Grid.SetColumn(label, 0);

            numericTextBox = new NumericTextBox
            {
                Minimum = min,
                Maximum = max,
                Value = min,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Width = 80,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Center
            };

            Grid.SetColumn(numericTextBox, 1);

            fieldGrid.Children.Add(label);
            fieldGrid.Children.Add(numericTextBox);
            parent.Children.Add(fieldGrid);
        }

        private void CreateCheckboxField(StackPanel parent, string labelText, out CheckBox checkBox)
        {
            checkBox = new CheckBox
            {
                Content = labelText,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                Margin = new Thickness(16, 5, 0, 5),
                VerticalAlignment = VerticalAlignment.Center
            };

            parent.Children.Add(checkBox);
        }
        #endregion

        #region Event Handlers
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            _result = false;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _result = false;
            this.Close();
        }

        private void DefaultButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Default button clicked for frame '{_frame.Title}' - resetting all controls to defaults");

                // Reset all dropdown controls to "Default" (index 0)
                _cmbCustomColor.SelectedIndex = 0;
                _cmbCustomLaunchEffect.SelectedIndex = 0;
                _cmbframeBorderColor.SelectedIndex = 0;
                _cmbTitleTextColor.SelectedIndex = 0;
                _cmbTitleTextSize.SelectedIndex = 0;
                _cmbIconSize.SelectedIndex = 0;
                _cmbTextColor.SelectedIndex = 0;

                // Reset numeric controls to specified defaults
                _nudframeBorderThickness.Value = 0;
                _nudIconSpacing.Value = 5;

                // Reset checkbox controls to unchecked
                _chkBoldTitleText.IsChecked = false;
                _chkDisableTextShadow.IsChecked = false;
                _chkGrayscaleIcons.IsChecked = false;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"All controls reset to default values for frame '{_frame.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error resetting controls to defaults: {ex.Message}");
              MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error resetting to defaults: {ex.Message}", "Default Error");
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isCtrlPressed)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Save to all clicked");
                    ApplyChangesToAll();
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Save button clicked for frame '{_frame.Title}'");
                    ApplyChanges(); // Call shared logic
                }
                this.Close();   // Close the form
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error saving settings: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error saving settings: {ex.Message}", "Save Error");
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isCtrlPressed)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Apply to all clicked");
                    ApplyChangesToAll();
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Global changes applied successfully without closing form");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Apply button clicked for frame '{_frame.Title}'");
                    ApplyChanges(); // Call shared logic (Does NOT close the form)
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Changes applied successfully without closing form");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error applying settings: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error applying settings: {ex.Message}", "Apply Error");
            }
        }

        /// <summary>
        /// Gets the DPI scale factor for proper WPF window positioning
        /// </summary>
        private double GetFormDpiScaleFactor()
        {
            try
            {
                // Use Graphics to get the screen's DPI
                using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                {
                    float dpiX = graphics.DpiX; // Horizontal DPI
                    return dpiX / 96.0; // Standard DPI is 96, so scale factor = dpiX / 96
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Could not get DPI scale factor: {ex.Message}. Using default scale of 1.0");
                return 1.0; // Default to no scaling if DPI detection fails
            }
        }


                #endregion




        #region Multi-Monitor Positioning
  


        private void PositionFormOnMouseScreen()
        {
            try
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Mouse position: X={mousePosition.X}, Y={mousePosition.Y}");

                var mouseScreen = System.Windows.Forms.Screen.FromPoint(mousePosition);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Mouse is on screen: {mouseScreen.DeviceName}, Bounds: {mouseScreen.Bounds}");

                // Get DPI scale factor for proper WPF positioning
                double dpiScale = GetFormDpiScaleFactor();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"DPI scale factor: {dpiScale}");

                // Convert physical pixels to device-independent units (DIUs)
                double screenLeftDiu = mouseScreen.Bounds.Left / dpiScale;
                double screenTopDiu = mouseScreen.Bounds.Top / dpiScale;
                double screenWidthDiu = mouseScreen.Bounds.Width / dpiScale;
                double screenHeightDiu = mouseScreen.Bounds.Height / dpiScale;

                // Calculate center position in DIUs
                double centerX = screenLeftDiu + (screenWidthDiu - this.Width) / 2;
                double centerY = screenTopDiu + (screenHeightDiu - this.Height) / 2;

                this.Left = centerX;
                this.Top = centerY;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Positioned CustomizeFrameFormWPF at X={centerX}, Y={centerY} on screen '{mouseScreen.DeviceName}' (DPI scale: {dpiScale})");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error positioning form on mouse screen: {ex.Message}");
                this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Falling back to CenterScreen positioning");
            }
        }
        #endregion

        #region Current Frame Data Retrieval
        private dynamic GetCurrentFrameData(dynamic originalFrame)
        {
            try
            {
                string frameId = originalFrame.Id?.ToString();
                if (string.IsNullOrEmpty(frameId))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Original frame '{originalFrame.Title}' has no Id, using original reference");
                    return originalFrame;
                }

                var FrameData = Framemanager.GetFrameData();
                dynamic currentFrame = FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);

                if (currentFrame != null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Retrieved current frame data for '{currentFrame.Title}' (Id: {frameId})");
                    return currentFrame;
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Could not find current frame data for Id '{frameId}', using original reference");
                    return originalFrame;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error retrieving current frame data: {ex.Message}, using original reference");
                return originalFrame;
            }
        }
        #endregion

        #region Value Loading and Saving Methods - Helper Methods for Loading Values
        /// <summary>
        /// Shared logic to commit changes to JSON and update the runtime UI
        /// </summary>
        private void ApplyChanges()
        {
            ApplyRuntimeChanges(); // --- SPEED FIX: Update visuals instantly first ---
            SaveAllPropertiesToJson(); // Then save to JSON
            _result = true; // Mark as successful so if they close later, it counts as saved
        }

        /// <summary>
        /// Commits changes to ALL frames globally and updates their runtime UI
        /// </summary>
        private async void ApplyChangesToAll()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Applying customization changes to ALL frames globally.");

                // 1. Gather values from UI once
                string customColor = GetDropdownValue(_cmbCustomColor);
                string customLaunchEffect = GetDropdownValue(_cmbCustomLaunchEffect);
                string frameBorderColor = GetDropdownValue(_cmbframeBorderColor);
                string frameBorderThickness = _nudframeBorderThickness.Value.ToString();
                string titleTextColor = GetDropdownValue(_cmbTitleTextColor);
                string titleTextSize = GetDropdownValue(_cmbTitleTextSize);
                string boldTitleText = (_chkBoldTitleText.IsChecked ?? false).ToString().ToLower();
                string portalView = GetDropdownValue(_cmbPortalView); // --- NEW: Get Portal View for broadcast ---
                string iconSize = GetDropdownValue(_cmbIconSize);
                string iconSpacing = _nudIconSpacing.Value.ToString();
                string textColor = GetDropdownValue(_cmbTextColor);
                string disableTextShadow = (_chkDisableTextShadow.IsChecked ?? false).ToString().ToLower();
                string grayscaleIcons = (_chkGrayscaleIcons.IsChecked ?? false).ToString().ToLower();

                // 2. Iterate through all frames (Isolating the collection to prevent enumeration crashes)
                var allFramesRaw = Framemanager.GetFrameData();
                var allFrames = new System.Collections.Generic.List<dynamic>();
                foreach (var f in allFramesRaw) allFrames.Add(f);

                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();

                // TEST


                // --- PHASE 1: INSTANT VISUAL UPDATE ---
                // We update the screen first before touching the heavy JSON data engine
                foreach (dynamic targetFrame in allFrames)
                {
                    string frameId = targetFrame.Id?.ToString();
                    if (string.IsNullOrEmpty(frameId)) continue;

                    var win = windows.FirstOrDefault(w => w.Tag?.ToString() == frameId);
                    if (win != null)
                    {
                        var originalFrame = _frame;
                        _frame = targetFrame;

                        ApplyFrameBorderSettings(win);
                        ApplyTitleSettings(win);
                        ApplyCustomColorSetting(win);

                        string itemsType = targetFrame.ItemsType?.ToString();
                        if (itemsType == "Note")
                        {
                            ApplyNoteSettings(win);
                        }
                        else if (itemsType == "Portal")
                        {
                            // --- ULTRA FIX: Live switch Portal View without destroying the active manager ---
                            string targetView = GetDropdownValue(_cmbPortalView) ?? SettingsManager.DefaultPortalView;

                            // 1. Force update in-memory property immediately so internal syncs read the new state
                            Framemanager.UpdateFrameProperty(targetFrame, "ViewMode", targetView, "Global live view update");

                            // 2. Trigger live switch on the active window manager (Bulletproof against CS0165)
                            var pManagers = Framemanager.GetPortalFrames();
                            var targetKey = pManagers.Keys.FirstOrDefault(k => k.Id?.ToString() == frameId);

                            PortalFramemanager pManager = null; // Explicitly initialized to satisfy C# compiler
                            if (targetKey != null && pManagers.TryGetValue(targetKey, out pManager))
                            {
                                pManager?.SetViewMode(targetView);
                            }
                        }
                        else
                        {
                            ApplyIconSettings(win);
                        }

                        _frame = originalFrame;
                    }
                }

                // Force the UI to repaint the new visuals immediately
                await System.Threading.Tasks.Task.Delay(10); // A short delay forces the WPF render thread to flush

                // --- PHASE 2: BACKGROUND DATA SAVE ---
                // Offload the massive JSON disk writes to a background queue so the UI never freezes
                await System.Threading.Tasks.Task.Run(() =>
                {
                    foreach (dynamic targetFrame in allFrames)
                    {
                        string frameId = targetFrame.Id?.ToString();
                        if (string.IsNullOrEmpty(frameId)) continue;

                        // Safely execute the writes on the dispatcher, but chunked as Background Priority
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (targetFrame.ItemsType?.ToString() == "Portal")
                            {
                                Framemanager.UpdateFrameProperty(targetFrame, "ViewMode", portalView, "Global Apply: ViewMode updated");
                            }
                            Framemanager.UpdateFrameProperty(targetFrame, "CustomColor", customColor, "Global Apply: CustomColor updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "CustomLaunchEffect", customLaunchEffect, "Global Apply: CustomLaunchEffect updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "FrameBorderColor", frameBorderColor, "Global Apply: FrameBorderColor updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "FrameBorderThickness", frameBorderThickness, "Global Apply: FrameBorderThickness updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "TitleTextColor", titleTextColor, "Global Apply: TitleTextColor updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "TitleTextSize", titleTextSize, "Global Apply: TitleTextSize updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "BoldTitleText", boldTitleText, "Global Apply: BoldTitleText updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "IconSize", iconSize, "Global Apply: IconSize updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "IconSpacing", iconSpacing, "Global Apply: IconSpacing updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "TextColor", textColor, "Global Apply: TextColor updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "DisableTextShadow", disableTextShadow, "Global Apply: DisableTextShadow updated");
                            Framemanager.UpdateFrameProperty(targetFrame, "GrayscaleIcons", grayscaleIcons, "Global Apply: GrayscaleIcons updated");
                        }, System.Windows.Threading.DispatcherPriority.Background);
                    }
                });

                _result = true;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error applying changes to all frames: {ex.Message}");
                // Throw removed intentionally to avoid unhandled exception crashes in async void
            }
        }


        private void LoadCurrentValues()
        {
            try
            {
                // --- BULLETPROOF DYNAMIC EXTRACTOR ---
                string GetSafeProp(string propName)
                {
                    try
                    {
                        if (_frame is Newtonsoft.Json.Linq.JObject jObj) return jObj[propName]?.ToString();
                        if (_frame is System.Collections.Generic.IDictionary<string, object> dict && dict.ContainsKey(propName)) return dict[propName]?.ToString();
                        return null;
                    }
                    catch { return null; }
                }
                // -------------------------------------

                string frameTitle = GetSafeProp("Title") ?? "Unknown";
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Loading current values for frame '{frameTitle}'");

                string itemsType = GetSafeProp("ItemsType");
                bool isPortalFrame = itemsType == "Portal";
                bool isnoteFrame = itemsType == "Note";

                if (isPortalFrame)
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Portal frame detected - disabling icon controls for '{frameTitle}'");
                if (isnoteFrame)
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Note frame detected - disabling icon controls for '{frameTitle}'");

                // Load Frame Section properties
                LoadDropdownValue(_cmbCustomColor, GetSafeProp("CustomColor"), "CustomColor");
                LoadDropdownValue(_cmbCustomLaunchEffect, GetSafeProp("CustomLaunchEffect"), "CustomLaunchEffect");
                LoadDropdownValue(_cmbframeBorderColor, GetSafeProp("FrameBorderColor"), "FrameBorderColor");
                LoadNumericValue(_nudframeBorderThickness, GetSafeProp("FrameBorderThickness"), "FrameBorderThickness", 0);

                // Load Title Section properties
                LoadDropdownValue(_cmbTitleTextColor, GetSafeProp("TitleTextColor"), "TitleTextColor");
                LoadDropdownValue(_cmbTitleTextSize, GetSafeProp("TitleTextSize"), "TitleTextSize");
                LoadCheckboxValue(_chkBoldTitleText, GetSafeProp("BoldTitleText"), "BoldTitleText");

                // Load Icons Section properties
                if (isnoteFrame || isPortalFrame)
                {
                    DisableIconControls();
                    if (isPortalFrame)
                    {
                        _cmbPortalView.IsEnabled = true;
                        _cmbPortalView.Background = SystemColors.WindowBrush;
                        _cmbPortalView.ToolTip = null;
                        LoadDropdownValue(_cmbPortalView, GetSafeProp("ViewMode"), "ViewMode");
                    }
                }
                else
                {
                    EnableIconControls();
                    LoadDropdownValue(_cmbIconSize, GetSafeProp("IconSize"), "IconSize");
                    LoadNumericValue(_nudIconSpacing, GetSafeProp("IconSpacing"), "IconSpacing", 5);
                    LoadDropdownValue(_cmbTextColor, GetSafeProp("TextColor"), "TextColor");
                    LoadCheckboxValue(_chkDisableTextShadow, GetSafeProp("DisableTextShadow"), "DisableTextShadow");
                    LoadCheckboxValue(_chkGrayscaleIcons, GetSafeProp("GrayscaleIcons"), "GrayscaleIcons");
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Successfully loaded all current values for frame '{frameTitle}' (Portal: {isPortalFrame})");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading current values: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error loading frame properties: {ex.Message}", "Load Error");
            }
        }


        private void DisableIconControls()
        {
            try
            {
                _cmbIconSize.IsEnabled = false;
                _nudIconSpacing.IsEnabled = false;
                _cmbTextColor.IsEnabled = false;
                _chkDisableTextShadow.IsEnabled = false;
                _chkGrayscaleIcons.IsEnabled = false;
                _cmbCustomLaunchEffect.IsEnabled = false;

                _cmbIconSize.SelectedIndex = 0;
                _nudIconSpacing.Value = 5;
                _cmbTextColor.SelectedIndex = 0;
                _chkDisableTextShadow.IsChecked = false;
                _chkGrayscaleIcons.IsChecked = false;
                _cmbCustomLaunchEffect.SelectedIndex = 0;

                _cmbIconSize.Background = SystemColors.ControlBrush;
                _nudIconSpacing.Background = SystemColors.ControlBrush;
                _cmbTextColor.Background = SystemColors.ControlBrush;
                _cmbCustomLaunchEffect.Background = SystemColors.ControlBrush;

                string frameType = _frame.ItemsType?.ToString();
                string tooltipMessage = frameType == "Portal" ? "Icon appearance settings are not available for Portal Frames"
                                     : frameType == "Note" ? "Icon appearance settings are not available for Note Frames"
                                     : "Icon appearance settings are disabled";

                _cmbIconSize.ToolTip = tooltipMessage;
                _nudIconSpacing.ToolTip = tooltipMessage;
                _cmbTextColor.ToolTip = tooltipMessage;
                _chkDisableTextShadow.ToolTip = tooltipMessage;
                _chkGrayscaleIcons.ToolTip = tooltipMessage;
                _cmbCustomLaunchEffect.ToolTip = tooltipMessage;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Disabled icon controls for Portal Frame");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error disabling icon controls: {ex.Message}");
            }
        }

        private void EnableIconControls()
        {
            try
            {
                _cmbIconSize.IsEnabled = true;
                _nudIconSpacing.IsEnabled = true;
                _cmbTextColor.IsEnabled = true;
                _chkDisableTextShadow.IsEnabled = true;
                _chkGrayscaleIcons.IsEnabled = true;
                _cmbCustomLaunchEffect.IsEnabled = true;

                _cmbIconSize.Background = SystemColors.WindowBrush;
                _nudIconSpacing.Background = SystemColors.WindowBrush;
                _cmbTextColor.Background = SystemColors.WindowBrush;
                _cmbCustomLaunchEffect.Background = SystemColors.WindowBrush;

                _cmbIconSize.ToolTip = null;
                _nudIconSpacing.ToolTip = null;
                _cmbTextColor.ToolTip = null;
                _chkDisableTextShadow.ToolTip = null;
                _chkGrayscaleIcons.ToolTip = null;
                _cmbCustomLaunchEffect.ToolTip = null;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Enabled icon controls for regular frame");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error enabling icon controls: {ex.Message}");
            }
        }

        private void LoadDropdownValue(ComboBox comboBox, string currentValue, string propertyName)
        {
            try
            {
                if (string.IsNullOrEmpty(currentValue))
                {
                    comboBox.SelectedIndex = 0; // "Default" is always first item
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to Default (null/empty value)");
                }
                else
                {
                    for (int i = 0; i < comboBox.Items.Count; i++)
                    {
                        if (string.Equals((comboBox.Items[i] as ComboBoxItem)?.Tag as string, currentValue, StringComparison.OrdinalIgnoreCase))
                        {
                            comboBox.SelectedIndex = i;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to '{currentValue}'");
                            return;
                        }
                    }

                    comboBox.SelectedIndex = 0;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Unknown value '{currentValue}' for {propertyName}, defaulted to Default");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading {propertyName}: {ex.Message}");
                comboBox.SelectedIndex = 0;
            }
        }

        private void LoadNumericValue(NumericTextBox numericTextBox, string currentValue, string propertyName, int defaultValue)
        {
            try
            {
                if (string.IsNullOrEmpty(currentValue))
                {
                    numericTextBox.Value = defaultValue;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to default value {defaultValue} (null/empty value)");
                }
                else
                {
                    if (int.TryParse(currentValue, out int parsedValue))
                    {
                        if (parsedValue >= numericTextBox.Minimum && parsedValue <= numericTextBox.Maximum)
                        {
                            numericTextBox.Value = parsedValue;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to '{parsedValue}'");
                        }
                        else
                        {
                            numericTextBox.Value = defaultValue;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Value '{parsedValue}' out of bounds for {propertyName}, used default {defaultValue}");
                        }
                    }
                    else
                    {
                        numericTextBox.Value = defaultValue;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Failed to parse '{currentValue}' for {propertyName}, used default {defaultValue}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading {propertyName}: {ex.Message}");
                numericTextBox.Value = defaultValue;
            }
        }

        private void LoadCheckboxValue(CheckBox checkBox, string currentValue, string propertyName)
        {
            try
            {
                if (string.IsNullOrEmpty(currentValue))
                {
                    checkBox.IsChecked = false;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to false (null/empty value)");
                }
                else
                {
                    bool parsedValue = currentValue.Equals("true", StringComparison.OrdinalIgnoreCase);
                    checkBox.IsChecked = parsedValue;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to '{parsedValue}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading {propertyName}: {ex.Message}");
                checkBox.IsChecked = false;
            }
        }



        private void SaveAllPropertiesToJson()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Saving all properties to JSON for frame '{_frame.Title}'");

                // Get values from all 12 controls
                string customColor = GetDropdownValue(_cmbCustomColor);
                string customLaunchEffect = GetDropdownValue(_cmbCustomLaunchEffect);
                string frameBorderColor = GetDropdownValue(_cmbframeBorderColor);
                string frameBorderThickness = _nudframeBorderThickness.Value.ToString();

                string titleTextColor = GetDropdownValue(_cmbTitleTextColor);
                string titleTextSize = GetDropdownValue(_cmbTitleTextSize);
                string boldTitleText = (_chkBoldTitleText.IsChecked ?? false).ToString().ToLower();

                string portalView = GetDropdownValue(_cmbPortalView); // --- NEW: Get Portal View selection ---
                string iconSize = GetDropdownValue(_cmbIconSize);
                string iconSpacing = _nudIconSpacing.Value.ToString();
                string textColor = GetDropdownValue(_cmbTextColor);
                string disableTextShadow = (_chkDisableTextShadow.IsChecked ?? false).ToString().ToLower();
                string grayscaleIcons = (_chkGrayscaleIcons.IsChecked ?? false).ToString().ToLower();

                // Save all 12 properties using existing UpdateFrameProperty method
                Framemanager.UpdateFrameProperty(_frame, "CustomColor", customColor, $"CustomColor updated to '{customColor}'");
                Framemanager.UpdateFrameProperty(_frame, "CustomLaunchEffect", customLaunchEffect, $"CustomLaunchEffect updated to '{customLaunchEffect}'");
                Framemanager.UpdateFrameProperty(_frame, "FrameBorderColor", frameBorderColor, $"FrameBorderColor updated to '{frameBorderColor}'");
                Framemanager.UpdateFrameProperty(_frame, "FrameBorderThickness", frameBorderThickness, $"FrameBorderThickness updated to '{frameBorderThickness}'");
                Framemanager.UpdateFrameProperty(_frame, "TitleTextColor", titleTextColor, $"TitleTextColor updated to '{titleTextColor}'");
                Framemanager.UpdateFrameProperty(_frame, "TitleTextSize", titleTextSize, $"TitleTextSize updated to '{titleTextSize}'");
                Framemanager.UpdateFrameProperty(_frame, "BoldTitleText", boldTitleText, $"BoldTitleText updated to '{boldTitleText}'");
                if (_frame.ItemsType?.ToString() == "Portal")
                {
                    Framemanager.UpdateFrameProperty(_frame, "ViewMode", portalView, $"ViewMode updated to '{portalView}'");
                }
                Framemanager.UpdateFrameProperty(_frame, "IconSize", iconSize, $"IconSize updated to '{iconSize}'");
                Framemanager.UpdateFrameProperty(_frame, "IconSpacing", iconSpacing, $"IconSpacing updated to '{iconSpacing}'");
                Framemanager.UpdateFrameProperty(_frame, "TextColor", textColor, $"TextColor updated to '{textColor}'");
                Framemanager.UpdateFrameProperty(_frame, "DisableTextShadow", disableTextShadow, $"DisableTextShadow updated to '{disableTextShadow}'");
                Framemanager.UpdateFrameProperty(_frame, "GrayscaleIcons", grayscaleIcons, $"GrayscaleIcons updated to '{grayscaleIcons}'");

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"All 12 properties saved to JSON for frame '{_frame.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error saving properties to JSON: {ex.Message}");
                throw;
            }
        }

        private void ApplyRuntimeChanges()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Applying ALL runtime changes for frame '{_frame.Title}'");

                string frameId = _frame.Id?.ToString();
                if (string.IsNullOrEmpty(frameId)) return;

                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                var win = windows.FirstOrDefault(w => w.Tag?.ToString() == frameId);

                if (win != null)
                {
                    ApplyFrameBorderSettings(win);
                    ApplyTitleSettings(win);
                    ApplyCustomColorSetting(win);

                    // --- CHANGED LOGIC START ---
                    string itemsType = _frame.ItemsType?.ToString();

                    if (itemsType == "Note")
                    {
                        // Explicitly update Note visuals
                        ApplyNoteSettings(win);
                    }
                    else if (itemsType == "Portal")
                    {
                        // --- ULTRA FIX: Live switch Portal View without destroying the active manager ---
                        string targetView = GetDropdownValue(_cmbPortalView) ?? SettingsManager.DefaultPortalView;

                        var pManagers = Framemanager.GetPortalFrames();
                        var targetKey = pManagers.Keys.FirstOrDefault(k => k.Id?.ToString() == frameId);

                        PortalFramemanager pManager = null; // Explicitly initialized to satisfy C# compiler
                        if (targetKey != null && pManagers.TryGetValue(targetKey, out pManager))
                        {
                            pManager?.SetViewMode(targetView);
                        }
                    }
                    else
                    {
                        // Update Icon visuals for standard shortcut frames
                        ApplyIconSettings(win);
                    }
                    // --- CHANGED LOGIC END ---

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Applied all runtime changes to frame '{_frame.Title}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error applying runtime changes: {ex.Message}");
            }
        }


        // In CustomizeFrameFormManager.cs

        private void ApplyNoteSettings(NonActivatingWindow win)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Refreshing Note visuals for frame '{_frame.Title}'");

                var border = win.Content as Border;
                var dockPanel = border?.Child as DockPanel;

                if (dockPanel != null)
                {
                    var noteTextBox = dockPanel.Children.OfType<TextBox>().FirstOrDefault();

                    if (noteTextBox != null)
                    {
                        // --- ROBUST FIX START ---
                        // Create a temporary "Effective Frame" object using the CURRENT FORM VALUES.
                        // This ensures the visual update uses exactly what the user just clicked "Save" on,
                        // ignoring any stale data in the global list.

                        // 1. Clone properties from the original _frame into a dictionary
                        var effectiveFrame = new Dictionary<string, object>();
                        try
                        {
                            IDictionary<string, object> originalDict = null;
                            // Handle JObject vs ExpandoObject
                            if (_frame is Newtonsoft.Json.Linq.JObject jObj)
                                originalDict = jObj.ToObject<Dictionary<string, object>>();
                            else if (_frame is IDictionary<string, object> dict)
                                originalDict = dict;

                            if (originalDict != null)
                            {
                                foreach (var kvp in originalDict) effectiveFrame[kvp.Key] = kvp.Value;
                            }
                        }
                        catch { }

						// 2. OVERRIDE with values from the FORM CONTROLS
						effectiveFrame["CustomColor"] = GetDropdownValue(_cmbCustomColor);
						effectiveFrame["TextColor"] = GetDropdownValue(_cmbTextColor);
						// Map TitleTextSize dropdown to NoteFontSize property
						effectiveFrame["NoteFontSize"] = GetDropdownValue(_cmbTitleTextSize);

                        // 3. Convert to ExpandoObject for dynamic compatibility
                        dynamic dynamicFrame = new System.Dynamic.ExpandoObject();
                        var dynamicDict = (IDictionary<string, object>)dynamicFrame;
                        foreach (var kvp in effectiveFrame) dynamicDict[kvp.Key] = kvp.Value;
                        // --- ROBUST FIX END ---

                        // 4. Force NoteFramemanager to repaint using this fresh data
                        NoteFramemanager.RefreshNoteVisuals(dynamicFrame, noteTextBox);

                        // 5. Update font settings directly
                        try
                        {
                            string fontSizeStr = GetDropdownValue(_cmbTitleTextSize);
                            if (!string.IsNullOrEmpty(fontSizeStr))
                                noteTextBox.FontSize = NoteFramemanager.GetNoteFontSizeValue(fontSizeStr);
                        }
                        catch { }

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Successfully refreshed Note TextBox visuals using form values");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error applying note settings: {ex.Message}");
            }
        }

        private void ApplyFrameBorderSettings(NonActivatingWindow win)
        {
            try
            {
                string borderColor = GetDropdownValue(_cmbframeBorderColor);
                int borderThickness = _nudframeBorderThickness.Value;

                if (win.Content is System.Windows.Controls.Border frameBorder)
                {
                    if (!string.IsNullOrEmpty(borderColor) && borderColor != "Default")
                    {
                        var color = Utility.GetColorFromName(borderColor);
						frameBorder.BorderBrush = new System.Windows.Media.SolidColorBrush(color);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Applied border color '{borderColor}' to frame");
                    }
                    else if (borderThickness > 0)
                    {
						frameBorder.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Gray);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Applied default border color with thickness {borderThickness}");
                    }
                    else
                    {
						frameBorder.BorderBrush = null;
                    }

					frameBorder.BorderThickness = new System.Windows.Thickness(borderThickness);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Applied border thickness {borderThickness} to frame");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error applying frame border settings: {ex.Message}");
            }
        }

        private void ApplyTitleSettings(NonActivatingWindow win)
        {
            try
            {
                string titleColor = GetDropdownValue(_cmbTitleTextColor);
                string titleSize = GetDropdownValue(_cmbTitleTextSize);
                bool boldTitle = _chkBoldTitleText.IsChecked ?? false;

                var titleGrid = FindVisualChild<System.Windows.Controls.Grid>(win);
                if (titleGrid != null)
                {
                    var titleLabel = FindVisualChildInParent<System.Windows.Controls.Label>(titleGrid);
                    if (titleLabel != null)
                    {
                        if (!string.IsNullOrEmpty(titleColor) && titleColor != "Default")
                        {
                            var color = Utility.GetColorFromName(titleColor);
                            titleLabel.Foreground = new System.Windows.Media.SolidColorBrush(color);
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Applied title color '{titleColor}' to frame");
                        }
                        else
                        {
                            titleLabel.Foreground = System.Windows.Media.Brushes.White;
                        }

                        double fontSize = 12;
                        if (!string.IsNullOrEmpty(titleSize) && titleSize != "Default")
                        {
                            fontSize = titleSize switch
                            {
                                "Small" => 10,
                                "Medium" => 12,
                                "Large" => 14,
                                _ => 12
                            };
                        }
                        titleLabel.FontSize = fontSize;

                        titleLabel.FontWeight = boldTitle ? System.Windows.FontWeights.Bold : System.Windows.FontWeights.Normal;

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Applied title font size {fontSize} and bold={boldTitle} to frame");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Could not find title Label in titleGrid for frame '{_frame.Title}'");
                    }
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Could not find titleGrid for frame '{_frame.Title}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error applying title settings: {ex.Message}");
            }
        }

        private void ApplyCustomColorSetting(NonActivatingWindow win)
        {
            try
            {
                string customColor = GetDropdownValue(_cmbCustomColor);

                if (!string.IsNullOrEmpty(customColor) && customColor != "Default")
                {
                    Utility.ApplyTintAndColorToFrame(win, customColor);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Applied custom color '{customColor}' to frame");
                }
                else
                {
                    Utility.ApplyTintAndColorToFrame(win, null);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Applied default color to frame");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error applying custom color: {ex.Message}");
            }
        }

        private void ApplyIconSettings(NonActivatingWindow win)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Refreshing icons for frame '{_frame.Title}' to apply icon settings");

                var wrapPanel = FindVisualChild<System.Windows.Controls.WrapPanel>(win);
                if (wrapPanel != null)
                {
                    var FrameData = Framemanager.GetFrameData();
                    dynamic currentFrame = FrameData.FirstOrDefault(f => f.Id?.ToString() == _frame.Id?.ToString());

                    if (currentFrame != null)
                    {
                        string itemsType = currentFrame.ItemsType?.ToString();

                        if (itemsType == "Portal")
                        {
                            RefreshPortalFrameIcons(wrapPanel, currentFrame);
                        }
                        else
                        {
                            RefreshRegularFrameIcons(wrapPanel, currentFrame);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error applying icon settings: {ex.Message}");
            }
        }

        private void RefreshPortalFrameIcons(System.Windows.Controls.WrapPanel wrapPanel, dynamic portalFrame)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Refreshing Portal Frame icons for '{portalFrame.Title}'");

                wrapPanel.Children.Clear();

                var portalManagers = Framemanager.GetPortalFrames();
                if (portalManagers.ContainsKey(portalFrame))
                {
                    portalManagers[portalFrame].Dispose();
                    portalManagers[portalFrame] = new PortalFramemanager(portalFrame, wrapPanel);

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Successfully refreshed Portal Frame icons for '{portalFrame.Title}'");
                }
                else
                {
                    var newManager = new PortalFramemanager(portalFrame, wrapPanel);
                    portalManagers[portalFrame] = newManager;

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Created new PortalFramemanager for '{portalFrame.Title}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing Portal Frame icons: {ex.Message}");
            }
        }

        private void RefreshRegularFrameIcons(System.Windows.Controls.WrapPanel wrapPanel, dynamic regularFrame)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Refreshing regular frame icons for '{regularFrame.Title}'");

                wrapPanel.Children.Clear();

                bool tabsEnabled = regularFrame.TabsEnabled?.ToString().ToLower() == "true";
                Newtonsoft.Json.Linq.JArray items = null;

                if (tabsEnabled)
                {
                    var tabs = regularFrame.Tabs as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray();
                    int currentTab = Convert.ToInt32(regularFrame.CurrentTab?.ToString() ?? "0");

                    if (currentTab >= 0 && currentTab < tabs.Count)
                    {
                        var activeTab = tabs[currentTab] as Newtonsoft.Json.Linq.JObject;
                        if (activeTab != null)
                        {
                            items = activeTab["Items"] as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray();
                            string tabName = activeTab["TabName"]?.ToString() ?? $"Tab {currentTab}";
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                $"Refreshing icons from tab '{tabName}' for frame '{regularFrame.Title}'");
                        }
                    }
                }
                else
                {
                    items = regularFrame.Items as Newtonsoft.Json.Linq.JArray;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                        $"Refreshing icons from main Items for frame '{regularFrame.Title}'");
                }

                if (items != null)
                {
                    var sortedItems = items
                        .OfType<Newtonsoft.Json.Linq.JObject>()
                        .OrderBy(item => item["DisplayOrder"]?.Value<int>() ?? 0)
                        .ToList();

                    foreach (dynamic item in sortedItems)
                    {
						// FIX: Pass 'regularFrame' context here too!
						Framemanager.AddIcon(item, wrapPanel, regularFrame);

                        if (wrapPanel.Children.Count > 0)
                        {
                            var sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as System.Windows.Controls.StackPanel;
                            if (sp != null)
                            {
                                string filePath = item.Filename?.ToString() ?? "Unknown";
                                bool isFolder = item.IsFolder?.ToString().ToLower() == "true";
                                // FIX: Extract arguments
                                string arguments = Utility.GetShortcutArguments(filePath);

                                Framemanager.ClickEventAdder(sp, filePath, isFolder, arguments);

                                // FIX: Attach the centralized context menu
                                // We need to find the parent window (NonActivatingWindow) of the WrapPanel
                                var parentWindow = Window.GetWindow(wrapPanel) as NonActivatingWindow;
                                if (parentWindow != null)
                                {
                                    Framemanager.AttachIconContextMenu(sp, item, regularFrame, parentWindow);
                                }
                            }
                        }
                    }

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Successfully refreshed {sortedItems.Count} icons for frame '{regularFrame.Title}' {(tabsEnabled ? "(from active tab)" : "(from main Items)")}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"No items found for frame '{regularFrame.Title}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing regular frame icons: {ex.Message}");
            }
        }

        private T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            try
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
                {
                    var child = VisualTreeHelper.GetChild(parent, i);

                    if (child is T typedChild)
                    {
                        return typedChild;
                    }

                    var result = FindVisualChild<T>(child);
                    if (result != null)
                        return result;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error finding visual child: {ex.Message}");
            }

            return null;
        }

        private T FindVisualChildInParent<T>(DependencyObject parent) where T : DependencyObject
        {
            try
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
                {
                    var child = VisualTreeHelper.GetChild(parent, i);

                    if (child is T typedChild)
                    {
                        return typedChild;
                    }

                    var result = FindVisualChildInParent<T>(child);
                    if (result != null)
                        return result;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error finding visual child in parent: {ex.Message}");
            }

            return null;
        }

        private string GetDropdownValue(ComboBox comboBox)
        {
            try
            {
                string tag = (comboBox.SelectedItem as ComboBoxItem)?.Tag as string;
                if (comboBox.SelectedIndex <= 0 || tag == null || tag == "Default")
                {
                    return null; // Return null for "Default" to match existing JSON structure
                }
                return tag;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error getting dropdown value: {ex.Message}");
                return null;
            }
        }
        #endregion
    }

    #region NumericTextBox Custom Control
    /// <summary>
    /// Custom WPF NumericTextBox control that emulates Windows Forms NumericUpDown
    /// with TextBox + Up/Down buttons (▲▼) and min/max validation
    /// </summary>
    public class NumericTextBox : UserControl
    {
        #region Private Fields
        private TextBox _textBox;
        private Button _upButton;
        private Button _downButton;
        private int _value = 0;
        private int _minimum = 0;
        private int _maximum = 100;
        #endregion

        #region Public Properties
        public int Value
        {
            get => _value;
            set
            {
                int newValue = Math.Max(_minimum, Math.Min(_maximum, value));
                if (_value != newValue)
                {
                    _value = newValue;
                    if (_textBox != null)
                        _textBox.Text = _value.ToString();
                    ValueChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        public int Minimum
        {
            get => _minimum;
            set
            {
                _minimum = value;
                if (_value < _minimum)
                    Value = _minimum;
            }
        }

        public int Maximum
        {
            get => _maximum;
            set
            {
                _maximum = value;
                if (_value > _maximum)
                    Value = _maximum;
            }
        }

        public event EventHandler ValueChanged;
        #endregion

        #region Constructor
        public NumericTextBox()
        {
            InitializeComponent();
        }
        #endregion

        #region Initialization
        private void InitializeComponent()
        {
            // Main grid layout: TextBox on left, buttons on right
            Grid mainGrid = new Grid();
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });

            // TextBox for numeric input
            _textBox = new TextBox
            {
                Text = _value.ToString(),
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderBrush = new SolidColorBrush(Color.FromRgb(171, 173, 179)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(4, 2, 4, 2)
            };
            _textBox.TextChanged += TextBox_TextChanged;
            _textBox.KeyDown += TextBox_KeyDown;
            Grid.SetColumn(_textBox, 0);

            // --- THE WPF WAY: Dynamic Grid Layout ---
            // Replaces the StackPanel with a Grid that evenly splits available height 50/50.
            // This means no hardcoded heights! It will perfectly fit 20px on the Options menu, 
            // and seamlessly scale up to 24px+ on the Customize menu.
            Grid buttonPanel = new Grid();
            buttonPanel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            buttonPanel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            // Up button (▲)
            _upButton = new Button
            {
                Content = "▲",
                Width = 18,
                VerticalAlignment = VerticalAlignment.Stretch, // Fills its 50% half automatically
                FontSize = 6,
                FontFamily = new FontFamily("Segoe UI"),
                Padding = new Thickness(0),
                Margin = new Thickness(1, 0, 0, 0),
                Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(171, 173, 179)),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand
            };
            _upButton.Click += UpButton_Click;
            Grid.SetRow(_upButton, 0);

            // Down button (▼)
            _downButton = new Button
            {
                Content = "▼",
                Width = 18,
                VerticalAlignment = VerticalAlignment.Stretch, // Fills its 50% half automatically
                FontSize = 6,
                FontFamily = new FontFamily("Segoe UI"),
                Padding = new Thickness(0),
                Margin = new Thickness(1, 0, 0, 0),
                Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(171, 173, 179)),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand
            };
            _downButton.Click += DownButton_Click;
            Grid.SetRow(_downButton, 1);

            buttonPanel.Children.Add(_upButton);
            buttonPanel.Children.Add(_downButton);
            Grid.SetColumn(buttonPanel, 1);

            mainGrid.Children.Add(_textBox);
            mainGrid.Children.Add(buttonPanel);

            this.Content = mainGrid;
            // --- UI FIX: Removed "this.Height = 24" so the control dynamically respects the 20px limit passed from the parent ---
        }
        #endregion

        #region Event Handlers
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(_textBox.Text, out int newValue))
            {
                int validValue = Math.Max(_minimum, Math.Min(_maximum, newValue));
                if (_value != validValue)
                {
                    _value = validValue;
                    // Only update textbox if the valid value is different from what user typed
                    if (validValue.ToString() != _textBox.Text)
                    {
                        _textBox.Text = validValue.ToString();
                        _textBox.SelectionStart = _textBox.Text.Length; // Move cursor to end
                    }
                    ValueChanged?.Invoke(this, EventArgs.Empty);
                }
            }
            else if (string.IsNullOrEmpty(_textBox.Text))
            {
                // Allow empty text temporarily for user input
                return;
            }
            else
            {
                // Invalid input, restore previous value
                _textBox.Text = _value.ToString();
                _textBox.SelectionStart = _textBox.Text.Length;
            }
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            // Allow navigation keys
            if (e.Key == Key.Left || e.Key == Key.Right || e.Key == Key.Home || e.Key == Key.End ||
                e.Key == Key.Tab || e.Key == Key.Delete || e.Key == Key.Back)
                return;

            // Allow numbers
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
                return;

            // Allow up/down arrow keys to change value
            if (e.Key == Key.Up)
            {
                Value = Math.Min(_maximum, _value + 1);
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Down)
            {
                Value = Math.Max(_minimum, _value - 1);
                e.Handled = true;
                return;
            }

            // Block all other keys
            e.Handled = true;
        }

        private void UpButton_Click(object sender, RoutedEventArgs e)
        {
            Value = Math.Min(_maximum, _value + 1);
            _textBox.Focus(); // Keep focus on textbox for keyboard input
        }

        private void DownButton_Click(object sender, RoutedEventArgs e)
        {
            Value = Math.Max(_minimum, _value - 1);
            _textBox.Focus(); // Keep focus on textbox for keyboard input
        }

        protected override void OnGotFocus(RoutedEventArgs e)
        {
            base.OnGotFocus(e);
            _textBox?.Focus();
            _textBox?.SelectAll();
        }
        #endregion
    }
    #endregion
}

#endregion